/**
 * Grøenk — Inventory Transactions weekly archive
 *
 * ADDED 2026-09-26 (requested by Emese). The ledger grows by ~250 rows a day
 * and is the slowest thing the app loads. Every Monday this keeps the current
 * week + the last 2 completed weeks as individual rows. Everything older only
 * matters through the stock it adds up to, so per product + location it is
 * collapsed into ONE "Opening balance" row with exactly the same effect, using
 * the exact ledger rules of the app / PK order / daily consumption scripts.
 *
 * How (per product + location):
 *   - WITH a physical count: rows created before the latest count and dated
 *     before the cutoff are collapsed into one of those same rows (it keeps its
 *     own time stamp, so it is still "before the count").
 *   - With NO count: rows dated before the cutoff are collapsed the same way.
 *   - Count rows that are the latest count are NEVER changed, nor is anything
 *     dated on/after the cutoff. So "Updated X days ago", NO INVENTORY, the
 *     Friday audit (last 7 days) and the Monday reports (last week) are untouched.
 *
 * Safety (never relaxed):
 *   1. Stock for every product+location computed before and on the planned
 *      result; any difference > 0.000001 => abort, nothing written.
 *   2. Every deleted row and the original + new values of every changed row go
 *      to a CSV in the controlling Drive folder, downloaded back and verified
 *      (bytes, checksum, every record ID) BEFORE any write.
 *   3. After the run the whole ledger is re-read and every stock recomputed;
 *      the result is emailed ("done ✅" or "check needed").
 *   4. Refuses to write between 01:00 and 08:00 UTC (daily consumption window).
 *   DRY_RUN=true: steps 1-2 + email only.
 *
 * Env: AIRTABLE_TOKEN, RESEND_API_KEY, APPS_SCRIPT_URL, APPS_SCRIPT_SECRET, DRY_RUN
 */

const crypto = require('crypto');

const BASE_ID = 'appPcdy4HEJuDOF4j';
const TABLE = 'Inventory Transactions';
const RETENTION_WEEKS = 2;   // completed weeks kept as individual rows (+ current week)
const TIME_ZONE = 'Europe/Madrid';
const todayMadrid = () => new Intl.DateTimeFormat('en-CA', { timeZone: TIME_ZONE, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
const addDaysStr = (d, n) => { const x = new Date(d + 'T00:00:00Z'); x.setUTCDate(x.getUTCDate() + n); return x.toISOString().slice(0, 10); };
const mondayOf = d => addDaysStr(d, -((new Date(d + 'T00:00:00Z').getUTCDay() + 6) % 7));
const TODAY = todayMadrid();
const CUTOFF = process.env.CUTOFF_OVERRIDE || addDaysStr(mondayOf(TODAY), -7 * RETENTION_WEEKS); // first date KEPT
const ARCHIVE_NAME = `inventory-transactions_archive_${TODAY}_before_${CUTOFF}.csv`;
const EMAIL_FROM = 'emese@groenk.com';
const REPORT_EMAILS = ['emese@groenk.com', 'controlling@groenk.com'];
const EPS = 1e-6;

const { AIRTABLE_TOKEN, RESEND_API_KEY, APPS_SCRIPT_URL, APPS_SCRIPT_SECRET } = process.env;
const DRY_RUN = String(process.env.DRY_RUN || 'false').toLowerCase() === 'true';
if (!AIRTABLE_TOKEN || !RESEND_API_KEY || !APPS_SCRIPT_URL || !APPS_SCRIPT_SECRET) {
  console.error('Missing env: AIRTABLE_TOKEN, RESEND_API_KEY, APPS_SCRIPT_URL, APPS_SCRIPT_SECRET');
  process.exit(1);
}
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ---------- Ledger rules: identical to index.html computeCurrentStock (unclamped) ----------
const LEDGER_EFFECT = (type, qty) => (type === 'Waste' || type === 'Consumption') ? -Math.abs(qty) : qty;
const eff = t => LEDGER_EFFECT(t.fields['Type'], Number(t.fields['Quantity']) || 0);
const IS_MANUAL_COUNT = t => t.fields['Type'] === 'Manual Adjustment' && String(t.fields['Notes'] || '').toLowerCase().includes('manual count');
function stockOf(rows) {
  let lastCount = null;
  for (const t of rows) if (IS_MANUAL_COUNT(t) && (!lastCount || t.createdTime > lastCount.createdTime)) lastCount = t;
  if (!lastCount) return rows.reduce((s, t) => s + eff(t), 0);
  const baseline = rows.filter(t => t.createdTime < lastCount.createdTime).reduce((s, t) => s + eff(t), 0) + eff(lastCount);
  return rows.filter(t => t.createdTime > lastCount.createdTime && (t.fields['Date'] || '') >= (lastCount.fields['Date'] || ''))
    .reduce((s, t) => s + eff(t), baseline);
}

// ---------- Airtable ----------
async function at(url, options = {}) {
  let lastErr;
  for (let i = 0; i < 6; i++) {
    try {
      const res = await fetch(url, { ...options, headers: { Authorization: `Bearer ${AIRTABLE_TOKEN}`, 'Content-Type': 'application/json' } });
      const text = await res.text();
      if (res.ok) return JSON.parse(text);
      lastErr = new Error(`Airtable HTTP ${res.status}: ${text.slice(0, 300)}`);
      if (res.status !== 429 && res.status < 500) throw lastErr;
    } catch (e) { lastErr = e; if (/HTTP 4(?!29)/.test(e.message)) throw e; }
    await sleep(1000 * 2 ** i);
  }
  throw lastErr;
}
const tableUrl = () => `https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(TABLE)}`;
async function getAll() {
  let out = [], offset;
  do {
    const u = new URL(tableUrl()); u.searchParams.set('pageSize', '100');
    if (offset) u.searchParams.set('offset', offset);
    const d = await at(u.toString()); out = out.concat(d.records); offset = d.offset; await sleep(220);
  } while (offset);
  return out;
}
async function patchMany(updates) {
  for (let i = 0; i < updates.length; i += 10) { await at(tableUrl(), { method: 'PATCH', body: JSON.stringify({ records: updates.slice(i, i + 10) }) }); await sleep(220); }
}
async function createMany(fieldsList) {
  const created = [];
  for (let i = 0; i < fieldsList.length; i += 10) {
    const d = await at(tableUrl(), { method: 'POST', body: JSON.stringify({ records: fieldsList.slice(i, i + 10).map(fields => ({ fields })) }) });
    created.push(...d.records); await sleep(220);
  }
  return created;
}
async function deleteMany(ids) {
  for (let i = 0; i < ids.length; i += 10) {
    const u = new URL(tableUrl()); ids.slice(i, i + 10).forEach(id => u.searchParams.append('records[]', id));
    await at(u.toString(), { method: 'DELETE' }); await sleep(220);
  }
}

// ---------- Drive via the existing Apps Script ----------
async function appsScript(payload) {
  const res = await fetch(APPS_SCRIPT_URL, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ ...payload, secret: APPS_SCRIPT_SECRET }), redirect: 'follow' });
  const text = await res.text(); let d;
  try { d = JSON.parse(text); } catch (_) { throw new Error(`Apps Script non-JSON (HTTP ${res.status}): ${text.slice(0, 200)}`); }
  if (!d.ok) throw new Error(`Apps Script: ${d.error}`); return d;
}

// ---------- CSV ----------
const COLS = ['Action', 'Airtable Record ID', 'Created', 'Date', 'Type', 'Quantity', 'New Quantity', 'Related Product', 'Location', 'Notes', 'Transaction Added By', 'Waste Reason', 'Stock Status After Transaction'];
const cell = v => { const s = v == null ? '' : String(v); return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
const toCsv = rows => '\uFEFF' + [COLS.join(','), ...rows.map(r => COLS.map(c => cell(r[c])).join(','))].join('\r\n') + '\r\n';
function parseIds(text) { // record IDs are simple tokens; count them per line start
  return new Set((text.match(/\b(rec[A-Za-z0-9]{14})\b/g) || []));
}
async function sendEmail(to, subject, text, attachments) {
  const res = await fetch('https://api.resend.com/emails', { method: 'POST', headers: { Authorization: `Bearer ${RESEND_API_KEY}`, 'Content-Type': 'application/json' }, body: JSON.stringify({ from: EMAIL_FROM, to, subject, text, attachments }) });
  const r = await res.json(); if (!res.ok || r.error) throw new Error('Resend: ' + JSON.stringify(r).slice(0, 300));
}

function buildPlan(txns, cutoff = CUTOFF) {
  const groups = {}; const untouchable = [];
  for (const t of txns) {
    const p = t.fields['Related Product'] || [], l = t.fields['Location'] || [];
    if (p.length !== 1 || l.length !== 1) { untouchable.push(t); continue; }
    (groups[`${p[0]}|${l[0]}`] = groups[`${p[0]}|${l[0]}`] || []).push(t);
  }
  const before = {}, after = {}, plan = { deletes: [], updates: [] };
  for (const [key, rows] of Object.entries(groups)) {
    before[key] = stockOf(rows);
    let lastCount = null;
    for (const t of rows) if (IS_MANUAL_COUNT(t) && (!lastCount || t.createdTime > lastCount.createdTime)) lastCount = t;
    const [productId, locationId] = key.split('|');
    // Rows that only matter through their sum: dated before the cutoff and (if
    // there is a count) created before the latest count. The latest count
    // itself is never part of it.
    const D = rows.filter(t => (t.fields['Date'] || '') < cutoff && (!lastCount || (t.id !== lastCount.id && t.createdTime < lastCount.createdTime)));
    if (D.length < 2) { after[key] = before[key]; continue; }   // nothing to gain
    const total = D.reduce((s, t) => s + eff(t), 0);
    // Keep the most recently created of them as the holder of the total.
    const holder = D.reduce((a, b) => (b.createdTime > a.createdTime ? b : a));
    const newFields = { 'Type': 'Manual Adjustment', 'Quantity': Math.round(total * 1e6) / 1e6, 'Notes': `Opening balance — archived history before ${cutoff}`, 'Waste Reason': null };
    const changedHolder = { ...holder, fields: { ...holder.fields, ...newFields } };
    const simulated = rows.filter(t => !D.includes(t)).concat([changedHolder]);
    after[key] = stockOf(simulated);
    plan.updates.push({ id: holder.id, fields: newFields, original: holder, productId, locationId });
    D.filter(t => t.id !== holder.id).forEach(t => plan.deletes.push({ t, productId, locationId }));
  }
  return { groups, untouchable, before, after, plan };
}

const RUN_STARTED_AT = new Date().toISOString();
async function main() {
  const hour = new Date().getUTCHours();
  if (!DRY_RUN && hour >= 1 && hour < 8) throw new Error('Refusing a real run between 01:00 and 08:00 UTC (daily consumption job window).');
  console.log(`${DRY_RUN ? 'DRY RUN' : 'REAL RUN'} — cutoff ${CUTOFF}`);

  const [txns, products, locations] = [await getAll(), null, null];
  const prodRes = await (async () => { let o = [], off; do { const u = new URL(`https://api.airtable.com/v0/${BASE_ID}/Products`); u.searchParams.set('pageSize', '100'); u.searchParams.append('fields[]', 'Name'); if (off) u.searchParams.set('offset', off); const d = await at(u.toString()); o = o.concat(d.records); off = d.offset; await sleep(220); } while (off); return o; })();
  const locRes = (await at(`https://api.airtable.com/v0/${BASE_ID}/Locations?fields[]=Name`)).records;
  const pName = Object.fromEntries(prodRes.map(p => [p.id, p.fields['Name']]));
  const lName = Object.fromEntries(locRes.map(l => [l.id, l.fields['Name']]));
  console.log(`Loaded ${txns.length} transactions.`);

  const { groups, untouchable, before, after, plan } = buildPlan(txns);
  if (untouchable.length) console.log(`${untouchable.length} rows without exactly one product+location — left untouched.`);

  // Safety 1: identical stock everywhere.
  const mismatches = Object.keys(before).filter(k => Math.abs(before[k] - after[k]) > EPS);
  if (mismatches.length) throw new Error(`Plan changes stock for ${mismatches.length} product/location pairs — aborted, nothing written:\n` + mismatches.slice(0, 30).map(k => `${pName[k.split('|')[0]]} @ ${lName[k.split('|')[1]]}: ${before[k]} -> ${after[k]}`).join('\n'));
  console.log(`Plan OK: stock identical for all ${Object.keys(before).length} product/location pairs. Delete ${plan.deletes.length}, ${plan.updates.length} opening-balance rows. Table: ${txns.length} -> ${txns.length - plan.deletes.length} rows.`);

  if (!plan.deletes.length) { console.log('Nothing to archive this week.'); return; }

  // Safety 2: archive + verify in Drive.
  const row = (action, t, extra = {}) => ({
    'Action': action, 'Airtable Record ID': t.id, 'Created': t.createdTime, 'Date': t.fields['Date'], 'Type': t.fields['Type'], 'Quantity': t.fields['Quantity'],
    'Related Product': (t.fields['Related Product'] || []).map(id => pName[id] || id).join(' | '), 'Location': (t.fields['Location'] || []).map(id => lName[id] || id).join(' | '),
    'Notes': t.fields['Notes'], 'Transaction Added By': t.fields['Transaction Added By'], 'Waste Reason': t.fields['Waste Reason'], 'Stock Status After Transaction': t.fields['Stock Status After Transaction'], ...extra,
  });
  const archiveRows = [...plan.deletes.map(d => row('DELETED', d.t)), ...plan.updates.map(u => row('CHANGED TO OPENING BALANCE', u.original, { 'New Quantity': u.fields.Quantity }))];
  const csv = toCsv(archiveRows); const bytes = Buffer.from(csv, 'utf8'); const md5 = crypto.createHash('md5').update(bytes).digest('hex');
  const existing = (await appsScript({ action: 'find', name: ARCHIVE_NAME })).file;
  const up = (await appsScript({ action: 'upload', name: ARCHIVE_NAME, existingId: existing ? existing.id : null, contentBase64: bytes.toString('base64') })).file;
  const back = Buffer.from((await appsScript({ action: 'download', id: up.id })).contentBase64, 'base64');
  const ids = parseIds(back.toString('utf8'));
  const missing = archiveRows.filter(r => !ids.has(r['Airtable Record ID']));
  if (!back.equals(bytes) || up.md5Checksum !== md5 || missing.length) throw new Error(`Drive archive verification FAILED (identical=${back.equals(bytes)}, md5 ${up.md5Checksum} vs ${md5}, missing ${missing.length}) — nothing written.`);
  console.log(`Archive ${ARCHIVE_NAME}: ${archiveRows.length} rows saved and verified in Drive.`);

  const summary = `Inventory Transactions weekly archive — ${DRY_RUN ? 'DRY RUN (nothing changed)' : 'REAL RUN'}\n\n` +
    `Kept as individual rows: everything dated ${CUTOFF} or later\nRows in table: ${txns.length}\nDeleted (older history): ${plan.deletes.length}\nOpening-balance rows (one per product+location): ${plan.updates.length}\nRows after: ${txns.length - plan.deletes.length}\n\n` +
    `Stock check: identical for all ${Object.keys(before).length} product/location pairs.\nArchive in Drive (verified): ${ARCHIVE_NAME}\n`;
  if (DRY_RUN) { await sendEmail(REPORT_EMAILS, `Inventory archive — DRY RUN plan (${TODAY})`, summary); console.log(summary); return; }

  // Real run: fold/create first, then delete.
  await patchMany(plan.updates.map(u => ({ id: u.id, fields: u.fields })));
  await deleteMany(plan.deletes.map(d => d.t.id));

  // Safety 4: re-read and verify.
  const fresh = await getAll();
  const g2 = {};
  for (const t of fresh) { const p = t.fields['Related Product'] || [], l = t.fields['Location'] || []; if (p.length === 1 && l.length === 1) (g2[`${p[0]}|${l[0]}`] = g2[`${p[0]}|${l[0]}`] || []).push(t); }
  // Rows created by the app or automations during the run are the only
  // legitimate reason for a difference; the report shows them so it's obvious.
  const runStart = RUN_STARTED_AT;
  const bad = Object.keys(before).filter(k => {
    const now = g2[k] ? stockOf(g2[k]) : 0;
    if (Math.abs(now - before[k]) <= EPS) return false;
    const concurrent = (g2[k] || []).filter(t => t.createdTime >= runStart && !String(t.fields['Notes'] || '').startsWith('Opening balance'));
    const expected = before[k] + concurrent.reduce((s2, t) => s2 + eff(t), 0);
    return Math.abs(now - expected) > EPS;
  });
  const report = summary + `\nAfter-run check: ${bad.length ? `⚠️ ${bad.length} pair(s) differ:\n` + bad.slice(0, 50).map(k => `${pName[k.split('|')[0]]} @ ${lName[k.split('|')[1]]}: expected ${before[k]}, now ${g2[k] ? stockOf(g2[k]) : 0}`).join('\n') : 'all stock levels identical ✅'}\nTable now: ${fresh.length} rows.`;
  await sendEmail(REPORT_EMAILS, bad.length ? '⚠️ Inventory archive — check needed' : `Inventory archive — done ✅ (${TODAY})`, report);
  console.log(report);
  if (bad.length) process.exitCode = 1;
}

if (require.main === module) main().catch(async e => {
  console.error(e);
  try { await sendEmail(REPORT_EMAILS, '⚠️ Inventory archive stopped — nothing or only part changed', String(e && e.stack || e)); } catch (_) {}
  process.exit(1);
});

module.exports = { buildPlan, stockOf };
