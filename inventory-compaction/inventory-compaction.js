/**
 * Grøenk — Inventory Transactions compaction (one-off, run by hand)
 *
 * ADDED 2026-09-23 (requested by Emese). The Inventory Transactions ledger grows by
 * hundreds of rows a day and is the slowest thing the app loads. Everything dated
 * before the Monday 2026-09-21 physical inventory is history: it only matters
 * through the stock figure it adds up to. This script replaces that history with
 * the smallest set of rows that gives EXACTLY the same stock for every
 * product + location, using the exact ledger rules of the app / PK order /
 * daily consumption scripts:
 *
 *   Product+location WITH a physical count (the normal case):
 *     rows created before the latest count and dated before 2026-09-21 are
 *     deleted, and their total effect is folded into that count row's Quantity
 *     (its date, creator and time stamp stay the same, so "Updated X days ago"
 *     and NO INVENTORY are unaffected).
 *   Product+location with NO count ever:
 *     rows dated before 2026-09-21 are deleted and replaced by ONE
 *     "Opening balance" row with the same total (not a count, so it still shows
 *     NO INVENTORY).
 *
 * Safety (never relaxed):
 *   1. Stock for every product+location is computed before, and on the planned
 *      result in memory. Any difference > 0.000001 => abort, nothing written.
 *   2. Every row to be deleted, and the original value of every row to be
 *      changed, goes to a CSV in the controlling Drive folder. It is downloaded
 *      back and verified (row count, every record ID, checksum) BEFORE any write.
 *   3. DRY_RUN=true (default for manual runs) stops after step 2 and emails the
 *      plan. Nothing in Airtable is touched.
 *   4. After a real run the whole ledger is re-read and every stock recomputed;
 *      any mismatch is emailed immediately with the details needed to repair it.
 *   Rows with several products or locations, or none, are never touched.
 *
 * Run it OUTSIDE the daily consumption window (01:00-08:00 UTC) and after the
 * Friday inventory audit, so that audit never sees the folded count rows.
 *
 * Env: AIRTABLE_TOKEN, RESEND_API_KEY, APPS_SCRIPT_URL, APPS_SCRIPT_SECRET, DRY_RUN
 */

const crypto = require('crypto');

const BASE_ID = 'appPcdy4HEJuDOF4j';
const TABLE = 'Inventory Transactions';
const CUTOFF = '2026-09-21';                 // first date KEPT as individual rows
const OPENING_DATE = '2026-09-20';
const ARCHIVE_NAME = `inventory-transactions_archive_before_${CUTOFF}.csv`;
const EMAIL_FROM = 'emese@groenk.com';
const REPORT_EMAILS = ['emese@groenk.com', 'controlling@groenk.com'];
const EPS = 1e-6;

const { AIRTABLE_TOKEN, RESEND_API_KEY, APPS_SCRIPT_URL, APPS_SCRIPT_SECRET } = process.env;
const DRY_RUN = String(process.env.DRY_RUN || 'true').toLowerCase() !== 'false';
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

function buildPlan(txns) {
  // Group by product+location; rows that aren't exactly 1 product + 1 location are never touched.
  const groups = {}; const untouchable = [];
  for (const t of txns) {
    const p = t.fields['Related Product'] || [], l = t.fields['Location'] || [];
    if (p.length !== 1 || l.length !== 1) { untouchable.push(t); continue; }
    (groups[`${p[0]}|${l[0]}`] = groups[`${p[0]}|${l[0]}`] || []).push(t);
  }

  const before = {}, after = {}, plan = { deletes: [], updates: [], creates: [], log: [] };
  for (const [key, rows] of Object.entries(groups)) {
    before[key] = stockOf(rows);
    let lastCount = null;
    for (const t of rows) if (IS_MANUAL_COUNT(t) && (!lastCount || t.createdTime > lastCount.createdTime)) lastCount = t;
    const [productId, locationId] = key.split('|');
    let D, simulated;
    if (lastCount) {
      D = rows.filter(t => t.id !== lastCount.id && t.createdTime < lastCount.createdTime && (t.fields['Date'] || '') < CUTOFF);
      if (!D.length) { after[key] = before[key]; continue; }
      const fold = D.reduce((s, t) => s + eff(t), 0);
      // The count row is a Manual Adjustment, so its effect equals its Quantity.
      const newQty = (Number(lastCount.fields['Quantity']) || 0) + fold;
      const newNotes = `${lastCount.fields['Notes'] || 'Manual count'} | incl. opening balance of history before ${CUTOFF} (was ${lastCount.fields['Quantity']})`;
      const changedCount = { ...lastCount, fields: { ...lastCount.fields, Quantity: newQty, Notes: newNotes } };
      simulated = rows.filter(t => !D.includes(t) && t.id !== lastCount.id).concat([changedCount]);
      plan.updates.push({ id: lastCount.id, fields: { Quantity: newQty, Notes: newNotes }, original: lastCount, productId, locationId });
    } else {
      D = rows.filter(t => (t.fields['Date'] || '') < CUTOFF);
      if (!D.length) { after[key] = before[key]; continue; }
      const total = D.reduce((s, t) => s + eff(t), 0);
      const kept = rows.filter(t => !D.includes(t));
      simulated = kept.slice();
      if (Math.abs(total) > EPS) {
        const fields = { 'Date': OPENING_DATE, 'Type': 'Manual Adjustment', 'Quantity': total, 'Related Product': [productId], 'Location': [locationId], 'Notes': `Opening balance — archived history before ${CUTOFF}`, 'Transaction Added By': 'Archive' };
        simulated.push({ id: 'new', createdTime: '9999-12-31T00:00:00.000Z', fields });
        plan.creates.push(fields);
      }
    }
    after[key] = stockOf(simulated);
    D.forEach(t => plan.deletes.push({ t, productId, locationId }));
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
  console.log(`Plan OK: stock identical for all ${Object.keys(before).length} product/location pairs. Delete ${plan.deletes.length}, fold into ${plan.updates.length} count rows, create ${plan.creates.length} opening rows. Table: ${txns.length} -> ${txns.length - plan.deletes.length + plan.creates.length} rows.`);

  // Safety 2: archive + verify in Drive.
  const row = (action, t, extra = {}) => ({
    'Action': action, 'Airtable Record ID': t.id, 'Created': t.createdTime, 'Date': t.fields['Date'], 'Type': t.fields['Type'], 'Quantity': t.fields['Quantity'],
    'Related Product': (t.fields['Related Product'] || []).map(id => pName[id] || id).join(' | '), 'Location': (t.fields['Location'] || []).map(id => lName[id] || id).join(' | '),
    'Notes': t.fields['Notes'], 'Transaction Added By': t.fields['Transaction Added By'], 'Waste Reason': t.fields['Waste Reason'], 'Stock Status After Transaction': t.fields['Stock Status After Transaction'], ...extra,
  });
  const archiveRows = [...plan.deletes.map(d => row('DELETED', d.t)), ...plan.updates.map(u => row('COUNT QTY CHANGED', u.original, { 'New Quantity': u.fields.Quantity }))];
  const csv = toCsv(archiveRows); const bytes = Buffer.from(csv, 'utf8'); const md5 = crypto.createHash('md5').update(bytes).digest('hex');
  const existing = (await appsScript({ action: 'find', name: ARCHIVE_NAME })).file;
  const up = (await appsScript({ action: 'upload', name: ARCHIVE_NAME, existingId: existing ? existing.id : null, contentBase64: bytes.toString('base64') })).file;
  const back = Buffer.from((await appsScript({ action: 'download', id: up.id })).contentBase64, 'base64');
  const ids = parseIds(back.toString('utf8'));
  const missing = archiveRows.filter(r => !ids.has(r['Airtable Record ID']));
  if (!back.equals(bytes) || up.md5Checksum !== md5 || missing.length) throw new Error(`Drive archive verification FAILED (identical=${back.equals(bytes)}, md5 ${up.md5Checksum} vs ${md5}, missing ${missing.length}) — nothing written.`);
  console.log(`Archive ${ARCHIVE_NAME}: ${archiveRows.length} rows saved and verified in Drive.`);

  const summary = `Inventory Transactions compaction — ${DRY_RUN ? 'DRY RUN (nothing changed)' : 'REAL RUN'}\n\n` +
    `Rows in table: ${txns.length}\nTo delete (history before ${CUTOFF}): ${plan.deletes.length}\nCount rows with history folded in: ${plan.updates.length}\nNew opening-balance rows: ${plan.creates.length}\nRows after: ${txns.length - plan.deletes.length + plan.creates.length}\n\n` +
    `Stock check: identical for all ${Object.keys(before).length} product/location pairs.\nArchive in Drive (verified): ${ARCHIVE_NAME}\n`;
  if (DRY_RUN) { await sendEmail(REPORT_EMAILS, 'Inventory compaction — DRY RUN plan', summary); console.log(summary); return; }

  // Real run: fold/create first, then delete.
  await patchMany(plan.updates.map(u => ({ id: u.id, fields: u.fields })));
  await createMany(plan.creates);
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
    const concurrent = (g2[k] || []).filter(t => t.createdTime >= runStart && !(t.fields['Notes'] || '').startsWith('Opening balance'));
    const expected = before[k] + concurrent.reduce((s2, t) => s2 + eff(t), 0);
    return Math.abs(now - expected) > EPS;
  });
  const report = summary + `\nAfter-run check: ${bad.length ? `⚠️ ${bad.length} pair(s) differ:\n` + bad.slice(0, 50).map(k => `${pName[k.split('|')[0]]} @ ${lName[k.split('|')[1]]}: expected ${before[k]}, now ${g2[k] ? stockOf(g2[k]) : 0}`).join('\n') : 'all stock levels identical ✅'}\nTable now: ${fresh.length} rows.`;
  await sendEmail(REPORT_EMAILS, bad.length ? '⚠️ Inventory compaction — check needed' : 'Inventory compaction — done ✅', report);
  console.log(report);
  if (bad.length) process.exitCode = 1;
}

if (require.main === module) main().catch(async e => {
  console.error(e);
  try { await sendEmail(REPORT_EMAILS, '⚠️ Inventory compaction stopped', String(e && e.stack || e)); } catch (_) {}
  process.exit(1);
});

module.exports = { buildPlan, stockOf };
