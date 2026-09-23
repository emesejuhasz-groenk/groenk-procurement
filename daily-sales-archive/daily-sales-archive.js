/**
 * Grøenk — Daily Sales weekly archive
 *
 * ADDED 2026-09-23 (requested by Emese) to keep the Airtable base small and the
 * app fast. Airtable keeps only the current week + the last 4 completed
 * Monday-Sunday weeks of Daily Sales. Everything older is, one week per file:
 *
 *   1. written to a CSV
 *   2. uploaded to the controlling Drive folder "Daily Sales CSV"
 *   3. DOWNLOADED BACK from Drive and verified against Airtable:
 *        - same number of rows
 *        - exactly the same set of Airtable record IDs
 *        - same total units sold
 *        - identical file bytes (MD5 checksum reported by Drive)
 *   4. emailed to controlling@groenk.com (all weeks of the run in one email)
 *   5. ONLY IF step 3 passed for that week: deleted from Airtable.
 *
 * Hard safety rules (never relaxed):
 *   - A week whose Drive copy did not verify is NEVER deleted. An alert email
 *     goes out instead and the week is retried on the next run.
 *   - Records dated inside the retention window are never touched.
 *   - Records not yet processed by the daily consumption job (Stock Deducted
 *     unchecked) are never deleted — that would silently lose stock deductions.
 *     They stay in Airtable and are reported in the alert email.
 *   - DRY_RUN=true does steps 1-4 but deletes nothing.
 *
 * The first run also backfills: every complete week older than the retention
 * window is archived, one file per week. After that there is normally exactly
 * one week to archive each Monday. Re-running is safe: a week already in Drive
 * is verified instead of uploaded twice.
 *
 * Required environment variables:
 *   AIRTABLE_TOKEN, RESEND_API_KEY, APPS_SCRIPT_URL, APPS_SCRIPT_SECRET
 * Optional: DRY_RUN=true
 */

const crypto = require('crypto');

const BASE_ID = 'appPcdy4HEJuDOF4j';
const DAILY_SALES_TABLE = 'Daily Sales';
const DRIVE_FOLDER_ID = '1H6qGtPQtIVAaApFAhA6NLffHRdAcIjAJ'; // controlling Drive › "Daily Sales CSV"
const RETENTION_WEEKS = 4;          // completed weeks kept in Airtable (plus the current week)
const EMAIL_FROM = 'emese@groenk.com';
const CONTROLLING_EMAIL = 'controlling@groenk.com';
const ALERT_EMAILS = ['controlling@groenk.com', 'emese@groenk.com'];
const TIME_ZONE = 'Europe/Madrid';

const AIRTABLE_TOKEN = process.env.AIRTABLE_TOKEN;
const RESEND_API_KEY = process.env.RESEND_API_KEY;
const APPS_SCRIPT_URL = process.env.APPS_SCRIPT_URL;
const APPS_SCRIPT_SECRET = process.env.APPS_SCRIPT_SECRET;
const DRY_RUN = String(process.env.DRY_RUN || '').toLowerCase() === 'true';

if (!AIRTABLE_TOKEN || !RESEND_API_KEY || !APPS_SCRIPT_URL || !APPS_SCRIPT_SECRET) {
  console.error('Missing environment variables. Need AIRTABLE_TOKEN, RESEND_API_KEY, APPS_SCRIPT_URL, APPS_SCRIPT_SECRET.');
  process.exit(1);
}

const sleep = ms => new Promise(r => setTimeout(r, ms));

// ---------- Date helpers (plain YYYY-MM-DD strings, no time-zone drift) ----------
function todayInMadrid() {
  return new Intl.DateTimeFormat('en-CA', { timeZone: TIME_ZONE, year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
}
function addDaysStr(dateStr, n) {
  const d = new Date(dateStr + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}
function mondayOf(dateStr) {
  const d = new Date(dateStr + 'T00:00:00Z');
  const daysSinceMonday = (d.getUTCDay() + 6) % 7;
  return addDaysStr(dateStr, -daysSinceMonday);
}
function isoWeekLabel(mondayStr) {
  // ISO week number of the week starting on mondayStr
  const d = new Date(mondayStr + 'T00:00:00Z');
  const thursday = new Date(d); thursday.setUTCDate(d.getUTCDate() + 3);
  const yearStart = new Date(Date.UTC(thursday.getUTCFullYear(), 0, 1));
  const week = Math.ceil(((thursday - yearStart) / 86400000 + 1) / 7);
  return `${thursday.getUTCFullYear()}-W${String(week).padStart(2, '0')}`;
}

// ---------- Airtable ----------
async function airtableRequest(url, options = {}, maxRetries = 5) {
  let lastErr;
  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const res = await fetch(url, { ...options, headers: { Authorization: `Bearer ${AIRTABLE_TOKEN}`, 'Content-Type': 'application/json', ...(options.headers || {}) } });
      const text = await res.text();
      if (res.ok) return JSON.parse(text);
      lastErr = new Error(`Airtable HTTP ${res.status}: ${text.slice(0, 300)}`);
      if (res.status !== 429 && res.status < 500) throw lastErr;
    } catch (e) {
      lastErr = e;
      if (/HTTP 4(?!29)/.test(e.message)) throw e;
    }
    await sleep(1000 * Math.pow(2, attempt));
  }
  throw lastErr;
}
async function airtableGetAll(table, fields) {
  let records = [], offset;
  do {
    const url = new URL(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`);
    url.searchParams.set('pageSize', '100');
    (fields || []).forEach(f => url.searchParams.append('fields[]', f));
    if (offset) url.searchParams.set('offset', offset);
    const data = await airtableRequest(url.toString());
    records = records.concat(data.records);
    offset = data.offset;
    await sleep(220); // stay under 5 req/s
  } while (offset);
  return records;
}
async function airtableDelete(table, ids) {
  for (let i = 0; i < ids.length; i += 10) {
    const url = new URL(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`);
    ids.slice(i, i + 10).forEach(id => url.searchParams.append('records[]', id));
    await airtableRequest(url.toString(), { method: 'DELETE' });
    await sleep(220);
  }
}

// ---------- Google Drive via Apps Script (CHANGED 2026-09-23) ----------
// The groenk.com Google Cloud organisation blocks service-account keys, so Drive
// access goes through a small Google Apps Script web app that runs as the
// controlling@groenk.com account (see apps-script/Code.gs). Every call carries
// a shared secret. Content travels base64-encoded so the bytes stored in Drive
// are exactly the bytes generated here (needed for the checksum comparison).
async function appsScript(payload) {
  let lastErr;
  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      const res = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ ...payload, secret: APPS_SCRIPT_SECRET }),
        redirect: 'follow',
      });
      const text = await res.text();
      let data;
      try { data = JSON.parse(text); } catch (_) { throw new Error(`Apps Script returned non-JSON (HTTP ${res.status}): ${text.slice(0, 200)}`); }
      if (!data.ok) throw new Error(`Apps Script error: ${data.error}`);
      return data;
    } catch (e) {
      lastErr = e;
      if (/unauthorized/i.test(e.message)) throw e;
      await sleep(2000 * (attempt + 1));
    }
  }
  throw lastErr;
}
async function driveFindFile(name) {
  const data = await appsScript({ action: 'find', name });
  return data.file || null; // { id, name, size, md5Checksum }
}
async function driveUpload(name, content, existingId) {
  const data = await appsScript({ action: 'upload', name, existingId: existingId || null, contentBase64: Buffer.from(content, 'utf8').toString('base64') });
  return data.file;
}
async function driveDownload(fileId) {
  const data = await appsScript({ action: 'download', id: fileId });
  return Buffer.from(data.contentBase64, 'base64').toString('utf8');
}

// ---------- CSV ----------
const CSV_COLUMNS = ['Date', 'Location', 'Menu Item', 'Units sold', 'Revenue (EUR)', 'Stock Deducted', 'Airtable Record ID'];
function csvCell(v) {
  const s = v === null || v === undefined ? '' : String(v);
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}
function buildCsv(rows) {
  return '\uFEFF' + [CSV_COLUMNS.join(','), ...rows.map(r => CSV_COLUMNS.map(c => csvCell(r[c])).join(','))].join('\r\n') + '\r\n';
}
function parseCsv(text) {
  const src = text.replace(/^\uFEFF/, '');
  const rows = []; let row = [], cell = '', inQuotes = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (inQuotes) {
      if (ch === '"' && src[i + 1] === '"') { cell += '"'; i++; }
      else if (ch === '"') inQuotes = false;
      else cell += ch;
    } else if (ch === '"') inQuotes = true;
    else if (ch === ',') { row.push(cell); cell = ''; }
    else if (ch === '\n') { row.push(cell.replace(/\r$/, '')); rows.push(row); row = []; cell = ''; }
    else cell += ch;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  const [header, ...data] = rows.filter(r => r.length > 1 || r[0] !== '');
  return data.map(r => Object.fromEntries(header.map((h, i) => [h, r[i]])));
}

// Compares what came back from Drive with what Airtable holds for that week.
function verifyWeek(downloadedText, expectedRows, expectedCsv, driveMeta) {
  const problems = [];
  const parsed = parseCsv(downloadedText);
  if (parsed.length !== expectedRows.length) problems.push(`row count ${parsed.length} in Drive vs ${expectedRows.length} in Airtable`);
  const driveIds = new Set(parsed.map(r => r['Airtable Record ID']));
  const missing = expectedRows.filter(r => !driveIds.has(r['Airtable Record ID']));
  if (missing.length) problems.push(`${missing.length} Airtable records missing from the Drive file`);
  const sum = arr => arr.reduce((a, r) => a + (Number(r['Units sold']) || 0), 0);
  if (Math.abs(sum(parsed) - sum(expectedRows)) > 1e-6) problems.push(`units sold total ${sum(parsed)} in Drive vs ${sum(expectedRows)} in Airtable`);
  const localMd5 = crypto.createHash('md5').update(Buffer.from(expectedCsv, 'utf8')).digest('hex');
  if (driveMeta && driveMeta.md5Checksum && driveMeta.md5Checksum !== localMd5) problems.push('file checksum differs from the generated file');
  return problems;
}

// ---------- Email ----------
async function sendEmail({ to, subject, text, attachments }) {
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${RESEND_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ from: EMAIL_FROM, to, subject, text, attachments }),
  });
  const result = await res.json();
  if (result.error || !res.ok) throw new Error(`Resend failed: ${JSON.stringify(result.error || result).slice(0, 300)}`);
}

// ---------- Main ----------
async function main() {
  const today = todayInMadrid();
  const currentMonday = mondayOf(today);
  const cutoff = addDaysStr(currentMonday, -7 * RETENTION_WEEKS); // first day KEPT in Airtable
  // Belt and braces: the cutoff must always leave at least 4 full weeks.
  if (cutoff > addDaysStr(today, -28)) throw new Error(`Refusing to run: computed cutoff ${cutoff} is too recent for today ${today}.`);
  console.log(`Today ${today} (Madrid). Keeping ${cutoff} onward. Archiving complete weeks before ${cutoff}.${DRY_RUN ? ' DRY RUN — nothing will be deleted.' : ''}`);

  const [sales, menuItems, locations] = [
    await airtableGetAll(DAILY_SALES_TABLE, ['Date', 'Location', 'Menu Item', 'Units sold', 'Revenue (EUR)', 'Stock Deducted']),
    await airtableGetAll('Menu Items', ['Name']),
    await airtableGetAll('Locations', ['Name']),
  ];
  const menuName = Object.fromEntries(menuItems.map(m => [m.id, (m.fields['Name'] || '').trim()]));
  const locName = Object.fromEntries(locations.map(l => [l.id, (l.fields['Name'] || '').trim()]));

  // Group old records by week (client-side date comparison — the formula-based
  // date filter proved unreliable on this base, see daily-sales-watchdog.js).
  const weeks = {};
  const noDate = [];
  for (const r of sales) {
    const date = r.fields['Date'];
    if (!date) { noDate.push(r.id); continue; }
    if (date >= cutoff) continue;
    const monday = mondayOf(date);
    (weeks[monday] = weeks[monday] || []).push(r);
  }
  const mondays = Object.keys(weeks).sort();
  if (!mondays.length) { console.log('Nothing to archive.'); return; }

  const results = [];
  const attachments = [];
  for (const monday of mondays) {
    const sunday = addDaysStr(monday, 6);
    const label = isoWeekLabel(monday);
    const fileName = `daily-sales_${label}_${monday}_${sunday}.csv`;
    const records = weeks[monday];
    const rows = records
      .map(r => ({
        'Date': r.fields['Date'],
        'Location': (r.fields['Location'] || []).map(id => locName[id] || id).join(' | '),
        'Menu Item': (r.fields['Menu Item'] || []).map(id => menuName[id] || id).join(' | '),
        'Units sold': r.fields['Units sold'] ?? '',
        'Revenue (EUR)': r.fields['Revenue (EUR)'] ?? '',
        'Stock Deducted': r.fields['Stock Deducted'] ? 'yes' : 'no',
        'Airtable Record ID': r.id,
      }))
      .sort((a, b) => (a['Date'] + a['Location'] + a['Menu Item']).localeCompare(b['Date'] + b['Location'] + b['Menu Item']));
    const csv = buildCsv(rows);
    const result = { label, monday, sunday, fileName, rows: rows.length, verified: false, deleted: 0, kept: [], problems: [] };
    results.push(result);
    try {
      // Upload (or reuse the file from an earlier interrupted run), then verify.
      let meta = await driveFindFile(fileName);
      let problems = meta ? verifyWeek(await driveDownload(meta.id), rows, csv, meta) : ['not yet in Drive'];
      if (problems.length) {
        meta = await driveUpload(fileName, csv, meta ? meta.id : null);
        await sleep(1500);
        meta = (await driveFindFile(fileName)) || meta;
        problems = verifyWeek(await driveDownload(meta.id), rows, csv, meta);
      }
      if (problems.length) throw new Error(`Drive verification failed: ${problems.join('; ')}`);
      result.verified = true;
      result.driveFileId = meta.id;
      attachments.push({ filename: fileName, content: Buffer.from(csv, 'utf8').toString('base64') });
      console.log(`${fileName}: ${rows.length} rows uploaded and verified in Drive.`);
    } catch (e) {
      result.problems.push(e.message);
      console.error(`${fileName}: ${e.message} — this week will NOT be deleted.`);
    }
  }

  // Email the verified files to controlling before anything is deleted.
  const verified = results.filter(r => r.verified);
  let emailError = null;
  if (verified.length) {
    try {
      await sendEmail({
        to: [CONTROLLING_EMAIL],
        subject: `Daily Sales archive — ${verified.map(r => r.label).join(', ')}`,
        text: `Attached: Daily Sales CSV archive, one file per week.\n\n${verified.map(r => `${r.label} (${r.monday} – ${r.sunday}): ${r.rows} rows`).join('\n')}\n\nThe same files are saved in the controlling Drive folder "Daily Sales CSV".`,
        attachments,
      });
    } catch (e) { emailError = e.message; console.error(`Controlling email failed: ${e.message}`); }
  }

  // Delete ONLY verified weeks, and never unprocessed (Stock Deducted = no) rows.
  for (const r of verified) {
    const records = weeks[r.monday];
    const deletable = records.filter(x => x.fields['Stock Deducted'] === true).map(x => x.id);
    r.kept = records.filter(x => x.fields['Stock Deducted'] !== true).map(x => x.id);
    if (DRY_RUN) { console.log(`${r.fileName}: DRY RUN — would delete ${deletable.length} records.`); continue; }
    try {
      await airtableDelete(DAILY_SALES_TABLE, deletable);
      r.deleted = deletable.length;
      console.log(`${r.fileName}: deleted ${deletable.length} records from Airtable.${r.kept.length ? ` Kept ${r.kept.length} not-yet-deducted records.` : ''}`);
    } catch (e) { r.problems.push(`delete failed after ${r.deleted} records: ${e.message}`); }
  }

  // Alert if anything at all did not go perfectly.
  const issues = results.filter(r => r.problems.length || r.kept.length);
  if (issues.length || emailError || noDate.length) {
    const lines = [
      ...issues.map(r => `${r.label} (${r.monday} – ${r.sunday}): ${r.problems.length ? r.problems.join('; ') + ' — NOT deleted.' : ''}${r.kept.length ? ` ${r.kept.length} record(s) kept in Airtable because Stock Deducted is not ticked yet.` : ''}`),
      ...(emailError ? [`Controlling email failed (files ARE saved and verified in Drive): ${emailError}`] : []),
      ...(noDate.length ? [`${noDate.length} Daily Sales record(s) have no Date and were left untouched.`] : []),
    ];
    await sendEmail({ to: ALERT_EMAILS, subject: '⚠️ Daily Sales archive — needs attention', text: lines.join('\n') });
  }
  console.log(JSON.stringify(results.map(({ label, rows, verified, deleted, problems }) => ({ label, rows, verified, deleted, problems })), null, 2));
  if (results.some(r => r.problems.length)) process.exitCode = 1;
}

main().catch(async e => {
  console.error(e);
  try { await sendEmail({ to: ALERT_EMAILS, subject: '⚠️ Daily Sales archive failed — nothing was deleted', text: String(e && e.stack || e) }); } catch (_) {}
  process.exit(1);
});
