/**
 * Grøenk — Weekly Inventory Audit
 *
 * Runs every Saturday at 06:00 (Madrid). Looks at every "Manual Adjustment"
 * Inventory Transaction from the last 7 days (i.e. every time someone used
 * "Update count" in the Procurement app's Inventory screen to correct the
 * system-computed stock with what they physically counted) and reports:
 *
 *   Date | Location | Product | System qty | Manually counted qty | Difference | Counted by
 *
 * "System qty" is the running stock the ledger had calculated for that
 * product+location immediately BEFORE the correction — i.e. what the count
 * disagreed with. This is recomputed from the full transaction history (not
 * just the last 7 days), because the running balance depends on everything
 * that came before it.
 *
 * Negative differences (physically counted LESS than the system expected —
 * the case Emese wants to catch) are highlighted red in the XLSX.
 *
 * This does NOT look at Waste-type transactions — those already have their
 * own reason/category and aren't part of this fraud/shrinkage check.
 *
 * Output: one email to controlling@groenk.com with an XLSX attachment.
 *
 * Required environment variables (GitHub Actions secrets):
 *   AIRTABLE_TOKEN   — same token used by the other scripts (read access is enough)
 *   RESEND_API_KEY   — same Resend key used by the T+1 script
 */

const ExcelJS = require('exceljs');

const BASE_ID = 'appPcdy4HEJuDOF4j';
const AIRTABLE_TOKEN = process.env.AIRTABLE_TOKEN;
const RESEND_API_KEY = process.env.RESEND_API_KEY;
const EMAIL_TO = 'productionkitchengroenk@gmail.com';
const EMAIL_FROM = 'onboarding@resend.dev';

const LOOKBACK_DAYS = 7;

// Locations table record id -> display name (mirrors index.html's LOCATION_IDS, reversed)
const LOCATION_NAMES = {
  'reckUG4DXrJTMYtte': 'Soller - Groenk Pizza',
  'recnoXjgMS7jPYgE7': 'Deia - Groenk Bistro',
  'recyfcAwYYZgFzSyd': 'Fornalutx - Groenk Bistro',
  'recZespbXJli9GI4r': 'Groenk Production Kitchen Fornalutx',
};

if (!AIRTABLE_TOKEN || !RESEND_API_KEY) {
  console.error('Missing required environment variables. Need AIRTABLE_TOKEN, RESEND_API_KEY.');
  process.exit(1);
}

async function airtableGetAll(table) {
  const headers = { Authorization: `Bearer ${AIRTABLE_TOKEN}` };
  let records = [];
  let offset;
  do {
    const url = new URL(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`);
    url.searchParams.set('pageSize', '100');
    if (offset) url.searchParams.set('offset', offset);
    const res = await fetch(url, { headers });
    const data = await res.json();
    if (data.error) throw new Error(`Airtable getAll(${table}): ${data.error.message}`);
    records = records.concat(data.records);
    offset = data.offset;
  } while (offset);
  return records;
}

function isoDate(d) { return d.toISOString().slice(0, 10); }
function addDays(date, n) { const d = new Date(date); d.setDate(d.getDate() + n); return d; }

async function main() {
  const today = new Date();
  const cutoffStr = isoDate(addDays(today, -LOOKBACK_DAYS));
  console.log(`Weekly audit run: ${isoDate(today)} — looking at Manual Adjustment transactions since ${cutoffStr}`);

  const [txns, products] = await Promise.all([
    airtableGetAll('Inventory Transactions'),
    airtableGetAll('Products'),
  ]);

  const productNameById = Object.fromEntries(products.map(p => [p.id, p.fields['Name'] || '(unknown product)']));

  // Group all transactions by product+location, sorted chronologically, so we can walk
  // the running balance and know exactly what "system qty" was right before each
  // Manual Adjustment.
  const byKey = {};
  for (const t of txns) {
    const productId = (t.fields['Related Product'] || [])[0];
    const locationId = (t.fields['Location'] || [])[0];
    if (!productId || !locationId) continue;
    const key = `${productId}|${locationId}`;
    (byKey[key] = byKey[key] || []).push(t);
  }
  for (const key of Object.keys(byKey)) {
    byKey[key].sort((a, b) => {
      const d = (a.fields['Date'] || '').localeCompare(b.fields['Date'] || '');
      if (d !== 0) return d;
      return (a.createdTime || '').localeCompare(b.createdTime || '');
    });
  }

  const auditRows = [];
  for (const [key, list] of Object.entries(byKey)) {
    const [productId, locationId] = key.split('|');
    let balance = 0;
    for (const t of list) {
      const qty = Number(t.fields['Quantity']) || 0;
      const type = t.fields['Type'];
      const isManualAdjustment = type === 'Manual Adjustment';
      const dateStr = t.fields['Date'] || '';

      if (isManualAdjustment && dateStr >= cutoffStr) {
        auditRows.push({
          date: dateStr,
          location: LOCATION_NAMES[locationId] || locationId,
          product: productNameById[productId] || productId,
          systemQty: balance,
          manualQty: balance + qty,
          difference: qty,
          countedBy: t.fields['Transaction Added By'] || '',
        });
      }

      // Keep the running balance current for every transaction type, same rule as the app.
      balance += type === 'Waste' ? -Math.abs(qty) : qty;
    }
  }

  auditRows.sort((a, b) => a.date.localeCompare(b.date) || a.location.localeCompare(b.location) || a.product.localeCompare(b.product));

  console.log(`Found ${auditRows.length} manual count correction(s) in the last ${LOOKBACK_DAYS} days.`);

  // ---------- Build XLSX ----------

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Weekly Audit');
  sheet.columns = [
    { header: 'Date', key: 'date', width: 12 },
    { header: 'Location', key: 'location', width: 26 },
    { header: 'Product', key: 'product', width: 32 },
    { header: 'System qty', key: 'systemQty', width: 12 },
    { header: 'Manually counted qty', key: 'manualQty', width: 18 },
    { header: 'Difference', key: 'difference', width: 12 },
    { header: 'Counted by', key: 'countedBy', width: 18 },
  ];
  sheet.getRow(1).font = { bold: true };

  for (const row of auditRows) {
    const excelRow = sheet.addRow(row);
    if (row.difference < 0) {
      const diffCell = excelRow.getCell('difference');
      diffCell.font = { color: { argb: 'FFCC0000' }, bold: true };
      excelRow.getCell('systemQty').font = { color: { argb: 'FFCC0000' } };
      excelRow.getCell('manualQty').font = { color: { argb: 'FFCC0000' } };
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();

  // ---------- Send email ----------

  const base64Attachment = buffer.toString('base64');
  const rangeLabel = `${cutoffStr} to ${isoDate(today)}`;
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${RESEND_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      from: EMAIL_FROM,
      to: [EMAIL_TO],
      subject: `Weekly Inventory Audit — ${rangeLabel}`,
      text: auditRows.length
        ? `Attached: ${auditRows.length} manual inventory count correction(s) from ${rangeLabel}. Rows in red show a count that came in LOWER than the system-expected stock.`
        : `No manual inventory count corrections were recorded between ${rangeLabel}.`,
      attachments: auditRows.length ? [{
        filename: `weekly-inventory-audit-${isoDate(today)}.xlsx`,
        content: base64Attachment,
      }] : [],
    }),
  });
  const emailResult = await res.json();
  if (!res.ok) throw new Error(`Resend send failed: ${JSON.stringify(emailResult)}`);

  console.log('Email sent to', EMAIL_TO, '— Resend id:', emailResult.id);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
