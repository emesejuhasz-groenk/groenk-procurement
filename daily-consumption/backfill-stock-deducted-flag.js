/**
 * Grøenk — ONE-OFF: baseline Daily Sales "Stock Deducted" flag
 *
 * Run this ONCE, manually, before turning on the daily consumption deduction
 * script. It marks every EXISTING Daily Sales record as "Stock Deducted = true"
 * WITHOUT creating any Inventory Transactions — i.e. it does not touch stock at
 * all, it just fast-forwards the flag so historical sales are never processed.
 *
 * Why: Emese just did fresh physical inventory counts (drinks today, food two
 * days ago). Deducting the full sales history now would double-count against
 * those counts. This establishes a clean starting line — from tomorrow's Daily
 * Sales onward, the regular daily-consumption-deduction.js script takes over
 * and actually deducts stock for new sales only.
 *
 * Required environment variable: AIRTABLE_TOKEN
 * Usage: node backfill-stock-deducted-flag.js
 */

const BASE_ID = 'appPcdy4HEJuDOF4j';
const AIRTABLE_TOKEN = process.env.AIRTABLE_TOKEN;

if (!AIRTABLE_TOKEN) {
  console.error('Missing AIRTABLE_TOKEN environment variable.');
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

async function airtableUpdateMany(table, updates) {
  const results = [];
  for (let i = 0; i < updates.length; i += 25) {
    const chunk = updates.slice(i, i + 25);
    const res = await fetch(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${AIRTABLE_TOKEN}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ records: chunk, typecast: true }),
    });
    const data = await res.json();
    if (data.error) throw new Error(`Airtable updateMany(${table}): ${JSON.stringify(data.error)}`);
    results.push(...data.records);
  }
  return results;
}

async function main() {
  const dailySales = await airtableGetAll('Daily Sales');
  const toFlag = dailySales.filter(s => !s.fields['Stock Deducted']);
  console.log(`${dailySales.length} total Daily Sales records, ${toFlag.length} not yet flagged.`);
  if (!toFlag.length) { console.log('Nothing to do — already all flagged.'); return; }

  await airtableUpdateMany('Daily Sales', toFlag.map(s => ({ id: s.id, fields: { 'Stock Deducted': true } })));
  console.log(`Flagged ${toFlag.length} record(s) as Stock Deducted = true. No stock was touched.`);
  console.log('From here on, only new Daily Sales rows (tomorrow onward) will be deducted by the daily script.');
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
