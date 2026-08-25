/**
 * Grøenk — Daily Stock Consumption Deduction
 *
 * Runs daily (via GitHub Actions cron), shortly after the Cowork automation that
 * imports the previous day's Daily Sales from the POS emails (~06:00).
 *
 * For every Daily Sales record not yet marked "Stock Deducted": looks up the Menu
 * Item's Recipes (BOM), converts each component's per-portion usage into the
 * component Product's own stocking unit (same conversion logic as the other
 * scripts — direct weight/volume, bottle/box, bulk-liter, keg, generic
 * Weight-per-Unit, generic Pack Size), multiplies by Units sold, and writes ONE
 * aggregated "Consumption" Inventory Transaction per (date, location, product) —
 * so a T-bone and a Chuletón sold the same day at the same place add up into a
 * single BBQ-sauce consumption line, not two.
 *
 * This makes the "in stock" figure track actual sales in near-real-time, instead
 * of only being corrected at the weekly physical count. The Friday inventory
 * count remains useful as a periodic correction for human error / measurement
 * drift, but is no longer the only thing keeping stock accurate.
 *
 * On first run this will process the ENTIRE historical backlog of Daily Sales
 * (since none of it has "Stock Deducted" set yet) — expect a large first run;
 * every run after that is just the new day's sales.
 *
 * Required environment variable (GitHub Actions secret):
 *   AIRTABLE_TOKEN — same token used by the other scripts (needs write access)
 */

const BASE_ID = 'appPcdy4HEJuDOF4j';
const AIRTABLE_TOKEN = process.env.AIRTABLE_TOKEN;

if (!AIRTABLE_TOKEN) {
  console.error('Missing AIRTABLE_TOKEN environment variable.');
  process.exit(1);
}

// ---------- Airtable helpers ----------

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

async function airtableCreateMany(table, fieldsArray) {
  const results = [];
  for (let i = 0; i < fieldsArray.length; i += 25) {
    const chunk = fieldsArray.slice(i, i + 25);
    const res = await fetch(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${AIRTABLE_TOKEN}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ records: chunk.map(f => ({ fields: f })), typecast: true }),
    });
    const data = await res.json();
    if (data.error) throw new Error(`Airtable createMany(${table}): ${JSON.stringify(data.error)}`);
    results.push(...data.records);
  }
  return results;
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

// ---------- Unit conversion (same logic as the other scripts) ----------

const WEIGHT_TO_GRAMS = { g: 1, gr: 1, gramm: 1, kg: 1000 };
const VOLUME_TO_ML = { ml: 1, cl: 10, dl: 100, l: 1000, liter: 1000, litre: 1000 };
const BOTTLE_PRODUCTS = {
  'rec0GG6FuoGVhppOX': { bottleMl: 750, boxBottles: 6 },
  'recr5eRYcQ9DjtF8N': { bottleMl: 750, boxBottles: 12 },
  'recNPOqxkKjbxruPm': { bottleMl: 750, boxBottles: 6 },
  'recfaXavJa7Mfn95b': { bottleMl: 750, boxBottles: 6 },
  'recOur09D12vya1TO': { bottleMl: 750, boxBottles: 6 },
  'recFkhzAQHnJQFOv6': { bottleMl: 750, boxBottles: 6 },
  'recqcp3W6BLYP7HLz': { bottleMl: 750, boxBottles: 6 },
  'recbmyjGTpXCAgRdf': { bottleMl: 750, boxBottles: 6 },
  'recvWB7SPBW8yCbyw': { bottleMl: 750, boxBottles: 6 },
  'rectHW7Se1XBPFbF7': { bottleMl: 750, boxBottles: 6 },
  'recuDpwaPhpDe5Qh0': { bottleMl: 750, boxBottles: 6 },
  'rec5lgdtonSCCuXKn': { bottleMl: 750, boxBottles: 6 },
  'recBZqhgUHiq9bP1d': { bottleMl: 500, boxBottles: 6 },
  'recYff3VN3LLhBP3R': { bottleMl: 700, boxBottles: 6 },
  'recJVtCM5aKx3fkUs': { bottleMl: 700, boxBottles: 6 },
  'recO3hYm1bTCEbHwE': { bottleMl: 750, boxBottles: 6 },
  'recLnrqeL5oKywzpO': { bottleMl: 750, boxBottles: 6 },
  'recrvnmrObLfFgOZO': { bottleMl: 200, boxBottles: 24 },
  'recLWTdCvXT0VPUoV': { bottleMl: 700, boxBottles: 12 },
};
const BULK_LITER_PRODUCTS = {
  'recQuqcE97aPWh9gx': 3, 'recQZGUy989QpkbZe': 3, 'recJmkJvH3IRDwWrc': 3, 'recoPQdKRSxB8nKEt': 3,
  'recMaHap0Kptnyc8N': 10,
};
const KEG_PRODUCTS = { 'recg4kyyl1P4ionxe': 30 };

function convertQty(qty, fromUnit, toUnit, productId, packSize, weightPerUnitG) {
  const f = String(fromUnit || '').toLowerCase().trim();
  const t = String(toUnit || '').toLowerCase().trim();
  if (f === t) return qty;
  if (WEIGHT_TO_GRAMS[f] && WEIGHT_TO_GRAMS[t]) return qty * WEIGHT_TO_GRAMS[f] / WEIGHT_TO_GRAMS[t];
  if (VOLUME_TO_ML[f] && VOLUME_TO_ML[t]) return qty * VOLUME_TO_ML[f] / VOLUME_TO_ML[t];
  const bottleInfo = BOTTLE_PRODUCTS[productId];
  if (bottleInfo && t.includes('box')) {
    const bottles = f === 'bottle' ? qty : (VOLUME_TO_ML[f] ? (qty * VOLUME_TO_ML[f]) / bottleInfo.bottleMl : qty);
    return bottles / bottleInfo.boxBottles;
  }
  const litersPerBox = BULK_LITER_PRODUCTS[productId];
  if (litersPerBox && (t.includes('box') || t.includes('pack'))) {
    const ml = VOLUME_TO_ML[f] ? qty * VOLUME_TO_ML[f] : qty * 1000;
    return ml / 1000 / litersPerBox;
  }
  const litersPerKeg = KEG_PRODUCTS[productId];
  if (litersPerKeg && VOLUME_TO_ML[f]) return (qty * VOLUME_TO_ML[f]) / 1000 / litersPerKeg;
  if (WEIGHT_TO_GRAMS[f] && weightPerUnitG) return (qty * WEIGHT_TO_GRAMS[f]) / weightPerUnitG;
  const startedAsWeightOrVolume = WEIGHT_TO_GRAMS[f] !== undefined || VOLUME_TO_ML[f] !== undefined;
  if (!startedAsWeightOrVolume && packSize) return qty / packSize;
  return startedAsWeightOrVolume ? null : qty;
}

function isoDate(d) { return d.toISOString().slice(0, 10); }

async function main() {
  const [dailySales, recipes, products] = await Promise.all([
    airtableGetAll('Daily Sales'),
    airtableGetAll('Recipes (BOM)'),
    airtableGetAll('Products'),
  ]);

  const toProcess = dailySales.filter(s => !s.fields['Stock Deducted'] && (s.fields['Menu Item'] || []).length && (s.fields['Location'] || []).length);
  console.log(`${toProcess.length} Daily Sales record(s) not yet deducted.`);
  if (!toProcess.length) { console.log('Nothing to do.'); return; }

  const productUnitById = Object.fromEntries(products.map(p => [p.id, p.fields['Unit'] || 'unit']));
  const productPackSizeById = Object.fromEntries(products.map(p => [p.id, p.fields['Pack Size']]));
  const productWeightPerUnitById = Object.fromEntries(products.map(p => [p.id, p.fields['Weight per Unit (g)']]));

  const bomByMenuItem = {};
  for (const r of recipes) {
    const qtyPerUnit = Number(r.fields['Quantity per unit']) || 0;
    const bomUnit = r.fields['Unit'] || 'unit';
    for (const miId of (r.fields['Menu Item'] || [])) {
      for (const productId of (r.fields['Component (Product)'] || [])) {
        (bomByMenuItem[miId] = bomByMenuItem[miId] || []).push({ productId, qtyPerUnit, bomUnit });
      }
    }
  }

  // Aggregate consumption by (date, locationId, productId) across every sale being processed.
  const agg = {}; // key -> { date, locationId, productId, qty }
  const unreliable = new Set();
  for (const sale of toProcess) {
    const date = sale.fields['Date'];
    const locationId = (sale.fields['Location'] || [])[0];
    const unitsSold = Number(sale.fields['Units sold']) || 0;
    if (!date || !locationId || unitsSold <= 0) continue;
    for (const miId of (sale.fields['Menu Item'] || [])) {
      const bom = bomByMenuItem[miId];
      if (!bom) continue;
      for (const { productId, qtyPerUnit, bomUnit } of bom) {
        const converted = convertQty(qtyPerUnit, bomUnit, productUnitById[productId], productId, productPackSizeById[productId], productWeightPerUnitById[productId]);
        if (converted === null) { unreliable.add(productId); continue; }
        const key = `${date}|${locationId}|${productId}`;
        if (!agg[key]) agg[key] = { date, locationId, productId, qty: 0 };
        agg[key].qty += unitsSold * converted;
      }
    }
  }

  const consumptionRecords = Object.values(agg)
    .filter(a => a.qty > 0)
    .map(a => ({
      'Date': a.date,
      'Type': 'Consumption',
      'Quantity': Math.round(a.qty * 1000) / 1000,
      'Related Product': [a.productId],
      'Location': [a.locationId],
      'Notes': 'Auto-deducted from Daily Sales',
    }));

  console.log(`Writing ${consumptionRecords.length} aggregated Consumption transaction(s).`);
  if (consumptionRecords.length) await airtableCreateMany('Inventory Transactions', consumptionRecords);

  await airtableUpdateMany('Daily Sales', toProcess.map(s => ({ id: s.id, fields: { 'Stock Deducted': true } })));
  console.log(`Marked ${toProcess.length} Daily Sales record(s) as Stock Deducted.`);

  if (unreliable.size) {
    console.log(`\n⚠ ${unreliable.size} product(s) had sales but no reliable unit conversion (missing Weight per Unit / Pack Size) — their consumption was NOT deducted:`);
    for (const id of unreliable) {
      const p = products.find(p => p.id === id);
      console.log(`  - ${p ? p.fields['Name'] : id}`);
    }
  }
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
