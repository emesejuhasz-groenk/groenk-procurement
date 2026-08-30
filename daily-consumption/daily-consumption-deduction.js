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
 * Since 2026-08-29, every Consumption row this script writes also gets the
 * resulting running stock balance stamped into "Stock Status After Transaction",
 * so the current stock per product/location is readable directly in Airtable
 * without manually summing the ledger. NOTE: this only covers rows THIS script
 * creates — Manual Adjustment / Delivery Received / Waste rows created by the web
 * app are a separate piece of work, not yet done.
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
const RESEND_API_KEY = process.env.RESEND_API_KEY;
const GITHUB_EVENT_NAME = process.env.GITHUB_EVENT_NAME || '';
const EMAIL_TO = 'productionkitchengroenk@gmail.com';
const EMAIL_FROM = 'onboarding@resend.dev';

if (!AIRTABLE_TOKEN) {
  console.error('Missing AIRTABLE_TOKEN environment variable.');
  process.exit(1);
}

// ---------- Missed-window alert ----------
// ADDED 2026-08-30: this script previously had NO alert at all if it silently failed to
// pick up a day's sales (e.g. GitHub Actions delay, HiOPOS/Cowork import running late or
// failing) — nobody found out until the PK order came out wrong or someone checked by
// hand. Mirrors the same pattern already used in production-kitchen-order.js: only the
// LAST scheduled tick of the window checks readiness and alerts, so it fires at most once
// a day and only once we're sure no earlier tick is still going to catch up.
const LAST_TICK_UTC_HOUR = 2;
const LAST_TICK_UTC_MINUTE = 30;

function isLastTickOfWindow(date) {
  return date.getUTCHours() === LAST_TICK_UTC_HOUR && date.getUTCMinutes() >= LAST_TICK_UTC_MINUTE;
}

async function sendResendEmail({ subject, text }) {
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${RESEND_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ from: EMAIL_FROM, to: [EMAIL_TO], subject, text }),
  });
  const result = await res.json();
  if (result.error) throw new Error(`Resend send failed: ${result.error.message || JSON.stringify(result.error)}`);
  console.log('Email sent:', result.id || result);
  return result;
}

// Same 3 restaurants the PK order script covers — used only to check that yesterday's
// sales actually arrived for each of them before the window closes.
const RESTAURANT_LOCATION_IDS = {
  'Deià': 'recnoXjgMS7jPYgE7',
  'Fornalutx': 'recyfcAwYYZgFzSyd',
  'Sóller Pizza': 'reckUG4DXrJTMYtte',
};

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

function madridDateStr(d) {
  // en-CA locale formats as YYYY-MM-DD, matching the Date field's own format.
  return new Intl.DateTimeFormat('en-CA', { timeZone: 'Europe/Madrid', year: 'numeric', month: '2-digit', day: '2-digit' }).format(d);
}

function addDaysToDateStr(dateStr, n) {
  const d = new Date(dateStr + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return isoDate(d);
}

async function main() {
  const now = new Date();
  const [dailySales, recipes, products, invTxns] = await Promise.all([
    airtableGetAll('Daily Sales'),
    airtableGetAll('Recipes (BOM)'),
    airtableGetAll('Products'),
    airtableGetAll('Inventory Transactions'),
  ]);

  const toProcess = dailySales.filter(s => !s.fields['Stock Deducted'] && (s.fields['Menu Item'] || []).length && (s.fields['Location'] || []).length);
  console.log(`${toProcess.length} Daily Sales record(s) not yet deducted.`);

  // ---------- Missed-window check (runs whether or not there's anything to process) ----------
  // Only meaningful on the LAST scheduled tick of the window: is yesterday's Daily Sales
  // data actually here yet, for every restaurant? If a restaurant has zero eligible rows
  // for yesterday at all, the Cowork POS import hasn't caught up (or failed), and no
  // amount of running this script will fix that — it needs a human to look at the import.
  if (GITHUB_EVENT_NAME === 'schedule' && isLastTickOfWindow(now)) {
    const yesterday = addDaysToDateStr(madridDateStr(now), -1);
    const missing = Object.entries(RESTAURANT_LOCATION_IDS).filter(([, locId]) =>
      !dailySales.some(s => (s.fields['Date'] || '').slice(0, 10) === yesterday && (s.fields['Location'] || []).includes(locId))
    ).map(([name]) => name);
    if (missing.length) {
      console.log(`Last tick of the window and still missing yesterday's (${yesterday}) sales for: ${missing.join(', ')}. Sending a missed-window alert.`);
      try {
        await sendResendEmail({
          subject: `⚠️ Daily Stock Consumption Deduction MISSED for ${yesterday}`,
          text:
            `No Daily Sales data was found for ${yesterday} for: ${missing.join(', ')} — the morning window ` +
            `(01:45–02:30 UTC / 03:45–04:30 Madrid) ended without it. Consumption was NOT deducted for ${missing.length === 1 ? 'this restaurant' : 'these restaurants'}, ` +
            `so today's stock figures (and the Production Kitchen order built from them) will be too high.\n\n` +
            `Likely cause: the Cowork POS import task(s) didn't complete or failed. Check the Cowork Scheduled ` +
            `tasks page and the Airtable Daily Sales table, then re-run "Daily Stock Consumption Deduction" ` +
            `manually (Run workflow) once the data looks right — it's safe to run any time, it only ever ` +
            `processes rows not yet marked "Stock Deducted".`,
        });
      } catch (e) {
        console.log(`Missed-window alert email also failed to send: ${e.message}`);
      }
    } else {
      console.log(`Last tick of the window — yesterday's (${yesterday}) sales are present for all restaurants. No alert needed.`);
    }
  }

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

  // ---------- Running balance ("Stock Status After Transaction") ----------
  // Starting 2026-08-29: every new transaction this script writes also gets the
  // resulting running stock balance stamped onto it, so the ledger is readable
  // directly in Airtable without having to sum the whole history by hand.
  //
  // Same physical-count-aware logic as the PK order script (see its comments for
  // the full reasoning): a plain chronological sum double-counts consumption
  // whenever a physical count is entered LATE and a day's consumption then gets
  // backfilled afterward for a date before that count — the count already
  // reflects that consumption in the real world, so re-applying it subtracts
  // twice. The baseline uses INSERTION order (createdTime) to reconstruct what
  // the count's delta was calibrated against; anything applied on top must be
  // dated on/after the count's own date, whether it already existed or is being
  // created in THIS run.
  const ledgerEffect = (type, qty) => (type === 'Waste' || type === 'Consumption') ? -Math.abs(qty) : qty;
  const isManualCount = t => t.fields['Type'] === 'Manual Adjustment' && String(t.fields['Notes'] || '').toLowerCase().includes('manual count');

  function stockBeforeThisRun(productId, locationId) {
    const txns = invTxns.filter(t => (t.fields['Related Product'] || []).includes(productId) && (t.fields['Location'] || []).includes(locationId));
    let lastCount = null;
    for (const t of txns) {
      if (isManualCount(t) && (!lastCount || t.createdTime > lastCount.createdTime)) lastCount = t;
    }
    if (!lastCount) {
      return { balance: txns.reduce((sum, t) => sum + ledgerEffect(t.fields['Type'], Number(t.fields['Quantity']) || 0), 0), lastCountDate: null };
    }
    const baseline = txns.filter(t => t.createdTime < lastCount.createdTime)
      .reduce((sum, t) => sum + ledgerEffect(t.fields['Type'], Number(t.fields['Quantity']) || 0), 0)
      + ledgerEffect(lastCount.fields['Type'], Number(lastCount.fields['Quantity']) || 0);
    const after = txns.filter(t => t.createdTime > lastCount.createdTime && (t.fields['Date'] || '') >= (lastCount.fields['Date'] || ''));
    const balance = after.reduce((sum, t) => sum + ledgerEffect(t.fields['Type'], Number(t.fields['Quantity']) || 0), baseline);
    return { balance, lastCountDate: lastCount.fields['Date'] || null };
  }

  const byProductLocation = {};
  for (const rec of consumptionRecords) {
    const key = `${rec['Related Product'][0]}|${rec['Location'][0]}`;
    (byProductLocation[key] = byProductLocation[key] || []).push(rec);
  }
  for (const [key, recs] of Object.entries(byProductLocation)) {
    recs.sort((a, b) => (a['Date'] < b['Date'] ? -1 : a['Date'] > b['Date'] ? 1 : 0));
    const [productId, locationId] = key.split('|');
    const { balance, lastCountDate } = stockBeforeThisRun(productId, locationId);
    let running = balance;
    for (const rec of recs) {
      if (lastCountDate && rec['Date'] < lastCountDate) {
        // Backfilling a date before the most recent physical count — already
        // implicitly reflected in that count's number, so don't apply it to the
        // running total (would double-count); just stamp the unaffected balance.
        rec['Stock Status After Transaction'] = String(Math.round(running * 1000) / 1000);
        continue;
      }
      running = ledgerEffect(rec['Type'], rec['Quantity']) + running;
      rec['Stock Status After Transaction'] = String(Math.round(running * 1000) / 1000);
    }
  }

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
