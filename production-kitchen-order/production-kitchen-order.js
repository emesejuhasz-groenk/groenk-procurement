/**
 * Grøenk — Production Kitchen auto-order
 *
 * Runs daily. As of 2026-09-01, triggered by the Daily Stock Consumption Deduction
 * workflow's own completion (workflow_run), plus a single 09:00-10:00 Madrid
 * safety-net schedule tick for days Deduction never triggers at all. See the .yml
 * for the full reasoning — this eliminates the fixed-time race between the two
 * workflows that a purely calendar-based schedule could no longer guarantee.
 *
 * For each of the 3 restaurants (Deià, Fornalutx, Sóller Pizza), computes a
 * suggested order quantity for every product supplied by "Groenk Production
 * Kitchen":
 *
 *   yesterday_sales = units sold on the single most recent calendar day with
 *         Daily Sales data (confirmed 2026-08-29: NOT an average of several
 *         days — just yesterday's actual number) × Recipes(BOM)
 *   par = yesterday_sales * days to cover * 1.5      (required stock, +50% buffer)
 *   order qty = max(0, ceil(par - current stock))   (always rounds UP)
 *   current stock = Inventory Transactions ledger sum (Opening Count +
 *                   Delivery Received + Manual Adjustment - Waste -
 *                   Consumption). Consumption rows are written daily by the
 *                   separate daily-consumption-deduction script, which turns
 *                   the previous day's Daily Sales into ledger entries via
 *                   the same BOM — so "current stock" here already reflects
 *                   yesterday's sales, not just receiving/waste/manual counts.
 *
 * ---------- Delivery schedule ----------
 * CHANGED 2026-09-03: Production Kitchen no longer delivers on Mondays —
 * reverted from the "every day of the week" rule set on 2026-08-29 (Emese
 * confirmed 2026-09-03 this is effective immediately). On a Monday, this
 * script does nothing at all: no order is computed, no email is sent, no
 * "missed" alert fires — there's genuinely nothing to order for a day with no
 * delivery, this isn't a failure state.
 *
 * STILL T+0 for now (TARGET_DAY_OFFSET = 0, unchanged since 2026-08-27): this
 * script orders TODAY for TODAY's own delivery. Switching to real T+1 (order
 * today for tomorrow) is a separate, deliberately deferred change — flipping
 * it on an arbitrary evening would skip an entire day's delivery, because
 * tomorrow morning's run would then target the day AFTER tomorrow instead of
 * tomorrow itself, and nothing would have ordered for tomorrow under the old
 * T+0 rule either (that already happened, or didn't, under the code as it
 * was this morning). That transition needs its own careful one-time handling
 * and will be done as a separate step once this Monday-skip has settled in.
 *
 * Production Kitchen delivery days: Tuesday-Sunday (6 days/week).
 *
 * TEMPORARY (since 2026-08-27, while stock/auto-order/receiving got back in
 * sync after a double-booking incident): order for TODAY (T) instead of
 * tomorrow (T+1). Revert TARGET_DAY_OFFSET to 1 once things have settled —
 * see the note above for why that revert needs care, not just flipping the
 * constant.
 * NOTE: this shifts EVERY downstream date calculation (target date, weekday)
 * by the same amount — no other code changes needed elsewhere in this file.
 *
 * Output:
 *   1. Creates an Order + Order Items record per restaurant in Airtable
 *      (Supplier = Groenk Production Kitchen) so they show up in the
 *      Procurement app's "Receive goods" screen exactly like a normal order.
 *   2. Sends ONE email to productionkitchengroenk@gmail.com with an XLSX
 *      attachment: rows = products (grouped by Order Category), columns =
 *      Deià / Fornalutx / Sóller Pizza / Total, so the production kitchen
 *      chef can prep from one file.
 *
 * Required environment variables (set as GitHub Actions secrets):
 *   AIRTABLE_TOKEN   — Personal Access Token, scoped to base appPcdy4HEJuDOF4j
 *                      with read+write (data.records:read, data.records:write)
 *   RESEND_API_KEY   — API key from resend.com (sign up with
 *                      productionkitchengroenk@gmail.com so the free sandbox
 *                      sender can email that same address without domain setup)
 */

const ExcelJS = require('exceljs');

const BASE_ID = 'appPcdy4HEJuDOF4j';
const AIRTABLE_TOKEN = process.env.AIRTABLE_TOKEN;
const RESEND_API_KEY = process.env.RESEND_API_KEY;
const EMAIL_TO = 'productionkitchengroenk@gmail.com';
const EMAIL_FROM = 'onboarding@resend.dev'; // Resend's shared sandbox sender — works with no domain setup as long as EMAIL_TO matches the Resend account's own signup address

// ---------- Trigger classification ----------
// This workflow has THREE possible triggers (see .yml):
//   - workflow_run   — fired automatically right after Daily Stock Consumption
//                       Deduction completes successfully. This is now the PRIMARY path.
//   - schedule       — a single once-a-day safety-net tick (09:00-10:00 Madrid), only
//                       relevant on the rare day Deduction doesn't trigger at all.
//   - workflow_dispatch — manual "Run workflow" click, for testing/backfilling.
// Both workflow_run and schedule represent genuine, unattended, trustworthy runs that
// should behave identically (readiness check applies, email auto-sends, dedup applies)
// — only workflow_dispatch is the "manual test, stay silent by default" path.
const GITHUB_EVENT_NAME = process.env.GITHUB_EVENT_NAME || '';
const IS_AUTOMATED_RUN = GITHUB_EVENT_NAME === 'schedule' || GITHUB_EVENT_NAME === 'workflow_run';
const SEND_EMAIL_OVERRIDE = process.env.SEND_EMAIL_OVERRIDE === 'true';
const MORNING_CUTOFF_HOUR = 10; // Madrid local time; automated runs at/after this hour still send, but the email gets a "this ran late" note appended

function madridHour(date) {
  return parseInt(
    new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/Madrid', hour: '2-digit', hour12: false }).format(date),
    10
  );
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

const BUFFER_MULTIPLIER = 1.5;
const TARGET_DAY_OFFSET = 0; // T+0, still — see the delivery-schedule comment at the top of the file
const PRODUCTION_KITCHEN_SUPPLIER_ID = 'recPXErB7VgvkYd6F'; // "Groenk Production Kitchen" in Suppliers table

// Restaurant app-name -> Locations table record id (Retail-role records; see index.html for the
// Production-Kitchen-has-two-roles note. Only the 3 restaurants order FROM the kitchen.)
const RESTAURANTS = {
  'Deia - Groenk Bistro': 'recnoXjgMS7jPYgE7',
  'Fornalutx - Groenk Bistro': 'recyfcAwYYZgFzSyd',
  'Soller - Groenk Pizza': 'reckUG4DXrJTMYtte',
};

if (!AIRTABLE_TOKEN || !RESEND_API_KEY) {
  console.error('Missing required environment variables. Need AIRTABLE_TOKEN, RESEND_API_KEY.');
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

async function airtableCreate(table, fields) {
  const res = await fetch(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${AIRTABLE_TOKEN}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ records: [{ fields }] }),
  });
  const data = await res.json();
  if (data.error) throw new Error(`Airtable create(${table}): ${data.error.message}`);
  return data.records[0];
}

async function airtableCreateMany(table, fieldsArray) {
  const results = [];
  for (let i = 0; i < fieldsArray.length; i += 10) {
    const chunk = fieldsArray.slice(i, i + 10);
    const res = await fetch(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${AIRTABLE_TOKEN}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ records: chunk.map(f => ({ fields: f })) }),
    });
    const data = await res.json();
    if (data.error) throw new Error(`Airtable createMany(${table}): ${data.error.message}`);
    results.push(...data.records);
  }
  return results;
}

async function airtableUpdateMany(table, updates) {
  const results = [];
  for (let i = 0; i < updates.length; i += 10) {
    const chunk = updates.slice(i, i + 10);
    const res = await fetch(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`, {
      method: 'PATCH',
      headers: { Authorization: `Bearer ${AIRTABLE_TOKEN}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ records: chunk }),
    });
    const data = await res.json();
    if (data.error) throw new Error(`Airtable updateMany(${table}): ${data.error.message}`);
    results.push(...data.records);
  }
  return results;
}

async function airtableDeleteMany(table, ids) {
  for (let i = 0; i < ids.length; i += 10) {
    const chunk = ids.slice(i, i + 10);
    const url = new URL(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`);
    for (const id of chunk) url.searchParams.append('records[]', id);
    const res = await fetch(url, { method: 'DELETE', headers: { Authorization: `Bearer ${AIRTABLE_TOKEN}` } });
    const data = await res.json();
    if (data.error) throw new Error(`Airtable deleteMany(${table}): ${data.error.message}`);
  }
}

// ---------- Date helpers ----------

function isoDate(d) {
  return d.toISOString().slice(0, 10);
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

// See the big comment block at the top of the file for the full reasoning.
// getDay(): 0=Sun, 1=Mon, 2=Tue, ... 6=Sat.
function isMonday(date) {
  return date.getDay() === 1;
}

// ---------- Main ----------

async function main() {
  const today = new Date();

  // Production Kitchen doesn't deliver on Mondays — nothing to order, nothing to
  // send, this isn't a missed/failure state so no alert either. Just stop here.
  if (isMonday(today)) {
    console.log(`${isoDate(today)} is a Monday — Production Kitchen has no delivery today. Nothing to do.`);
    return;
  }

  const targetDate = addDays(today, TARGET_DAY_OFFSET);
  const targetDateStr = isoDate(targetDate);
  const targetWeekday = targetDate.getDay();
  const daysToCover = 1; // always 1 under T+0 — see delivery-schedule comment at top of file
  console.log(`Run date: ${isoDate(today)} — target (T+${TARGET_DAY_OFFSET}) date: ${targetDateStr} (weekday ${targetWeekday}), trigger: "${GITHUB_EVENT_NAME}"`);

  const [products, recipes, dailySales, invTxns, locations] = await Promise.all([
    airtableGetAll('Products'),
    airtableGetAll('Recipes (BOM)'),
    airtableGetAll('Daily Sales'),
    airtableGetAll('Inventory Transactions'),
    airtableGetAll('Locations'),
  ]);

  // ---------- Wait for the daily consumption deduction to have caught up ----------
  // This script's "current stock" figure depends on the Daily Stock Consumption
  // Deduction job having already turned yesterday's Daily Sales into "Consumption"
  // Inventory Transactions. This workflow is normally triggered directly by that
  // job's own successful completion, so this check should almost always pass
  // immediately — it remains as a real data-based safety check (not a fixed time
  // assumption) rather than trusting the trigger alone, and it's what still
  // protects the once-a-day schedule safety-net path, which has no such
  // guarantee. Manual runs skip this check so testing/backfilling is never blocked.
  if (IS_AUTOMATED_RUN) {
    const eligible = s => (s.fields['Menu Item'] || []).length && (s.fields['Location'] || []).length;
    const mostRecentDate = dailySales
      .filter(eligible)
      .reduce((max, s) => {
        const d = (s.fields['Date'] || '').slice(0, 10);
        return d > max ? d : max;
      }, '');
    const todaysBatch = dailySales.filter(s => eligible(s) && (s.fields['Date'] || '').slice(0, 10) === mostRecentDate);
    const deductionCaughtUp = mostRecentDate && todaysBatch.length > 0 && todaysBatch.every(s => s.fields['Stock Deducted']);
    if (!deductionCaughtUp) {
      const reason = `most recent Daily Sales date: ${mostRecentDate || 'none'}, ` +
        `${todaysBatch.filter(s => !s.fields['Stock Deducted']).length}/${todaysBatch.length} of that day's records still un-deducted`;
      console.log(`Not ready on this automated run (trigger: "${GITHUB_EVENT_NAME}"): ${reason}. Sending a missed-window alert.`);
      try {
        await sendResendEmail({
          subject: `⚠️ Production Kitchen order MISSED for ${targetDateStr}`,
          text:
            `No Production Kitchen order was generated or sent today. This workflow now runs right after the ` +
            `Daily Stock Consumption Deduction workflow completes successfully (plus a 09:00–10:00 Madrid ` +
            `safety-net run on days that doesn't happen) — but even so, Daily Sales data for ${mostRecentDate || 'today'} ` +
            `wasn't fully marked as deducted yet.\n\n` +
            `Details: ${reason}.\n\n` +
            `Likely cause: the Cowork POS import task(s) didn't complete, or the Deduction workflow ran into a ` +
            `problem partway through despite reporting success. Check the Cowork Scheduled tasks page and the ` +
            `Airtable Daily Sales table, then run "Production Kitchen T+1 auto-order" manually (Run workflow, ` +
            `with send_email checked) once the data looks right.`,
        });
      } catch (e) {
        console.log(`Missed-window alert email also failed to send: ${e.message}`);
      }
      return;
    }
  }

  // Only products supplied by the Production Kitchen matter for this order.
  const pkProducts = products.filter(p => (p.fields['Supplier'] || []).includes(PRODUCTION_KITCHEN_SUPPLIER_ID));
  const pkProductIds = new Set(pkProducts.map(p => p.id));

  // Group BOM rows by Menu Item, keeping only components that are Production-Kitchen products.
  // ---------- Unit conversion ----------
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
  const KEG_PRODUCTS = {
    'recg4kyyl1P4ionxe': 30,
  };

  function convertQty(qty, fromUnit, toUnit, productId) {
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
    if (litersPerKeg && VOLUME_TO_ML[f]) {
      return (qty * VOLUME_TO_ML[f]) / 1000 / litersPerKeg;
    }

    const weightPerUnitG = Number(productWeightPerUnitById[productId]) || null;
    if (WEIGHT_TO_GRAMS[f] && weightPerUnitG) {
      return (qty * WEIGHT_TO_GRAMS[f]) / weightPerUnitG;
    }

    const startedAsWeightOrVolume = WEIGHT_TO_GRAMS[f] !== undefined || VOLUME_TO_ML[f] !== undefined;
    if (!startedAsWeightOrVolume) {
      const packSize = Number(productPackSizeById[productId]) || null;
      if (packSize) return qty / packSize;
    }

    return startedAsWeightOrVolume ? null : qty;
  }
  const productUnitById = Object.fromEntries(products.map(p => [p.id, p.fields['Unit'] || 'unit']));
  const productPackSizeById = Object.fromEntries(products.map(p => [p.id, p.fields['Pack Size']]));
  const productWeightPerUnitById = Object.fromEntries(products.map(p => [p.id, p.fields['Weight per Unit (g)']]));

  const bomByMenuItem = {};
  for (const r of recipes) {
    const menuItemIds = r.fields['Menu Item'] || [];
    const componentIds = r.fields['Component (Product)'] || [];
    const qtyPerUnit = Number(r.fields['Quantity per unit']) || 0;
    const bomUnit = r.fields['Unit'] || 'unit';
    for (const miId of menuItemIds) {
      for (const compId of componentIds) {
        if (!pkProductIds.has(compId)) continue;
        (bomByMenuItem[miId] = bomByMenuItem[miId] || []).push({ productId: compId, qtyPerUnit, bomUnit });
      }
    }
  }

  // ---------- Current stock (physical-count-aware) ----------
  const ledgerEffect = (type, qty) => (type === 'Waste' || type === 'Consumption') ? -Math.abs(qty) : qty;
  const isManualCount = t => t.fields['Type'] === 'Manual Adjustment' && String(t.fields['Notes'] || '').toLowerCase().includes('manual count');

  function currentStock(productId, locationId) {
    const txns = invTxns.filter(t => (t.fields['Related Product'] || []).includes(productId) && (t.fields['Location'] || []).includes(locationId));

    let lastCount = null;
    for (const t of txns) {
      if (isManualCount(t) && (!lastCount || t.createdTime > lastCount.createdTime)) lastCount = t;
    }
    if (!lastCount) {
      return txns.reduce((sum, t) => sum + ledgerEffect(t.fields['Type'], Number(t.fields['Quantity']) || 0), 0);
    }

    const baseline = txns
      .filter(t => t.createdTime < lastCount.createdTime)
      .reduce((sum, t) => sum + ledgerEffect(t.fields['Type'], Number(t.fields['Quantity']) || 0), 0)
      + ledgerEffect(lastCount.fields['Type'], Number(lastCount.fields['Quantity']) || 0);

    const after = txns.filter(t =>
      t.createdTime > lastCount.createdTime && (t.fields['Date'] || '') >= (lastCount.fields['Date'] || '')
    );
    return after.reduce((sum, t) => sum + ledgerEffect(t.fields['Type'], Number(t.fields['Quantity']) || 0), baseline);
  }

  // results[restaurantName][productId] = order quantity
  const results = {};
  for (const [restaurantName, locationId] of Object.entries(RESTAURANTS)) {
    const salesForLocation = dailySales.filter(s => (s.fields['Location'] || []).includes(locationId));
    const recentDates = [...new Set(salesForLocation.map(s => s.fields['Date']).filter(Boolean))]
      .sort()
      .reverse()
      .slice(0, 1);
    const recentSales = salesForLocation.filter(s => recentDates.includes(s.fields['Date']));
    console.log(`${restaurantName}: using yesterday's sales from ${recentDates[0] || '(no data)'}`);

    const salesByMenuItem = {};
    for (const s of recentSales) {
      const miIds = s.fields['Menu Item'] || [];
      const units = Number(s.fields['Units sold']) || 0;
      for (const miId of miIds) {
        (salesByMenuItem[miId] = salesByMenuItem[miId] || []).push(units);
      }
    }
    const avgUnitsSold = {};
    for (const [miId, arr] of Object.entries(salesByMenuItem)) {
      avgUnitsSold[miId] = arr.reduce((a, b) => a + b, 0);
    }

    const avgConsumption = {};
    for (const [miId, avgUnits] of Object.entries(avgUnitsSold)) {
      const bom = bomByMenuItem[miId];
      if (!bom) continue;
      for (const { productId, qtyPerUnit, bomUnit } of bom) {
        const converted = convertQty(qtyPerUnit, bomUnit, productUnitById[productId], productId);
        if (converted === null) {
          console.log(`Skipping ${productId}: no reliable conversion from "${bomUnit}" to "${productUnitById[productId]}" (missing pack size?)`);
          continue;
        }
        avgConsumption[productId] = (avgConsumption[productId] || 0) + avgUnits * converted;
      }
    }

    const restaurantResult = {};
    for (const [productId, avg] of Object.entries(avgConsumption)) {
      const par = avg * daysToCover * BUFFER_MULTIPLIER;
      const stock = currentStock(productId, locationId);
      const qty = Math.max(0, Math.ceil(par - stock));
      if (qty > 0) restaurantResult[productId] = qty;
    }
    results[restaurantName] = restaurantResult;
    console.log(`${restaurantName}: ${Object.keys(restaurantResult).length} products to order`);
  }

  // ---------- Write Orders + Order Items to Airtable (so they appear in Receive Goods) ----------

  const existingOrders = await airtableGetAll('Orders');
  const existingOrderItems = await airtableGetAll('Order Items');
  const ordersByRestaurant = {};
  for (const o of existingOrders) {
    const supplierIds = o.fields['Supplier'] || [];
    const orderDate = (o.fields['Order Date'] || '').slice(0, 10);
    if (supplierIds.includes(PRODUCTION_KITCHEN_SUPPLIER_ID) && orderDate === targetDateStr) {
      ordersByRestaurant[o.fields['Restaurant']] = o;
    }
  }
  const alreadyOrdered = new Set(Object.keys(ordersByRestaurant));

  for (const [restaurantName, productQtys] of Object.entries(results)) {
    const items = Object.entries(productQtys);
    const existingOrder = ordersByRestaurant[restaurantName];

    if (!existingOrder) {
      if (!items.length) continue;
      const order = await airtableCreate('Orders', {
        'Order Date': targetDateStr,
        'Supplier': [PRODUCTION_KITCHEN_SUPPLIER_ID],
        'Restaurant': restaurantName,
        'Status': 'Pending',
        'Created By': 'Auto (T+1)',
      });
      await airtableCreateMany('Order Items', items.map(([productId, qty]) => ({
        'Order': [order.id],
        'Product': [productId],
        'Quantity': qty,
        'Received': false,
        'Invoice Match': 'Pending Review',
      })));
      console.log(`Created Order for ${restaurantName} with ${items.length} items.`);
      continue;
    }

    const currentItems = existingOrderItems.filter(oi => (oi.fields['Order'] || []).includes(existingOrder.id));
    if (currentItems.some(oi => oi.fields['Received'])) {
      console.log(`Skipping ${restaurantName}: goods receiving has already started for today's order — not touching its items to avoid corrupting an in-progress receipt.`);
      continue;
    }

    const currentByProduct = {};
    for (const oi of currentItems) {
      const productId = (oi.fields['Product'] || [])[0];
      if (productId) currentByProduct[productId] = oi;
    }
    const wantedIds = new Set(items.map(([productId]) => productId));

    const toUpdate = items
      .filter(([productId, qty]) => currentByProduct[productId] && currentByProduct[productId].fields['Quantity'] !== qty)
      .map(([productId, qty]) => ({ id: currentByProduct[productId].id, fields: { 'Quantity': qty } }));
    const toCreate = items
      .filter(([productId]) => !currentByProduct[productId])
      .map(([productId, qty]) => ({
        'Order': [existingOrder.id],
        'Product': [productId],
        'Quantity': qty,
        'Received': false,
        'Invoice Match': 'Pending Review',
      }));
    const toDelete = currentItems.filter(oi => !wantedIds.has((oi.fields['Product'] || [])[0])).map(oi => oi.id);

    if (toUpdate.length) await airtableUpdateMany('Order Items', toUpdate);
    if (toCreate.length) await airtableCreateMany('Order Items', toCreate);
    if (toDelete.length) await airtableDeleteMany('Order Items', toDelete);
    if (toUpdate.length || toCreate.length || toDelete.length) {
      console.log(`Synced ${restaurantName}'s existing order: ${toUpdate.length} updated, ${toCreate.length} added, ${toDelete.length} removed.`);
    }
  }

  // ---------- Build XLSX ----------

  const productById = Object.fromEntries(products.map(p => [p.id, p.fields]));
  const restaurantNames = Object.keys(RESTAURANTS);
  const allProductIds = pkProductIds;

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Order');
  sheet.columns = [
    { header: 'Product', key: 'product', width: 36 },
    { header: 'Unit', key: 'unit', width: 14 },
    { header: 'Deià', key: 'deia', width: 12 },
    { header: 'Fornalutx', key: 'fornalutx', width: 12 },
    { header: 'Sóller Pizza', key: 'soller', width: 14 },
    { header: 'Total', key: 'total', width: 12 },
  ];
  sheet.getRow(1).font = { bold: true, size: 16 };
  sheet.getRow(1).height = 22;

  const CATEGORY_ORDER = ['Meat & Fish', 'Sauce', 'Bakery, pastry, dessert', 'Extra topping', 'Drink', 'Other'];
  const byCategory = {};

  for (const productId of allProductIds) {
    const p = productById[productId] || {};
    const deia = results['Deia - Groenk Bistro'][productId] || 0;
    const fornalutx = results['Fornalutx - Groenk Bistro'][productId] || 0;
    const soller = results['Soller - Groenk Pizza'][productId] || 0;
    const category = p['Order Category'] || 'Other';
    (byCategory[category] = byCategory[category] || []).push({
      product: p['Name'] || '(unknown product)',
      unit: p['Unit'] || '',
      deia, fornalutx, soller,
      total: deia + fornalutx + soller,
    });
  }

  for (const cat of CATEGORY_ORDER) {
    const rows = byCategory[cat];
    if (!rows || !rows.length) continue;
    rows.sort((a, b) => a.product.trim().localeCompare(b.product.trim()));
    const headerRow = sheet.addRow({ product: cat });
    headerRow.font = { bold: true, size: 15 };
    headerRow.height = 20;
    headerRow.eachCell(cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } }; });
    for (const row of rows) {
      const dataRow = sheet.addRow({
        product: row.product,
        unit: row.unit,
        deia: row.deia || '',
        fornalutx: row.fornalutx || '',
        soller: row.soller || '',
        total: row.total || '',
      });
      dataRow.font = { size: 13 };
      dataRow.height = 18;
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();

  // ---------- Send email via Resend ----------

  const currentMadridHour = madridHour(today);
  const isLateRun = currentMadridHour >= MORNING_CUTOFF_HOUR;
  const shouldSendEmail = IS_AUTOMATED_RUN || SEND_EMAIL_OVERRIDE;

  if (!shouldSendEmail) {
    console.log(
      `Skipping email send (event: "${GITHUB_EVENT_NAME}", override: ${SEND_EMAIL_OVERRIDE}). ` +
      `Email only sends automatically on an automated run (workflow_run or the schedule ` +
      `safety-net), or on a manual run with the send_email input set to true. Airtable ` +
      `writes above (if any) still happened normally.`
    );
    return;
  }

  if (IS_AUTOMATED_RUN) {
    const restaurantsNeedingOrder = Object.entries(results).filter(([, q]) => Object.keys(q).length > 0).map(([r]) => r);
    if (restaurantsNeedingOrder.length > 0 && restaurantsNeedingOrder.every(r => alreadyOrdered.has(r))) {
      console.log('All restaurants that need an order already have one for today — this looks like a later automated run catching up after a real send already went out. Skipping to avoid a duplicate email.');
      return;
    }
  }

  const lateNote = isLateRun
    ? ` (Note: this ran later than the usual morning time — around ${currentMadridHour}:00 Madrid time — most likely due to a GitHub Actions scheduling delay. The order contents are still correct.)`
    : '';
  const hasAnyOrders = restaurantNames.some(r => Object.keys(results[r]).length > 0);
  const offsetLabel = TARGET_DAY_OFFSET === 0 ? 'T (same-day)' : `T+${TARGET_DAY_OFFSET}`;
  const base64Attachment = buffer.toString('base64');
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${RESEND_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: EMAIL_FROM,
      to: [EMAIL_TO],
      subject: `Production Kitchen Order — ${targetDateStr}`,
      text: (hasAnyOrders
        ? `Attached: suggested order for ${targetDateStr} (${offsetLabel}), by restaurant and total. The sheet always lists every Production Kitchen product; rows with nothing to order are left blank.`
        : `No items to order for ${targetDateStr} — nothing crossed the buffer threshold. Attached anyway for reference (all rows blank).`) + lateNote,
      attachments: [{
        filename: `production-kitchen-order-${targetDateStr}.xlsx`,
        content: base64Attachment,
      }],
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
