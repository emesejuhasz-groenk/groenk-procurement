/**
 * Grøenk — Production Kitchen auto-order
 *
 * Runs daily (via GitHub Actions cron). For each of the 3 restaurants
 * (Deià, Fornalutx, Sóller Pizza), computes a suggested order quantity for
 * every product supplied by "Groenk Production Kitchen":
 *
 *   avg = average daily consumption over the LAST 2 CALENDAR DAYS with Daily
 *         Sales data (no weekday-matching — demand is fairly stable day to
 *         day in Mallorca, per Emese) × Recipes(BOM)
 *   par = avg * days to cover * 1.5            (buffer, +50%)
 *   order qty = max(0, ceil(par - current stock))   (always rounds UP)
 *   current stock = Inventory Transactions ledger sum (Opening Count +
 *                   Delivery Received + Manual Adjustment - Waste -
 *                   Consumption). Consumption rows are written daily by the
 *                   separate daily-consumption-deduction script, which turns
 *                   the previous day's Daily Sales into ledger entries via
 *                   the same BOM — so "current stock" here already reflects
 *                   yesterday's sales, not just receiving/waste/manual counts.
 *
 * Target date = run date + TARGET_DAY_OFFSET days (see constant below).
 * NORMALLY this is T+1 (order today for tomorrow's delivery), but it is
 * TEMPORARILY set to T+0 (order today for TODAY's delivery) starting
 * 2026-08-27, while stock / auto-order / receiving get back in sync after
 * a double-booking incident. Revert TARGET_DAY_OFFSET to 1 once things
 * have settled — nothing else in this file needs to change to switch back.
 *
 * Whichever mode is active, the Production Kitchen delivers Tuesday–Sunday
 * (NOT Monday), checked against the TARGET date's weekday (so this adapts
 * automatically whether TARGET_DAY_OFFSET is 0 or 1):
 *   - if the target date is a Monday, there's no Monday delivery, so the
 *     script exits without doing anything for that restaurant/day — the
 *     next day's run will compute Tuesday's order instead
 *   - if the target date is a Sunday, that delivery has to cover BOTH
 *     Sunday and Monday (2 days) since nothing arrives again until Tuesday
 *   - every other target weekday covers exactly 1 day
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

// ---------- Email send guard ----------
// The Airtable idempotency guard below only stops duplicate Orders — it does NOT stop
// this script from re-sending the email every time it runs, which confused the kitchen
// staff when a manual test run (or a scheduled run that GitHub Actions delayed by several
// hours) fired outside the normal morning window.
//
// Rule: only actually send the email when this is a genuine SCHEDULED run (not a manual
// "Run workflow" / workflow_dispatch test) AND it's still morning in Madrid. Everything
// else (Airtable writes, xlsx build) still happens on every run, so manual runs remain
// useful for testing/backfilling data — they just stay silent by default.
// To deliberately test the email itself from a manual run, use "Run workflow" and set
// the send_email input to true.
const GITHUB_EVENT_NAME = process.env.GITHUB_EVENT_NAME || '';
const SEND_EMAIL_OVERRIDE = process.env.SEND_EMAIL_OVERRIDE === 'true';
const MORNING_CUTOFF_HOUR = 10; // Madrid local time; scheduled runs at/after this hour are treated as "not morning" and skip the email

function madridHour(date) {
  return parseInt(
    new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/Madrid', hour: '2-digit', hour12: false }).format(date),
    10
  );
}

const BUFFER_MULTIPLIER = 1.5;

// TEMPORARY (a few weeks, starting 2026-08-27): order for TODAY (T) instead of
// tomorrow (T+1), while stock / auto-order / receiving get back in sync after the
// double-booking incident. Revert to 1 once things have settled.
// NOTE: this shifts EVERY downstream date calculation (target date, weekday checks,
// Monday-skip, Saturday double-coverage) by the same amount — no other code changes
// needed elsewhere in this file.
const TARGET_DAY_OFFSET = 0;
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

// ---------- Date helpers ----------

function isoDate(d) {
  return d.toISOString().slice(0, 10);
}

function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}

// ---------- Main ----------

async function main() {
  const today = new Date();
  const targetDate = addDays(today, TARGET_DAY_OFFSET);
  const targetDateStr = isoDate(targetDate);
  const targetWeekday = targetDate.getDay(); // 0=Sun..6=Sat
  console.log(`Run date: ${isoDate(today)} — target (T+${TARGET_DAY_OFFSET}) date: ${targetDateStr} (weekday ${targetWeekday})`);

  if (targetWeekday === 1) {
    console.log('Target date is a Monday — Production Kitchen does not deliver on Mondays. Nothing to do; Tuesday\'s order will be computed by tomorrow\'s (Monday) run instead.');
    return;
  }
  // Saturday's run (target = Sunday) has to cover Sunday AND Monday, since there's no
  // Monday delivery to top up again before Tuesday.
  const daysToCover = targetWeekday === 0 ? 2 : 1;
  console.log(`Days to cover: ${daysToCover}`);

  const [products, recipes, dailySales, invTxns, locations] = await Promise.all([
    airtableGetAll('Products'),
    airtableGetAll('Recipes (BOM)'),
    airtableGetAll('Daily Sales'),
    airtableGetAll('Inventory Transactions'),
    airtableGetAll('Locations'),
  ]);

  // Only products supplied by the Production Kitchen matter for this order.
  const pkProducts = products.filter(p => (p.fields['Supplier'] || []).includes(PRODUCTION_KITCHEN_SUPPLIER_ID));
  const pkProductIds = new Set(pkProducts.map(p => p.id));

  // Group BOM rows by Menu Item, keeping only components that are Production-Kitchen products.
  // ---------- Unit conversion ----------
  // Recipes (BOM) usually store per-portion quantities in small units (g, ml), while
  // Products are stocked/ordered in bulk units (kg, l). Without this conversion, raw
  // BOM numbers get summed as if they were already in the product's own order unit,
  // producing wildly inflated results (e.g. grams treated as kilograms).
  const WEIGHT_TO_GRAMS = { g: 1, gr: 1, gramm: 1, kg: 1000 };
  const VOLUME_TO_ML = { ml: 1, cl: 10, dl: 100, l: 1000, liter: 1000, litre: 1000 };

  // Products sold as individual bottles that get packed into boxes. bottleMl = size of
  // one bottle in ml, boxBottles = bottles per box (mirrors Products."Pack Size", kept
  // here explicitly since the script also needs the per-bottle ml, which isn't itself
  // an Airtable field). Add new bottled drinks here as they come up.
  const BOTTLE_PRODUCTS = {
    'rec0GG6FuoGVhppOX': { bottleMl: 750, boxBottles: 6 },  // Cruz De Alba Tinto Roble
    'recr5eRYcQ9DjtF8N': { bottleMl: 750, boxBottles: 12 }, // K Naia Verdejo-Sauvig Blanc
    'recNPOqxkKjbxruPm': { bottleMl: 750, boxBottles: 6 },  // Cava M.Manz.Paloma Minguez
    'recfaXavJa7Mfn95b': { bottleMl: 750, boxBottles: 6 },  // Groenk La Isla Bonita Wine
    'recOur09D12vya1TO': { bottleMl: 750, boxBottles: 6 },  // Paco&Lola Albarino Blanco
    'recFkhzAQHnJQFOv6': { bottleMl: 750, boxBottles: 6 },  // Can Gelat White
    'recqcp3W6BLYP7HLz': { bottleMl: 750, boxBottles: 6 },  // Mucho Más red wine
    'recbmyjGTpXCAgRdf': { bottleMl: 750, boxBottles: 6 },  // Macia Batle 1856 Negre
    'recvWB7SPBW8yCbyw': { bottleMl: 750, boxBottles: 6 },  // Can Gelat Gran Vi Red
    'rectHW7Se1XBPFbF7': { bottleMl: 750, boxBottles: 6 },  // Obalo Rosado
    'recuDpwaPhpDe5Qh0': { bottleMl: 750, boxBottles: 6 },  // Can Gelat Rosé
    'rec5lgdtonSCCuXKn': { bottleMl: 750, boxBottles: 6 },  // Roda I Reserva
    'recBZqhgUHiq9bP1d': { bottleMl: 500, boxBottles: 6 },  // Gin Groenk
    'recYff3VN3LLhBP3R': { bottleMl: 700, boxBottles: 6 },  // Pampelle Aperitif
    'recJVtCM5aKx3fkUs': { bottleMl: 700, boxBottles: 6 },  // Amargero
    'recO3hYm1bTCEbHwE': { bottleMl: 750, boxBottles: 6 },  // André Clouet Brut Gran Reserva
    'recLnrqeL5oKywzpO': { bottleMl: 750, boxBottles: 6 },  // Lanson champagne
    'recrvnmrObLfFgOZO': { bottleMl: 200, boxBottles: 24 }, // Tónica Fever Tree 20cl C-24 (not PK-supplied today, kept for parity with index.html / daily-consumption-deduction.js)
    'recLWTdCvXT0VPUoV': { bottleMl: 700, boxBottles: 12 }, // Taroncello (not PK-supplied today, kept for parity)
  };
  // Products sold from a bulk-liter box (no discrete "bottle" — syrups, bag-in-box). litersPerBox.
  const BULK_LITER_PRODUCTS = {
    'recQuqcE97aPWh9gx': 3,  // Cordial Elderflower
    'recQZGUy989QpkbZe': 3,  // Cordial Strawberry
    'recJmkJvH3IRDwWrc': 3,  // Cordial Ginger
    'recoPQdKRSxB8nKEt': 3,  // Cordial Mango
    'recMaHap0Kptnyc8N': 10, // Vermouth Flors de Collserola
  };
  // Draught products, ordered by the keg. litersPerKeg.
  const KEG_PRODUCTS = {
    'recg4kyyl1P4ionxe': 30, // EG Barril 30L
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
      const ml = VOLUME_TO_ML[f] ? qty * VOLUME_TO_ML[f] : qty * 1000; // assume liters if unit unrecognized
      return ml / 1000 / litersPerBox;
    }

    const litersPerKeg = KEG_PRODUCTS[productId];
    if (litersPerKeg && VOLUME_TO_ML[f]) {
      return (qty * VOLUME_TO_ML[f]) / 1000 / litersPerKeg;
    }

    // Generic weight-based package conversion, read live from each Product's own
    // "Weight per Unit (g)" field — e.g. Panko (10000 g/bag), Harina (1000 g/bag).
    const weightPerUnitG = Number(productWeightPerUnitById[productId]) || null;
    if (WEIGHT_TO_GRAMS[f] && weightPerUnitG) {
      return (qty * WEIGHT_TO_GRAMS[f]) / weightPerUnitG;
    }

    // Generic count-based package conversion, read live from each Product's own "Pack
    // Size" field — e.g. straws (100/pack), eggs (30/carton), carrots (25/tray).
    const startedAsWeightOrVolume = WEIGHT_TO_GRAMS[f] !== undefined || VOLUME_TO_ML[f] !== undefined;
    if (!startedAsWeightOrVolume) {
      const packSize = Number(productPackSizeById[productId]) || null;
      if (packSize) return qty / packSize;
    }

    // Plain count-unit label mismatches ('unit' vs 'Pcs' vs 'db') safely pass through 1:1.
    // But if we started from an actual weight/volume amount and found no confident path
    // to the product's own order unit, don't silently pass the raw number through —
    // signal "unreliable" so the caller can skip it instead of mis-ordering.
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

  function currentStock(productId, locationId) {
    return invTxns
      .filter(t => (t.fields['Related Product'] || []).includes(productId) && (t.fields['Location'] || []).includes(locationId))
      .reduce((sum, t) => {
        const qty = Number(t.fields['Quantity']) || 0;
        return (t.fields['Type'] === 'Waste' || t.fields['Type'] === 'Consumption') ? sum - Math.abs(qty) : sum + qty;
      }, 0);
  }

  // results[restaurantName][productId] = order quantity
  const results = {};
  for (const [restaurantName, locationId] of Object.entries(RESTAURANTS)) {
    // Use the last 2 calendar days that have Daily Sales data for this location — no
    // weekday-matching, since demand is fairly stable day to day here (per Emese).
    const salesForLocation = dailySales.filter(s => (s.fields['Location'] || []).includes(locationId));
    const recentDates = [...new Set(salesForLocation.map(s => s.fields['Date']).filter(Boolean))]
      .sort()
      .reverse()
      .slice(0, 2);
    const recentSales = salesForLocation.filter(s => recentDates.includes(s.fields['Date']));
    console.log(`${restaurantName}: using last-2-days average from ${recentDates.join(', ') || '(no data)'}`);

    // avgUnitsSold[menuItemId] = total units sold across those dates / number of dates
    // (fixed denominator, so a day with zero sales for an item still pulls the average down)
    const salesByMenuItem = {};
    for (const s of recentSales) {
      const miIds = s.fields['Menu Item'] || [];
      const units = Number(s.fields['Units sold']) || 0;
      for (const miId of miIds) {
        (salesByMenuItem[miId] = salesByMenuItem[miId] || []).push(units);
      }
    }
    const avgUnitsSold = {};
    const denom = recentDates.length || 1;
    for (const [miId, arr] of Object.entries(salesByMenuItem)) {
      avgUnitsSold[miId] = arr.reduce((a, b) => a + b, 0) / denom;
    }

    // Aggregate average ingredient consumption per Production-Kitchen product.
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
      // Always round UP, never to nearest — under-ordering a batch-produced item (a whole
      // cake, a kg of dressing) means running out; a small surplus is the safer error.
      const qty = Math.max(0, Math.ceil(par - stock));
      if (qty > 0) restaurantResult[productId] = qty;
    }
    results[restaurantName] = restaurantResult;
    console.log(`${restaurantName}: ${Object.keys(restaurantResult).length} products to order`);
  }

  // ---------- Write Orders + Order Items to Airtable (so they appear in Receive Goods) ----------

  // Guard against duplicate runs (e.g. the scheduled cron firing plus a manual
  // "Run workflow" test on the same day): skip any restaurant that already has an
  // Order from this supplier for this exact target date, regardless of who/what
  // created it (Auto or manual), Pending or Delivered.
  const existingOrders = await airtableGetAll('Orders');
  const alreadyOrdered = new Set(
    existingOrders
      .filter(o => {
        const supplierIds = o.fields['Supplier'] || [];
        const orderDate = (o.fields['Order Date'] || '').slice(0, 10);
        return supplierIds.includes(PRODUCTION_KITCHEN_SUPPLIER_ID) && orderDate === targetDateStr;
      })
      .map(o => o.fields['Restaurant'])
  );

  for (const [restaurantName, productQtys] of Object.entries(results)) {
    if (alreadyOrdered.has(restaurantName)) {
      console.log(`Skipping ${restaurantName}: an order for ${targetDateStr} from Groenk Production Kitchen already exists (idempotency guard).`);
      continue;
    }
    const items = Object.entries(productQtys);
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
  }

  // ---------- Build XLSX ----------

  const productById = Object.fromEntries(products.map(p => [p.id, p.fields]));
  const restaurantNames = Object.keys(RESTAURANTS);
  // Always list every product the Production Kitchen supplies, not just the ones with
  // a nonzero suggestion right now — rows with nothing to order stay in the sheet with
  // blank quantity cells, so the file is always a complete picture, never a partial one.
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
  const isScheduledMorningRun = GITHUB_EVENT_NAME === 'schedule' && currentMadridHour < MORNING_CUTOFF_HOUR;
  const shouldSendEmail = isScheduledMorningRun || SEND_EMAIL_OVERRIDE;

  if (!shouldSendEmail) {
    console.log(
      `Skipping email send (event: "${GITHUB_EVENT_NAME}", Madrid hour: ${currentMadridHour}, override: ${SEND_EMAIL_OVERRIDE}). ` +
      `Email only sends automatically on a scheduled run before ${MORNING_CUTOFF_HOUR}:00 Madrid time, ` +
      `or on a manual run with the send_email input set to true. Airtable writes above (if any) still happened normally.`
    );
    return;
  }

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
      text: hasAnyOrders
        ? `Attached: suggested order for ${targetDateStr} (${offsetLabel}${daysToCover === 2 ? ', covering Sun+Mon since there is no Monday delivery' : ''}), by restaurant and total. The sheet always lists every Production Kitchen product; rows with nothing to order are left blank.`
        : `No items to order for ${targetDateStr} — nothing crossed the buffer threshold. Attached anyway for reference (all rows blank).`,
      // Always attach — the sheet is the full PK product list every day, not just
      // days where something needs ordering.
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
