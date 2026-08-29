/**
 * Grøenk — Production Kitchen auto-order
 *
 * Runs daily (via GitHub Actions cron). For each of the 3 restaurants
 * (Deià, Fornalutx, Sóller Pizza), computes a suggested order quantity for
 * every product supplied by "Groenk Production Kitchen":
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
 * Target date = run date + TARGET_DAY_OFFSET days (see constant below).
 * NORMALLY this is T+1 (order today for tomorrow's delivery), but it is
 * TEMPORARILY set to T+0 (order today for TODAY's delivery) starting
 * 2026-08-27, while stock / auto-order / receiving get back in sync after
 * a double-booking incident. Revert TARGET_DAY_OFFSET to 1 once things
 * have settled — nothing else in this file needs to change to switch back.
 *
 * Production Kitchen now delivers every day of the week (confirmed
 * 2026-08-29 — this replaces the earlier Tuesday–Sunday / no-Monday rule).
 * daysToCover is therefore always 1.
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
const MORNING_CUTOFF_HOUR = 10; // Madrid local time; scheduled runs at/after this hour still send, but the email gets a "this ran late" note appended

function madridHour(date) {
  return parseInt(
    new Intl.DateTimeFormat('en-GB', { timeZone: 'Europe/Madrid', hour: '2-digit', hour12: false }).format(date),
    10
  );
}

// ---------- Missed-window alert ----------
// The workflow ticks every 15 min, 05:00-05:30 UTC. If the readiness check below never
// passes all morning (the Cowork imports or the deduction job never caught up), every
// tick was silently exiting with just a console log — nobody would know the day's order
// was skipped until someone noticed the missing email. This sends ONE alert email
// instead, but only from the LAST tick of the window, so it fires at most once per day
// and only once we're sure no earlier tick is still going to succeed.
const LAST_TICK_UTC_HOUR = 5;
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

// The workflow now ticks every 15 minutes through the morning window (see the .yml),
// so this script runs many times per day. Two things stop that from causing spam:
// (1) the readiness check below (skips entirely until the deduction job has caught up),
// (2) the Orders-based dedup check right before the email send (skips once a real send
// already happened today). Both are based on actual Airtable state, not on run history —
// every GitHub Actions tick "succeeds" whether or not it did anything, so run history
// alone can't distinguish "not ready yet" from "already sent".

const BUFFER_MULTIPLIER = 1.5;

// TEMPORARY (a few weeks, starting 2026-08-27): order for TODAY (T) instead of
// tomorrow (T+1), while stock / auto-order / receiving get back in sync after the
// double-booking incident. Revert to 1 once things have settled.
// NOTE: this shifts EVERY downstream date calculation (target date, weekday) by the
// same amount — no other code changes needed elsewhere in this file.
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

// ---------- Main ----------

async function main() {
  const today = new Date();
  const targetDate = addDays(today, TARGET_DAY_OFFSET);
  const targetDateStr = isoDate(targetDate);
  const targetWeekday = targetDate.getDay(); // 0=Sun..6=Sat
  console.log(`Run date: ${isoDate(today)} — target (T+${TARGET_DAY_OFFSET}) date: ${targetDateStr} (weekday ${targetWeekday})`);

  // Production Kitchen now delivers every day of the week (confirmed 2026-08-29 — no
  // more Monday exception, no more Sunday-covers-2-days special case). daysToCover is
  // kept as a variable (rather than inlining "1" everywhere) so it's a single, obvious
  // place to reintroduce a multi-day rule later if the delivery schedule ever changes
  // again.
  const daysToCover = 1;
  console.log(`Days to cover: ${daysToCover}`);

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
  // Inventory Transactions. Rather than assuming a fixed time gap (which broke over
  // the DST change, and whenever Actions itself is delayed), check the actual data:
  // is the most recent day's eligible Daily Sales batch fully marked "Stock Deducted"?
  // If not, this run is too early — bail out quietly and let a later cron tick
  // (the workflow now runs every 15 min in the morning window) pick it up once ready.
  // Manual runs skip this check so testing/backfilling is never blocked by it.
  if (GITHUB_EVENT_NAME === 'schedule') {
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
      if (isLastTickOfWindow(today)) {
        console.log(`Last tick of the morning window and still not ready (${reason}). Sending a missed-window alert.`);
        try {
          await sendResendEmail({
            subject: `⚠️ Production Kitchen order MISSED for ${targetDateStr}`,
            text:
              `No Production Kitchen order was generated or sent today — the morning window (05:00–05:30 UTC / 07:00–07:30 Madrid) ` +
              `ended without the Daily Stock Consumption Deduction catching up on today's Daily Sales data.\n\n` +
              `Details: ${reason}.\n\n` +
              `Likely cause: the Cowork POS import task(s) (04:30 / 06:00) didn't complete, or the Deduction ` +
              `workflow failed. Check the Cowork Scheduled tasks page and the Airtable Daily Sales table, then ` +
              `run "Production Kitchen T+1 auto-order" manually (Run workflow, with send_email checked) once ` +
              `the data looks right.`,
          });
        } catch (e) {
          console.log(`Missed-window alert email also failed to send: ${e.message}`);
        }
      } else {
        console.log(`Daily Stock Consumption Deduction hasn't caught up yet (${reason}). Skipping this tick — a later scheduled run will pick it up once the data is ready.`);
      }
      return;
    }
  }

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

  // ---------- Current stock (physical-count-aware) ----------
  // Fixed 2026-08-29: a plain chronological sum of every ledger row double-counts
  // consumption whenever a physical count gets entered LATE (its own day's
  // consumption not yet deducted at count time, then backfilled afterward). The
  // physical count already reflects that consumption in the real world — so
  // re-applying a backfilled row dated BEFORE the count subtracts it twice.
  //
  // The fix needs two different orderings for two different questions:
  // 1) What was the count's delta CALIBRATED AGAINST? Whatever existed in Airtable
  //    at the moment it was entered — i.e. sorted by createdTime (insertion order),
  //    not by the ledger rows' own Date. history-by-insertion-order + the count's
  //    own delta = the actual physical count value that was typed in.
  // 2) Which LATER rows should still apply on top? Only ones whose real-world Date
  //    is on/after the count's Date — a row inserted afterward (backfilled) but
  //    dated BEFORE the count already happened before the shelf was counted, so is
  //    already baked into that physical number and must be excluded, not re-summed.
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
    // Use YESTERDAY'S actual sales only — the single most recent calendar day that has
    // Daily Sales data for this location. Confirmed 2026-08-29 (Emese): this is NOT an
    // average of several days anymore. Rule: required stock for today = yesterday's
    // units sold × 1.5, minus whatever is still actually on the shelf right now.
    // Example: sold 40 yesterday, had 60 in stock yesterday morning → 20 left tonight.
    // Required stock today = 40 × 1.5 = 60. Order = 60 − 20 = 40.
    const salesForLocation = dailySales.filter(s => (s.fields['Location'] || []).includes(locationId));
    const recentDates = [...new Set(salesForLocation.map(s => s.fields['Date']).filter(Boolean))]
      .sort()
      .reverse()
      .slice(0, 1);
    const recentSales = salesForLocation.filter(s => recentDates.includes(s.fields['Date']));
    console.log(`${restaurantName}: using yesterday's sales from ${recentDates[0] || '(no data)'}`);

    // avgUnitsSold[menuItemId] = units sold on that single most recent day (no averaging).
    // Kept the name "avgUnitsSold" / "avg" below for minimal diff — it's really just
    // "yesterday's units", not an average, as of the 2026-08-29 formula change.
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

  // Keep the Order/Order Items in sync with the freshly computed quantities every
  // time this script runs — the emailed xlsx and what staff see in "Receive Goods"
  // must always match (confirmed 2026-08-29, after a stale-order incident where the
  // email got a formula fix but the Airtable Order Items were left with the old,
  // wrong numbers because the old code just skipped writing once an Order existed).
  //
  // Safety guard: if ANY item on an existing Order is already marked Received (goods
  // receiving has started/finished for it), don't touch that Order's items at all —
  // reconciling quantities mid-delivery could corrupt an in-progress goods receipt.
  // Log it clearly so a genuine same-day recompute after receiving starts is visible
  // rather than silently skipped.
  const existingOrders = await airtableGetAll('Orders');
  const existingOrderItems = await airtableGetAll('Order Items');
  const ordersByRestaurant = {}; // restaurantName -> order record
  for (const o of existingOrders) {
    const supplierIds = o.fields['Supplier'] || [];
    const orderDate = (o.fields['Order Date'] || '').slice(0, 10);
    if (supplierIds.includes(PRODUCTION_KITCHEN_SUPPLIER_ID) && orderDate === targetDateStr) {
      ordersByRestaurant[o.fields['Restaurant']] = o;
    }
  }
  const alreadyOrdered = new Set(Object.keys(ordersByRestaurant));

  for (const [restaurantName, productQtys] of Object.entries(results)) {
    const items = Object.entries(productQtys); // [productId, qty][]
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

  // IMPORTANT: a genuine scheduled run ALWAYS sends the email, no matter how late in the
  // day it actually executes — GitHub Actions can delay scheduled runs by hours under load
  // (this happened on 2026-08-27: the 07:00 cron didn't fire until 18:19). Silently skipping
  // a late-but-real scheduled run would mean the kitchen gets NO order summary that day at
  // all, which is worse than one arriving late. We only add a "this is late" note to the
  // email itself so it's clear what happened, instead of leaving anyone to guess.
  //
  // Manual test runs (workflow_dispatch) stay silent by default regardless of time of day —
  // that's what was actually causing the confusing extra emails (two manual test runs this
  // morning, in addition to the delayed real one). Use the send_email input to opt in when
  // you deliberately want to test the email itself.
  const currentMadridHour = madridHour(today);
  const isLateRun = currentMadridHour >= MORNING_CUTOFF_HOUR;
  const shouldSendEmail = GITHUB_EVENT_NAME === 'schedule' || SEND_EMAIL_OVERRIDE;

  if (!shouldSendEmail) {
    console.log(
      `Skipping email send (event: "${GITHUB_EVENT_NAME}", override: ${SEND_EMAIL_OVERRIDE}). ` +
      `Email only sends automatically on a scheduled run, or on a manual run with the ` +
      `send_email input set to true. Airtable writes above (if any) still happened normally.`
    );
    return;
  }

  // With the workflow now ticking every 15 min through the morning window, this same
  // check needs to also stop a SECOND scheduled tick (after the real one already sent)
  // from re-sending. Base this on the Orders idempotency check above.
  //
  // IMPORTANT: only apply this to SCHEDULED ticks, never to a manual run with
  // send_email checked. An explicit manual "actually send" click is a direct human
  // command for THIS run — it must not be silently swallowed just because an earlier
  // run (e.g. an unchecked silent test) already created the Airtable Order records
  // without ever emailing anyone. Conflating "Orders exist" with "email was sent" was
  // exactly the bug that caused a checked, explicit send to go silent on 2026-08-29.
  if (GITHUB_EVENT_NAME === 'schedule') {
    const restaurantsNeedingOrder = Object.entries(results).filter(([, q]) => Object.keys(q).length > 0).map(([r]) => r);
    if (restaurantsNeedingOrder.length > 0 && restaurantsNeedingOrder.every(r => alreadyOrdered.has(r))) {
      console.log('All restaurants that need an order already have one for today — this looks like a later scheduled tick catching up after a real send already went out. Skipping to avoid a duplicate email.');
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
