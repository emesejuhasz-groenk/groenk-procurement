/**
 * Grøenk — Production Kitchen T+2 auto-order
 *
 * Runs daily (via GitHub Actions cron). For each of the 3 restaurants
 * (Deià, Fornalutx, Sóller Pizza), computes a suggested order quantity for
 * every product supplied by "Groenk Production Kitchen":
 *
 *   avg = historical average consumption for the TARGET WEEKDAY
 *         (Daily Sales × Recipes(BOM), same weekday as target date, all
 *         available history since data collection started)
 *   par = avg * 1.2                      (buffer, per Emese: +20%)
 *   order qty = max(0, round(par - current stock))
 *   current stock = Inventory Transactions ledger sum (Opening Count +
 *                   Delivery Received + Manual Adjustment - Waste)
 *
 * Target date = run date + 2 days (T+2).
 *
 * Output:
 *   1. Creates an Order + Order Items record per restaurant in Airtable
 *      (Supplier = Groenk Production Kitchen) so they show up in the
 *      Procurement app's "Receive goods" screen exactly like a normal order.
 *   2. Sends ONE email to productionkitchengroenk@gmail.com with an XLSX
 *      attachment: rows = products, columns = Deià / Fornalutx / Sóller
 *      Pizza / Total, so the production kitchen chef can prep from one file.
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

const BUFFER_MULTIPLIER = 1.2;
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
  const targetDate = addDays(today, 2);
  const targetDateStr = isoDate(targetDate);
  const targetWeekday = targetDate.getDay(); // 0=Sun..6=Sat
  console.log(`Run date: ${isoDate(today)} — target (T+2) date: ${targetDateStr} (weekday ${targetWeekday})`);

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
  const bomByMenuItem = {};
  for (const r of recipes) {
    const menuItemIds = r.fields['Menu Item'] || [];
    const componentIds = r.fields['Component (Product)'] || [];
    const qtyPerUnit = Number(r.fields['Quantity per unit']) || 0;
    for (const miId of menuItemIds) {
      for (const compId of componentIds) {
        if (!pkProductIds.has(compId)) continue;
        (bomByMenuItem[miId] = bomByMenuItem[miId] || []).push({ productId: compId, qtyPerUnit });
      }
    }
  }

  function currentStock(productId, locationId) {
    return invTxns
      .filter(t => (t.fields['Related Product'] || []).includes(productId) && (t.fields['Location'] || []).includes(locationId))
      .reduce((sum, t) => {
        const qty = Number(t.fields['Quantity']) || 0;
        return t.fields['Type'] === 'Waste' ? sum - Math.abs(qty) : sum + qty;
      }, 0);
  }

  // results[restaurantName][productId] = order quantity
  const results = {};
  for (const [restaurantName, locationId] of Object.entries(RESTAURANTS)) {
    // Historical Daily Sales for this location, matching the target weekday.
    const salesForLocation = dailySales.filter(s => (s.fields['Location'] || []).includes(locationId));
    const matchingWeekdaySales = salesForLocation.filter(s => {
      const d = s.fields['Date'];
      if (!d) return false;
      return new Date(d + 'T00:00:00').getDay() === targetWeekday;
    });

    // avgUnitsSold[menuItemId] = average Units sold across matching-weekday dates
    const salesByMenuItem = {};
    for (const s of matchingWeekdaySales) {
      const miIds = s.fields['Menu Item'] || [];
      const units = Number(s.fields['Units sold']) || 0;
      for (const miId of miIds) {
        (salesByMenuItem[miId] = salesByMenuItem[miId] || []).push(units);
      }
    }
    const avgUnitsSold = {};
    for (const [miId, arr] of Object.entries(salesByMenuItem)) {
      avgUnitsSold[miId] = arr.reduce((a, b) => a + b, 0) / arr.length;
    }

    // Aggregate average ingredient consumption per Production-Kitchen product.
    const avgConsumption = {};
    for (const [miId, avgUnits] of Object.entries(avgUnitsSold)) {
      const bom = bomByMenuItem[miId];
      if (!bom) continue;
      for (const { productId, qtyPerUnit } of bom) {
        avgConsumption[productId] = (avgConsumption[productId] || 0) + avgUnits * qtyPerUnit;
      }
    }

    const restaurantResult = {};
    for (const [productId, avg] of Object.entries(avgConsumption)) {
      const par = avg * BUFFER_MULTIPLIER;
      const stock = currentStock(productId, locationId);
      const qty = Math.max(0, Math.round(par - stock));
      if (qty > 0) restaurantResult[productId] = qty;
    }
    results[restaurantName] = restaurantResult;
    console.log(`${restaurantName}: ${Object.keys(restaurantResult).length} products to order`);
  }

  // ---------- Write Orders + Order Items to Airtable (so they appear in Receive Goods) ----------

  for (const [restaurantName, productQtys] of Object.entries(results)) {
    const items = Object.entries(productQtys);
    if (!items.length) continue;
    const order = await airtableCreate('Orders', {
      'Order Date': targetDateStr,
      'Supplier': [PRODUCTION_KITCHEN_SUPPLIER_ID],
      'Restaurant': restaurantName,
      'Status': 'Pending',
      'Created By': 'Auto (T+2)',
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
  const allProductIds = new Set();
  for (const r of restaurantNames) Object.keys(results[r]).forEach(id => allProductIds.add(id));

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Order');
  sheet.columns = [
    { header: 'Product', key: 'product', width: 32 },
    { header: 'Unit', key: 'unit', width: 12 },
    { header: 'Deià', key: 'deia', width: 10 },
    { header: 'Fornalutx', key: 'fornalutx', width: 10 },
    { header: 'Sóller Pizza', key: 'soller', width: 12 },
    { header: 'Total', key: 'total', width: 10 },
  ];
  sheet.getRow(1).font = { bold: true };

  for (const productId of allProductIds) {
    const p = productById[productId] || {};
    const deia = results['Deia - Groenk Bistro'][productId] || 0;
    const fornalutx = results['Fornalutx - Groenk Bistro'][productId] || 0;
    const soller = results['Soller - Groenk Pizza'][productId] || 0;
    sheet.addRow({
      product: p['Name'] || '(unknown product)',
      unit: p['Unit'] || '',
      deia: deia || '',
      fornalutx: fornalutx || '',
      soller: soller || '',
      total: deia + fornalutx + soller,
    });
  }

  const buffer = await workbook.xlsx.writeBuffer();

  // ---------- Send email via Resend ----------

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
      text: allProductIds.size
        ? `Attached: suggested order for ${targetDateStr} (T+2), by restaurant and total.`
        : `No items to order for ${targetDateStr} — nothing crossed the buffer threshold.`,
      attachments: allProductIds.size ? [{
        filename: `production-kitchen-order-${targetDateStr}.xlsx`,
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
