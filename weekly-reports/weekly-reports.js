/**
 * Grøenk — Weekly Monday reports
 *
 * Runs once a week, Monday morning (see the .yml — scheduled ~06:30 Madrid, before
 * staff arrive, with a workflow_dispatch option for manual re-runs/backfills).
 * Computes everything for the most recently COMPLETED Monday-Sunday week (i.e. if
 * this runs Monday 2026-09-15, "last week" is 2026-09-08 to 2026-09-14) and sends
 * two separate emails:
 *
 *   1. Sales-by-category report -> controlling@groenk.com
 *      A full, ever-growing weekly breakdown of every menu item's sales, grouped
 *      into the DRINK / FOOD category scheme Emese defined (2026-09-13), by
 *      restaurant (Deià / Fornalutx / Sóller Pizza) + a Total column, one 3-column
 *      block per week. Every run recomputes the WHOLE history from
 *      HISTORY_START_DATE through the end of last week and re-sends the complete
 *      file — there's no incremental state to maintain, so a re-run or a gap week
 *      always self-heals.
 *
 *   2. Production Kitchen weekly consumption -> productionkitchengroenk@gmail.com
 *      For last week ONLY: every PK-supplied product's actual total consumption
 *      (units sold x Recipes/BOM, no buffer — this is a historical "what actually
 *      got used" report, not a forecast), by restaurant + Total. This is what the
 *      kitchen actually needs to have produced/delivered last week, so they can
 *      plan this week's production run.
 *
 * This is a SEPARATE, independent process from production-kitchen-order.js (the
 * daily T+0 same-day order, still running every day including Monday) and from
 * daily-consumption-deduction.js (the daily ledger writer). It only reads Daily
 * Sales and Recipes (BOM) — it never writes to Airtable and never touches the
 * Inventory Transactions ledger.
 *
 * Required environment variables (same secrets as the other two scripts):
 *   AIRTABLE_TOKEN, RESEND_API_KEY
 */

const ExcelJS = require('exceljs');

const BASE_ID = 'appPcdy4HEJuDOF4j';
const AIRTABLE_TOKEN = process.env.AIRTABLE_TOKEN;
const RESEND_API_KEY = process.env.RESEND_API_KEY;
const EMAIL_FROM = 'onboarding@resend.dev';
const CONTROLLING_EMAIL = 'controlling@groenk.com';
const PRODUCTION_KITCHEN_EMAIL = 'productionkitchengroenk@gmail.com';
const PRODUCTION_KITCHEN_SUPPLIER_ID = 'recPXErB7VgvkYd6F';

// First real day any Daily Sales data exists for. Doesn't need to be a Monday —
// weeksBetween() below automatically snaps back to that week's own Monday, so the
// very first reported week may include a few days before this date with no data
// (harmless, they'll just show zeros).
const HISTORY_START_DATE = '2026-08-01';

// Daily Sales' "Location" field is a linked record — the real Airtable REST API
// returns an array of record IDs here (not display text; that's only how a CSV
// export or the Airtable UI shows it). Same three location ids used throughout
// production-kitchen-order.js / daily-consumption-deduction.js.
const LOCATION_IDS = {
  'Deià': 'recnoXjgMS7jPYgE7',
  Fornalutx: 'recyfcAwYYZgFzSyd',
  'Soller Pizza': 'reckUG4DXrJTMYtte',
};
const RESTAURANTS = Object.keys(LOCATION_IDS);
const RESTAURANT_DISPLAY = { 'Deià': 'Deià', Fornalutx: 'Fornalutx', 'Soller Pizza': 'Sóller Pizza' };
const LOCATION_NAME_BY_ID = Object.fromEntries(Object.entries(LOCATION_IDS).map(([name, id]) => [id, name]));

if (!AIRTABLE_TOKEN || !RESEND_API_KEY) {
  console.error('Missing required environment variables. Need AIRTABLE_TOKEN, RESEND_API_KEY.');
  process.exit(1);
}

// ---------- Airtable helpers (same pattern as the other two scripts) ----------

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

async function sendResendEmail({ to, subject, text, attachments }) {
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${RESEND_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ from: EMAIL_FROM, to: [to], subject, text, attachments }),
  });
  const result = await res.json();
  if (result.error) throw new Error(`Resend send failed to ${to}: ${result.error.message || JSON.stringify(result.error)}`);
  console.log(`Email sent to ${to} — Resend id:`, result.id || result);
  return result;
}

// ---------- Date / week helpers ----------

function isoDate(d) {
  return d.toISOString().slice(0, 10);
}
function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}
// Monday..Sunday range that most recently completed as of `today`.
function lastCompletedWeek(today) {
  const daysSinceMonday = (today.getDay() + 6) % 7; // 0 if today is itself Monday
  const thisWeekMonday = addDays(today, -daysSinceMonday);
  const lastWeekSunday = addDays(thisWeekMonday, -1);
  const lastWeekMonday = addDays(lastWeekSunday, -6);
  return { start: lastWeekMonday, end: lastWeekSunday };
}
// All Monday-Sunday weeks from `startDate` through `endDate` inclusive (endDate should
// itself be a Sunday — the end of the last completed week).
function weeksBetween(startDate, endDate) {
  const weeks = [];
  let monday = new Date(startDate);
  // Snap startDate back to its own Monday, in case HISTORY_START_DATE isn't one.
  monday = addDays(monday, -((monday.getDay() + 6) % 7));
  while (monday <= endDate) {
    const sunday = addDays(monday, 6);
    weeks.push({ start: new Date(monday), end: sunday > endDate ? new Date(endDate) : sunday, label: `${isoDate(monday)} – ${isoDate(sunday > endDate ? endDate : sunday)}` });
    monday = addDays(monday, 7);
  }
  return weeks;
}

// ---------- Category scheme (confirmed with Emese 2026-09-13) ----------
// Keyed by exact Menu Item name. Renamed/duplicate names (old POS spellings, typos,
// underscore-suffixed dupes fixed on Airtable 2026-09-13) are folded into their
// canonical name so historical weeks before the fix still roll up correctly.
const RENAME_MAP = {
  'Fritas': 'French fries', 'French fries_': 'French fries',
  'Capuccino': 'Cappuccino',
  'Kid menu: chicken': 'Kid menu:chicken', 'Kid menu: pasta': 'Kid menu:pasta',
  'Lasagne Bolognese (urban)': 'Lasagna Bolognese',
  'Aioli con pan y olivas': 'Aioli con pan y olivas (pizzeria)',
  'Ensalada Verde': 'Ensalada Verde (pizzeria)',
  'Aioli sauce': 'Aioli sauce (pizzeria)',
};
// Non-product modifier/voucher rows — excluded from both reports entirely.
const EXCLUDE_ITEMS = new Set([
  '*2 with fries', '+ice', '1medium', '2x hugo', 'VOUCHER', 'a punt', 'Hugo',
  'Elderflower cordial_', 'Strawberry cordial_', 'All-i-oli', 'Tarncello',
]);

const CATEGORY_MAP = {
  WINE: ['1856 / botella', '1856 glass', '2 glass cava', 'ANDRE CLOUET RESERVA CHAMPAGNE', 'Can Gelat BLANCO COPA',
    'Can Gelat Gran Cup red glass', 'Can Gelat Red GRAN VI (premium)', 'Can Gelat Rose Botella', 'Can Gelat blanco BOTELLA',
    'Can Gelat rosado COPA', 'Can Gelat rose glass', 'Can Gelat tinto Gran Cup BOTELLA', 'Can Gelat white glass',
    'Cruz de Alba / botella', 'Cruz de Alba COPA', 'k-naia blanco / botella', 'k-naia blanco COPA', 'Knaia glass',
    'La Isla Bonita BOTELLA', 'La Isla Bonita copa', 'Mucho Mas / botella', 'Mucho Mas COPA', 'Mucho Mas glass',
    'Obalao rosado/COPA', 'Obalo botella', 'Obalo copa', 'Obalo glass', 'Paco & Lola / botella', 'Paco & Lola glass_',
    'Paco Lola COPA', 'Paloma Minguez bottle', 'Paloma Minguez glass', 'Roda 1 reserva botella', 'Rosa blanca', 'Vino dulce copa'],
  Beer: ['Estrella Galicia SIN ALCOHOL', 'Estrella Galicia de barril 330ml', 'Estrella Galicia lager free 0%', 'Shandy 33cl', 'Groenk Beer'],
  'Water/Beverages': ['Limonada', 'Water Large', 'Agua sin gas 700ml', 'Agua con gas 700ml', 'Ginger soda', 'Soda elderflower',
    'Soda strawberry', 'Soda mango', 'Agua 700ml SIN GAS', 'Orange juice', 'Water small', 'Agua 700ml CON GAS', 'Tonic',
    'Aqua 0.5 L sparkling', 'Aqua 0.5 L still', 'Aqua 1 L still', 'Aqua 1L sparkling'],
  'Alcoholic drinks': ['Pampelle spritz', 'Tinto de verano', 'Groenk Gin&Tonic', 'Amargero', 'Hierbesmel Mandragora',
    'Groenk negroni', 'Vermouth', 'Taroncello', 'Evelyn Kirkpatrick', 'Grøenk Gin botella', 'Taroncello Spritz',
    'Suau brandy', 'Vodka Soda', 'Vodka', 'Diplomatico rum', 'Glennfiddich whisky'],
  Coffee: ['Espresso', 'Cortado', 'Cappuccino', 'Double espresso', 'Cafe con leche', 'Americano', 'Te', 'Ice coffee',
    'Latte macchiato', 'Carajillo'],
  Meat: ['Schnietzel PORK IBERICO', 'Beef tenderloin (steak)', '+ César Pollo', 'Iberico Solomillo', 'Sirloin (kg)',
    'Premium Beef Tenderloin', 'Chuleton beef (kg)', 'T-bone (kg)', 'Steak tartare', 'extra chicken', 'extra pulled pork',
    'Grilled chicken (normal portion)'],
  Fish: ['Lubina', 'Pulpo a la parrilla', 'Smoked salmon', 'Tuna (kg)'],
  Hamburger: ['Groenk burger', 'Crispy chicken burger', 'Pulled pork burger'],
  'Pizza + extras': ['Margarita', 'Prosciutto cotto', 'Chorizo', 'Setas', 'Cuatro quesos', 'Sobrasada pizza',
    'extra prosciutto', 'extra parmesan', 'extra mozzarella', 'extra chorizo', 'extra jalapeno', 'Pan de pizza',
    'extra setas / mushroom', 'Pizza sin gluten', 'tomato salsa (no fior di latte)'],
  Pasta: ['Lasagna Bolognese', 'Pasta al pomodoro', 'Pasta al pesto genovese', 'Tomato Pasta', 'Pasta al ragú'],
  'Starters, Salads, Other': ['Aioli with olives & sourdough bread / person', 'Aioli con pan y olivas (pizzeria)',
    'Superfood Salad', 'Rustic salad', 'Spanish style tomato toast', 'Ricotta cream with carrots', 'Bread Basket',
    'Olives', 'French fries', 'Ensalada Verde (pizzeria)', 'Beetroot toast', 'Grilled cauliflower', 'Beans & burrata',
    'Patata Asada', '+ Pan sin gluten', 'Side salad', 'Side salad steak', 'Cebolla caramelizada', 'Avocado toast',
    'Double Cheese toast', 'Scrambled eggs', 'extra scrambled Eggs', 'Potato salad'],
  Desserts: ['Tiramisú', 'Cheesecake', 'Carrot cake', 'Tapioca', 'Panna Cotta', 'Daily dessert', 'Postre del Día'],
  Sauces: ['Aioli', 'Aioli sauce (pizzeria)', 'Mustard', 'Homemade ketchup', 'Mayo relish', 'Jalapeno', 'Bbq sauce'],
  'Kids menu': ['Kid menu:chicken', 'Kid menu:pasta'],
};
const CATEGORY_ORDER = [
  ['DRINK', ['WINE', 'Beer', 'Water/Beverages', 'Alcoholic drinks', 'Coffee']],
  ['FOOD', ['Meat', 'Fish', 'Hamburger', 'Pizza + extras', 'Pasta', 'Starters, Salads, Other', 'Desserts', 'Sauces', 'Kids menu']],
];

function canonicalName(name) {
  return RENAME_MAP[name] || name;
}

// ---------- Report 1: sales by category, by restaurant, by week ----------

async function buildSalesByCategoryReport(dailySales, weeks) {
  // salesByWeekLocItem[weekLabel][location][itemName] = units
  const cube = {};
  for (const w of weeks) cube[w.label] = { 'Deià': {}, Fornalutx: {}, 'Soller Pizza': {} };

  for (const s of dailySales) {
    const f = s.fields;
    const date = f['Date'];
    const location = LOCATION_NAME_BY_ID[(f['Location'] || [])[0]];
    if (!date || !location) continue;
    const week = weeks.find(w => date >= isoDate(w.start) && date <= isoDate(w.end));
    if (!week) continue;
    const name = canonicalName(s.__menuItemName);
    if (!name || EXCLUDE_ITEMS.has(name)) continue;
    const units = Number(f['Units sold']) || 0;
    const bucket = cube[week.label][location] || (cube[week.label][location] = {});
    bucket[name] = (bucket[name] || 0) + units;
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Heti eladás éttermenként');
  const locs = ['Deià', 'Fornalutx', 'Soller Pizza'];
  const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF123F46' } };
  const WEEK_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A5460' } };
  const CAT_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E6B75' } };
  const SUB_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } };
  const SUBTOTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF5F5F0' } };
  const WHITE_BOLD = { name: 'Arial', bold: true, color: { argb: 'FFFFFFFF' }, size: 10 };
  const BOLD = { name: 'Arial', bold: true, size: 10 };
  const NORMAL = { name: 'Arial', size: 9 };

  const ncols = weeks.length * 4; // 3 restaurants + Total per week
  const headerRow1 = ws.getRow(1);
  headerRow1.getCell(1).value = 'Kategória / Alkategória / Tétel';
  let col = 2;
  for (const w of weeks) {
    ws.mergeCells(1, col, 1, col + 3);
    headerRow1.getCell(col).value = w.label;
    col += 4;
  }
  const headerRow2 = ws.getRow(2);
  col = 2;
  for (const w of weeks) {
    for (const label of [...locs.map(l => RESTAURANT_DISPLAY[l]), 'Total']) {
      headerRow2.getCell(col).value = label;
      col++;
    }
  }
  for (let c = 1; c <= ncols + 1; c++) {
    headerRow1.getCell(c).fill = c === 1 ? HEADER_FILL : WEEK_FILL;
    headerRow1.getCell(c).font = WHITE_BOLD;
    headerRow1.getCell(c).alignment = { horizontal: 'center' };
    headerRow2.getCell(c).fill = HEADER_FILL;
    headerRow2.getCell(c).font = WHITE_BOLD;
    headerRow2.getCell(c).alignment = { horizontal: 'center' };
  }
  ws.getColumn(1).width = 42;
  for (let c = 2; c <= ncols + 1; c++) ws.getColumn(c).width = 11;

  for (const [cat, subs] of CATEGORY_ORDER) {
    const catRow = ws.addRow([cat]);
    for (let c = 1; c <= ncols + 1; c++) { catRow.getCell(c).fill = CAT_FILL; catRow.getCell(c).font = WHITE_BOLD; }
    const catSubtotalRowNumbers = [];
    for (const sub of subs) {
      const items = CATEGORY_MAP[sub];
      // Only show items that had at least one sale anywhere in the whole window.
      const itemsWithSales = items.filter(it => weeks.some(w => locs.some(l => (cube[w.label][l] || {})[it])));
      if (!itemsWithSales.length) continue;
      const subRow = ws.addRow([`  ${sub}`]);
      for (let c = 1; c <= ncols + 1; c++) { subRow.getCell(c).fill = SUB_FILL; subRow.getCell(c).font = BOLD; }
      // Sort items by total units across the whole window, descending.
      const totals = itemsWithSales.map(it => {
        let t = 0;
        for (const w of weeks) for (const l of locs) t += (cube[w.label][l] || {})[it] || 0;
        return { it, t };
      }).sort((a, b) => b.t - a.t);
      const itemRowNumbers = [];
      for (const { it } of totals) {
        const row = [it];
        for (const w of weeks) {
          let weekTotal = 0;
          for (const l of locs) {
            const v = (cube[w.label][l] || {})[it] || 0;
            row.push(v);
            weekTotal += v;
          }
          row.push(weekTotal);
        }
        const r = ws.addRow(row);
        for (let c = 1; c <= ncols + 1; c++) r.getCell(c).font = NORMAL;
        itemRowNumbers.push(r.number);
      }
      const subTotalRow = ws.addRow([`  ${sub} — összesen`]);
      for (let c = 2; c <= ncols + 1; c++) {
        const colLetter = ws.getColumn(c).letter;
        subTotalRow.getCell(c).value = { formula: `SUM(${colLetter}${itemRowNumbers[0]}:${colLetter}${itemRowNumbers[itemRowNumbers.length - 1]})` };
        subTotalRow.getCell(c).fill = SUBTOTAL_FILL;
        subTotalRow.getCell(c).font = { name: 'Arial', size: 9, italic: true, bold: true };
      }
      subTotalRow.getCell(1).fill = SUBTOTAL_FILL;
      subTotalRow.getCell(1).font = { name: 'Arial', size: 9, italic: true, bold: true };
      catSubtotalRowNumbers.push(subTotalRow.number);
    }
    const catTotalRow = ws.addRow([`${cat} — MINDÖSSZESEN`]);
    for (let c = 2; c <= ncols + 1; c++) {
      const colLetter = ws.getColumn(c).letter;
      catTotalRow.getCell(c).value = { formula: catSubtotalRowNumbers.map(r => `${colLetter}${r}`).join('+') };
      catTotalRow.getCell(c).fill = CAT_FILL;
      catTotalRow.getCell(c).font = WHITE_BOLD;
    }
    catTotalRow.getCell(1).fill = CAT_FILL;
    catTotalRow.getCell(1).font = WHITE_BOLD;
    ws.addRow([]);
  }
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 2 }];
  return wb;
}

// ---------- Report 2: PK product consumption, last week only, no buffer ----------

async function buildPkConsumptionReport(dailySales, products, recipes, week) {
  const pkProducts = products.filter(p => (p.fields['Supplier'] || []).includes(PRODUCTION_KITCHEN_SUPPLIER_ID));
  const pkProductIds = new Set(pkProducts.map(p => p.id));
  const productById = Object.fromEntries(products.map(p => [p.id, p.fields]));

  const WEIGHT_TO_GRAMS = { g: 1, gr: 1, gramm: 1, kg: 1000 };
  const VOLUME_TO_ML = { ml: 1, cl: 10, dl: 100, l: 1000, liter: 1000, litre: 1000 };
  function convertQty(qty, fromUnit, toUnit, packSize, weightPerUnitG) {
    const f = String(fromUnit || '').toLowerCase().trim();
    const t = String(toUnit || '').toLowerCase().trim();
    if (f === t) return qty;
    if (WEIGHT_TO_GRAMS[f] && WEIGHT_TO_GRAMS[t]) return qty * WEIGHT_TO_GRAMS[f] / WEIGHT_TO_GRAMS[t];
    if (VOLUME_TO_ML[f] && VOLUME_TO_ML[t]) return qty * VOLUME_TO_ML[f] / VOLUME_TO_ML[t];
    if (WEIGHT_TO_GRAMS[f] && weightPerUnitG) return (qty * WEIGHT_TO_GRAMS[f]) / weightPerUnitG;
    const startedAsWeightOrVolume = WEIGHT_TO_GRAMS[f] !== undefined || VOLUME_TO_ML[f] !== undefined;
    if (!startedAsWeightOrVolume && packSize) return qty / packSize;
    return startedAsWeightOrVolume ? null : qty;
  }

  const bomByMenuItem = {};
  for (const r of recipes) {
    const compIds = (r.fields['Component (Product)'] || []).filter(id => pkProductIds.has(id));
    if (!compIds.length) continue;
    const qtyPerUnit = Number(r.fields['Quantity per unit']) || 0;
    const bomUnit = r.fields['Unit'] || 'unit';
    for (const miId of (r.fields['Menu Item'] || [])) {
      for (const productId of compIds) (bomByMenuItem[miId] = bomByMenuItem[miId] || []).push({ productId, qtyPerUnit, bomUnit });
    }
  }

  // result[restaurant][productId] = total units consumed last week
  const result = { 'Deià': {}, Fornalutx: {}, 'Soller Pizza': {} };
  for (const s of dailySales) {
    const f = s.fields;
    const date = f['Date'];
    const location = LOCATION_NAME_BY_ID[(f['Location'] || [])[0]];
    if (!date || !location || date < isoDate(week.start) || date > isoDate(week.end)) continue;
    const bucket = result[location];
    if (!bucket) continue;
    const units = Number(f['Units sold']) || 0;
    if (!units) continue;
    for (const miId of s.__menuItemIds) {
      const bom = bomByMenuItem[miId];
      if (!bom) continue;
      for (const { productId, qtyPerUnit, bomUnit } of bom) {
        const converted = convertQty(qtyPerUnit, bomUnit, productById[productId]['Unit'], productById[productId]['Pack Size'], productById[productId]['Weight per Unit (g)']);
        if (converted === null) continue;
        bucket[productId] = (bucket[productId] || 0) + units * converted;
      }
    }
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('PK heti fogyás');
  ws.columns = [
    { header: 'Product', key: 'product', width: 36 },
    { header: 'Unit', key: 'unit', width: 12 },
    { header: 'Deià', key: 'deia', width: 12 },
    { header: 'Fornalutx', key: 'fornalutx', width: 12 },
    { header: 'Sóller Pizza', key: 'soller', width: 14 },
    { header: 'Total', key: 'total', width: 12 },
  ];
  ws.getRow(1).font = { name: 'Arial', bold: true, size: 14 };

  const CATEGORY_ORDER_PK = ['Meat & Fish', 'Sauce', 'Bakery, pastry, dessert', 'Extra topping', 'Drink', 'Other'];
  const byCat = {};
  for (const pid of pkProductIds) {
    const p = productById[pid] || {};
    const deia = Math.round((result['Deià'][pid] || 0) * 100) / 100;
    const fornalutx = Math.round((result['Fornalutx'][pid] || 0) * 100) / 100;
    const soller = Math.round((result['Soller Pizza'][pid] || 0) * 100) / 100;
    const total = deia + fornalutx + soller;
    if (!total) continue; // nothing consumed anywhere last week — leave off the list
    const cat = p['Order Category'] || 'Other';
    (byCat[cat] = byCat[cat] || []).push({ name: (p['Name'] || '').trim(), unit: p['Unit'] || '', deia, fornalutx, soller, total });
  }
  for (const cat of CATEGORY_ORDER_PK) {
    const rows = byCat[cat];
    if (!rows || !rows.length) continue;
    rows.sort((a, b) => b.total - a.total);
    const headerRow = ws.addRow({ product: cat });
    headerRow.font = { name: 'Arial', bold: true, size: 13 };
    headerRow.eachCell(c => { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8E8E8' } }; });
    for (const row of rows) {
      const r = ws.addRow({ product: row.name, unit: row.unit, deia: row.deia || '', fornalutx: row.fornalutx || '', soller: row.soller || '', total: row.total || '' });
      r.font = { name: 'Arial', size: 12 };
    }
  }
  return wb;
}

// ---------- Main ----------

async function main() {
  const today = new Date();
  const { start: lastWeekStart, end: lastWeekEnd } = lastCompletedWeek(today);
  const weeks = weeksBetween(new Date(HISTORY_START_DATE), lastWeekEnd);
  console.log(`Last completed week: ${isoDate(lastWeekStart)} to ${isoDate(lastWeekEnd)}. History report covers ${weeks.length} week(s) from ${HISTORY_START_DATE}.`);

  const [dailySalesRaw, products, recipes] = await Promise.all([
    airtableGetAll('Daily Sales'),
    airtableGetAll('Products'),
    airtableGetAll('Recipes (BOM)'),
  ]);

  // Daily Sales' "Menu Item" field is a linked-record array of {id, name} in the
  // Airtable REST API when fetched this way is actually just an array of ids — the
  // display name isn't included by default. Resolve names via the Menu Items table.
  const menuItems = await airtableGetAll('Menu Items');
  const menuItemNameById = Object.fromEntries(menuItems.map(m => [m.id, (m.fields['Name'] || '').trim()]));
  for (const s of dailySalesRaw) {
    const ids = s.fields['Menu Item'] || [];
    s.__menuItemIds = ids;
    s.__menuItemName = menuItemNameById[ids[0]] || null;
  }

  const salesWb = await buildSalesByCategoryReport(dailySalesRaw, weeks);
  const salesBuffer = await salesWb.xlsx.writeBuffer();
  await sendResendEmail({
    to: CONTROLLING_EMAIL,
    subject: `Heti eladási riport kategóriánként — ${isoDate(lastWeekStart)} – ${isoDate(lastWeekEnd)} zárva`,
    text: `Csatolva a teljes heti eladási bontás, kategóriánként és éttermenként (Deià / Fornalutx / Sóller Pizza + Total), ${HISTORY_START_DATE}-tól a most lezárt hétig (${isoDate(lastWeekStart)} – ${isoDate(lastWeekEnd)}).`,
    attachments: [{ filename: `heti-eladas-kategoriankent-${isoDate(lastWeekEnd)}.xlsx`, content: Buffer.from(salesBuffer).toString('base64') }],
  });

  const pkWb = await buildPkConsumptionReport(dailySalesRaw, products, recipes, { start: lastWeekStart, end: lastWeekEnd });
  const pkBuffer = await pkWb.xlsx.writeBuffer();
  await sendResendEmail({
    to: PRODUCTION_KITCHEN_EMAIL,
    subject: `Heti termékfogyás (múlt hét: ${isoDate(lastWeekStart)} – ${isoDate(lastWeekEnd)})`,
    text: `Csatolva, mennyi Production Kitchen-es termék fogyott ténylegesen a múlt héten (${isoDate(lastWeekStart)} – ${isoDate(lastWeekEnd)}), éttermenként és összesen — puffer nélkül, ez a tényleges felhasználás, ebből tervezhető a heti gyártás.`,
    attachments: [{ filename: `pk-heti-fogyas-${isoDate(lastWeekEnd)}.xlsx`, content: Buffer.from(pkBuffer).toString('base64') }],
  });
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
