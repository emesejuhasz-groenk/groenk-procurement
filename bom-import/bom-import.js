/**
 * Grøenk — one-off BOM (Recipes) import from Emese's corrected spreadsheet.
 *
 * Reads an xlsx with columns:
 *   A: Item name on the menu
 *   B: Item name in the daily report from POS   <- matched against Menu Items.Name
 *   D,F,H,...: Ingredient Name.1, .2, .3, ...    <- matched against Products.Name
 *   E,G,I,...: Quantity.1, .2, .3, ...           <- e.g. "40 gr", "1 unit", "1 Pcs"
 *
 * For every (Menu Item, Ingredient) pair found:
 *   - if a Recipes (BOM) record already links that Menu Item + Component (Product),
 *     its Quantity per unit / Unit are UPDATED to match the sheet
 *   - otherwise a new Recipes (BOM) record is CREATED
 *
 * Rows/cells that can't be matched or parsed are NOT written — they're collected
 * and printed at the end so Emese can resolve them by hand. This is a ONE-OFF
 * script (run manually via workflow_dispatch), not part of the daily automation.
 *
 * Required environment variable:
 *   AIRTABLE_TOKEN — same token used for the T+2 script (data.records:read + write)
 *
 * Usage: node bom-import.js path/to/file.xlsx
 */

const XLSX = require('xlsx');

const BASE_ID = 'appPcdy4HEJuDOF4j';
const AIRTABLE_TOKEN = process.env.AIRTABLE_TOKEN;
const MENU_ITEMS_TABLE = 'Menu Items';
const PRODUCTS_TABLE = 'Products';
const BOM_TABLE = 'Recipes (BOM)';

if (!AIRTABLE_TOKEN) {
  console.error('Missing AIRTABLE_TOKEN environment variable.');
  process.exit(1);
}

const filePath = process.argv[2];
if (!filePath) {
  console.error('Usage: node bom-import.js path/to/file.xlsx');
  process.exit(1);
}

// ---------- Airtable helpers (same pattern as production-kitchen-order.js) ----------

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
  // updates: [{ id, fields }]
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

// ---------- Name matching ----------

function norm(s) {
  return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

// Parses a quantity cell like "40 gr", "1 unit", "1 Pcs", "5l every day / restaurant".
// Returns { value, unit } or null if it can't be parsed as a plain number + unit.
function parseQuantity(cell) {
  if (cell === null || cell === undefined) return null;
  const str = String(cell).trim();
  const match = str.match(/^([\d.,]+)\s*(.*)$/);
  if (!match) return null;
  const value = parseFloat(match[1].replace(',', '.'));
  if (Number.isNaN(value)) return null;
  let unit = match[2].trim();
  // Anything with extra words after the unit (e.g. "5l every day / restaurant") is not a
  // simple per-portion quantity — treat as unparseable so it gets flagged, not silently written.
  const unitWordCount = unit.split(/\s+/).filter(Boolean).length;
  if (unitWordCount > 1) return null;
  if (!unit) unit = 'unit';
  // normalize a couple of common variants seen in the sheet
  if (/^gr$/i.test(unit)) unit = 'g';
  if (/^pcs$/i.test(unit)) unit = 'Pcs';
  return { value, unit };
}

async function main() {
  const workbook = XLSX.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const dataRows = rows.slice(1); // skip header row

  console.log(`Read ${dataRows.length} rows from ${filePath}`);

  const [menuItems, products, existingBom] = await Promise.all([
    airtableGetAll(MENU_ITEMS_TABLE),
    airtableGetAll(PRODUCTS_TABLE),
    airtableGetAll(BOM_TABLE),
  ]);

  const menuItemByName = new Map(menuItems.map(m => [norm(m.fields['Name']), m.id]));
  const productByName = new Map(products.map(p => [norm(p.fields['Name']), p.id]));

  // existingBomKey(menuItemId, productId) -> bom record id, for update-vs-create decisions
  const bomByKey = new Map();
  for (const r of existingBom) {
    const miIds = r.fields['Menu Item'] || [];
    const compIds = r.fields['Component (Product)'] || [];
    for (const mi of miIds) for (const comp of compIds) bomByKey.set(`${mi}|${comp}`, r.id);
  }

  const toCreate = [];
  const toUpdate = [];
  const unmatchedMenuItems = new Set();
  const unmatchedIngredients = new Set();
  const unparseableQuantities = [];

  for (const row of dataRows) {
    const posName = row[1]; // "Item name in the daily report from POS"
    if (!posName) continue;
    const menuItemId = menuItemByName.get(norm(posName));
    if (!menuItemId) { unmatchedMenuItems.add(posName); continue; }

    // Ingredient Name.N is at columns 3,5,7,... (0-indexed) / Quantity.N right after it
    for (let col = 3; col < row.length; col += 2) {
      const ingredientName = row[col];
      const quantityCell = row[col + 1];
      if (!ingredientName) continue;

      const productId = productByName.get(norm(ingredientName));
      if (!productId) { unmatchedIngredients.add(ingredientName); continue; }

      const parsed = parseQuantity(quantityCell);
      if (!parsed) {
        unparseableQuantities.push({ menuItem: posName, ingredient: ingredientName, raw: quantityCell });
        continue;
      }

      const key = `${menuItemId}|${productId}`;
      const existingId = bomByKey.get(key);
      const fields = {
        'Quantity per unit': parsed.value,
        'Unit': parsed.unit,
      };
      if (existingId) {
        toUpdate.push({ id: existingId, fields });
      } else {
        toCreate.push({ ...fields, 'Menu Item': [menuItemId], 'Component (Product)': [productId] });
        bomByKey.set(key, 'pending'); // avoid creating duplicates for repeated rows in the same run
      }
    }
  }

  console.log(`\nTo create: ${toCreate.length}`);
  console.log(`To update: ${toUpdate.length}`);

  if (toCreate.length) await airtableCreateMany(BOM_TABLE, toCreate);
  if (toUpdate.length) await airtableUpdateMany(BOM_TABLE, toUpdate);

  console.log('\nDone writing to Airtable.\n');

  if (unmatchedMenuItems.size) {
    console.log(`⚠ ${unmatchedMenuItems.size} menu item name(s) from the sheet had no match in Menu Items:`);
    [...unmatchedMenuItems].forEach(n => console.log('  -', n));
  }
  if (unmatchedIngredients.size) {
    console.log(`\n⚠ ${unmatchedIngredients.size} ingredient name(s) from the sheet had no match in Products:`);
    [...unmatchedIngredients].forEach(n => console.log('  -', n));
  }
  if (unparseableQuantities.length) {
    console.log(`\n⚠ ${unparseableQuantities.length} cell(s) had a quantity that couldn't be parsed as a plain number + unit (left untouched):`);
    unparseableQuantities.forEach(u => console.log(`  - ${u.menuItem} / ${u.ingredient}: "${u.raw}"`));
  }
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
