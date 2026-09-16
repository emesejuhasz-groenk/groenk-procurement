/**
 * Grøenk — Daily Sales watchdog
 *
 * ADDED 2026-09-16, after a real incident: Deià's Sept 15 POS data never made it
 * into Airtable (the Cowork import task apparently didn't find/process that day's
 * PDF), and nobody noticed until the day's Production Kitchen order looked
 * strangely empty for meat/fish. Both production-kitchen-order.js and
 * weekly-reports.js now warn inline when THEY detect a gap while doing their own
 * job — but this script exists as a fully independent, third check that doesn't
 * depend on either of them running, or on the Cowork import chain at all. Its
 * only job is: did yesterday's Daily Sales data actually arrive, for all three
 * restaurants? If not, say so loudly. It never writes anything to Airtable.
 *
 * Runs once a day, after the Cowork import should have finished (see the .yml —
 * scheduled for 06:15 Madrid, after the Cowork tasks at 04:30/06:00, and well
 * before the 07:00 Production Kitchen order).
 *
 * Required environment variables: AIRTABLE_TOKEN, RESEND_API_KEY
 */

const BASE_ID = 'appPcdy4HEJuDOF4j';
const AIRTABLE_TOKEN = process.env.AIRTABLE_TOKEN;
const RESEND_API_KEY = process.env.RESEND_API_KEY;
const EMAIL_FROM = 'emese@groenk.com';
const ALERT_EMAIL = 'controlling@groenk.com';

const LOCATION_IDS = {
  'Deià': 'recnoXjgMS7jPYgE7',
  Fornalutx: 'recyfcAwYYZgFzSyd',
  'Soller Pizza': 'reckUG4DXrJTMYtte',
};

if (!AIRTABLE_TOKEN || !RESEND_API_KEY) {
  console.error('Missing required environment variables. Need AIRTABLE_TOKEN, RESEND_API_KEY.');
  process.exit(1);
}

function isoDate(d) {
  return d.toISOString().slice(0, 10);
}
function addDays(date, n) {
  const d = new Date(date);
  d.setDate(d.getDate() + n);
  return d;
}
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Same retry-on-transient-error approach as weekly-reports.js — a single
// paginated fetch, but only ever needs to look at one day's worth of records
// (filtered server-side), so this stays cheap and fast even with retries.
async function airtableGetFiltered(table, filterFormula, { maxRetries = 4 } = {}) {
  const headers = { Authorization: `Bearer ${AIRTABLE_TOKEN}` };
  let records = [];
  let offset;
  do {
    const url = new URL(`https://api.airtable.com/v0/${BASE_ID}/${encodeURIComponent(table)}`);
    url.searchParams.set('pageSize', '100');
    url.searchParams.set('filterByFormula', filterFormula);
    if (offset) url.searchParams.set('offset', offset);

    let data;
    let lastErr;
    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        const res = await fetch(url, { headers });
        const bodyText = await res.text();
        if (!res.ok) {
          lastErr = new Error(`Airtable getFiltered(${table}): HTTP ${res.status} — ${bodyText.slice(0, 300)}`);
        } else {
          try {
            data = JSON.parse(bodyText);
          } catch (parseErr) {
            lastErr = new Error(`Airtable getFiltered(${table}): non-JSON response (status ${res.status}) — ${bodyText.slice(0, 300)}`);
          }
        }
      } catch (networkErr) {
        lastErr = networkErr;
      }
      if (data) break;
      if (attempt < maxRetries) {
        const backoffMs = 1000 * Math.pow(2, attempt);
        console.log(`${table}: request failed (attempt ${attempt + 1}/${maxRetries + 1}), retrying in ${backoffMs}ms — ${lastErr.message}`);
        await sleep(backoffMs);
      }
    }
    if (!data) throw lastErr;
    if (data.error) throw new Error(`Airtable getFiltered(${table}): ${data.error.message}`);
    records = records.concat(data.records);
    offset = data.offset;
  } while (offset);
  return records;
}

async function sendResendEmail({ subject, text }) {
  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: { Authorization: `Bearer ${RESEND_API_KEY}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ from: EMAIL_FROM, to: [ALERT_EMAIL], subject, text }),
  });
  const result = await res.json();
  if (result.error) throw new Error(`Resend send failed: ${result.error.message || JSON.stringify(result.error)}`);
  console.log('Alert email sent — Resend id:', result.id || result);
  return result;
}

async function main() {
  const today = new Date();
  const yesterday = isoDate(addDays(today, -1));

  // Airtable's Location field on Daily Sales links by record id, and
  // filterByFormula can't easily test "does this linked field include id X" —
  // but it CAN test the record id directly against a lookup, so instead we just
  // pull every Daily Sales row for yesterday's date (a small, single-day slice)
  // and check locations against it client-side. Far cheaper than pulling the
  // whole table like the other two scripts do.
  const rows = await airtableGetFiltered('Daily Sales', `{Date} = '${yesterday}'`);

  const missing = [];
  for (const [locName, locId] of Object.entries(LOCATION_IDS)) {
    const hasAny = rows.some(r => (r.fields['Location'] || []).includes(locId));
    if (!hasAny) missing.push(locName);
  }

  if (missing.length === 0) {
    console.log(`OK — all 3 restaurants have Daily Sales data for ${yesterday}.`);
    return;
  }

  console.log(`MISSING for ${yesterday}: ${missing.join(', ')}`);
  await sendResendEmail({
    subject: `⚠️ Daily Sales hiányzik — ${missing.join(', ')} (${yesterday})`,
    text:
      `A(z) ${yesterday}-i napra ${missing.length === 3 ? 'egyik étteremhez sem' : 'a következő étteremhez/étteremekhez nem'} ` +
      `érkezett Daily Sales adat az Airtable-be: ${missing.join(', ')}.\n\n` +
      `Ez azt jelenti, hogy a mai Production Kitchen rendelés és a heti riportok ezekre a helyekre ` +
      `hibás vagy régi adatra fognak visszaesni, amíg ez nincs pótolva.\n\n` +
      `Ellenőrizd: (1) megvan-e a mai/tegnapi POS PDF a Google Drive "HioPOS Reports" mappájában ${missing.join(', ')} ` +
      `helyszín(ek)re, (2) ha megvan, futtasd le kézzel a Cowork daily import task-ot, hogy pótolja.`,
  });
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
