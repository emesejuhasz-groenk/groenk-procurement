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

// ---------- Trigger classification ----------
// CHANGED 2026-09-01: this workflow now has THREE possible triggers (see .yml):
//   - workflow_run   — fired automatically right after Daily Stock Consumption
//                       Deduction completes successfully. This is now the PRIMARY path.
//   - schedule       — a single once-a-day safety-net tick (09:00-10:00 Madrid), only
//                       relevant on the rare day Deduction doesn't trigger at all.
//   - workflow_dispatch — manual "Run workflow" click, for testing/backfilling.
// Both workflow_run and schedule represent genuine, unattended, trustworthy runs that
// should behave identically (readiness check applies, email auto-sends, dedup applies)
// — only workflow_dispatch is the "manual test, stay silent by default" path. Every
// place in this file that used to check GITHUB_EVENT_NAME === 'schedule' now checks
// IS_AUTOMATED_RUN instead, so the switch to workflow_run as primary trigger doesn't
// silently disable the readiness check or the automatic email send.
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

// ---------- Missed-window alert ----------
// CHANGED 2026-09-01: this used to poll on a multi-tick schedule (every 15 min,
// 05:00-05:30 UTC), waiting for Deduction to catch up, and only alerted from the LAST
// tick of that window. That window no longer exists — this workflow is now triggered
// directly by Deduction's completion (workflow_run), plus a single once-a-day
// safety-net schedule tick. Neither path gets multiple attempts in a morning anymore,
// so there's no more "wait for a later tick" to fall back on: every automated run IS
// the only shot for that day. If the readiness check below still isn't satisfied on
// an automated run, that's worth flagging immediately — workflow_run only fires after
// Deduction reports success, so seeing un-deducted data at that point is unusual and
// should surface right away, not get held back waiting for a tick that won't come.

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

// Dedup for the email/Airtable-write step is still based on actual Airtable state (the
// Orders table), not on run history — see the shouldSendEmail / alreadyOrdered check
// near the bottom of main().

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
