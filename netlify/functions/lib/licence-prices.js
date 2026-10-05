"use strict";

// APPROXIMATE list price per seat per month (USD), used ONLY to order the
// Licence assignment list by estimated spend (price x seats purchased). These
// are rough estimates, not tenant billing data, and are never shown in the
// report. Keyed by the friendly licence name produced in intune.js. Unknown
// licences score 0 and so sort after known ones (then by seats assigned).
const EST_PRICE_USD = {
  "Microsoft 365 Business Premium": 22,
  "Microsoft 365 Business Standard": 12.5,
  "Microsoft 365 Business Basic": 6,
  "Microsoft 365 Apps for Business": 8.25,
  "Microsoft 365 E3": 36,
  "Microsoft 365 E5": 57,
  "Microsoft 365 F3": 8,
  "Office 365 E3": 23,
  "Office 365 E1": 10,
  "Exchange Online (Plan 1)": 4,
  "Exchange Online (Plan 2)": 8,
  "Power BI Pro": 14,
  "Microsoft Teams Premium": 10,
  "Microsoft Teams Rooms Pro": 40,
  "Defender And Purview Suites For Business Premium New": 10,
  "Microsoft Defender for Business": 3,
  "Visio Plan 2": 15,
  "Project Plan 3": 30,
  "Microsoft Intune Plan 1": 8,
  "Microsoft Copilot": 30,
  "Microsoft 365 Copilot": 30,
};

// Seats actually bought. Very large "available" numbers are free/trial
// allowances (e.g. 10,000 or 1,000,000), so fall back to seats assigned.
function seatsForSpend(l) {
  return l.available > 0 && l.available < 5000 ? l.available : l.count;
}

function estimatedSpend(l) {
  return (EST_PRICE_USD[l.name] || 0) * seatsForSpend(l);
}

// Paid licences first, ordered by estimated spend (highest first); free after,
// ordered by seats assigned. The original index keeps equal scores stable.
function sortLicences(list) {
  return list
    .map((l, i) => ({ l, i }))
    .sort((a, b) =>
      (a.l.free - b.l.free) ||
      (a.l.free ? (b.l.count - a.l.count) : (estimatedSpend(b.l) - estimatedSpend(a.l) || b.l.count - a.l.count)) ||
      (a.i - b.i))
    .map(x => x.l);
}

module.exports = { EST_PRICE_USD, estimatedSpend, sortLicences };
