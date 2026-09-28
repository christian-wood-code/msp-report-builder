"use strict";

// Builds a small, stable metrics snapshot from a full intune.js report
// payload, and turns two snapshots (this period vs last) into deterministic,
// template-based commentary sentences -- "Method A" from the design
// discussion: no AI, no invented claims, every sentence traces directly to a
// number already elsewhere in the report. Kept intentionally simple; if the
// phrasing starts feeling repetitive after a few months in production, the
// natural upgrade is a constrained LLM wording pass over these same facts
// (never a free-form one), not a rewrite of this file's approach.

// Only the numbers needed for commentary -- NOT the full report payload.
// This is the object that gets archived in SharePoint each month.
function snapshot(d) {
  const u = d.users || {};
  const sp = d.sharepoint || {};
  return {
    totalDevices: d.total || 0,
    compliant: d.comp?.compliant || 0,
    noncompliant: d.comp?.noncompliant || 0,
    encrypted: d.encryption?.encrypted || 0,
    notEncrypted: d.encryption?.notEncrypted || 0,
    securePct: d.score?.pct ?? null,
    riskyUsers: d.risky || 0,
    totalUsers: u.total || 0,
    guests: u.guests || 0,
    notSignedIn90: u.notSignedIn90Licensed || 0,
    adminRoles: (u.adminRoles || []).length,
    spSiteCount: sp.siteCount || 0,
    spUsedGB: sp.totalUsedGB || 0,
    spInactiveSites: sp.inactiveSiteCount || 0,
  };
}

function pct(c, p) {
  if (!p) return "n/a";
  const delta = Math.round(((c - p) / p) * 100);
  return `${delta > 0 ? "+" : ""}${delta}%`;
}

// Each rule fires only when `test` says the change is material -- an exact
// value comparison for small counts, a % threshold for larger/noisier ones.
// `text` never introduces a number not present in c/p themselves.
const RULES = [
  {
    key: "totalDevices",
    test: (c, p) => c !== p,
    text: (c, p) => `Total managed devices ${c > p ? "grew" : "fell"} from ${p} to ${c}.`,
  },
  {
    key: "compliant",
    test: (c, p) => p > 0 && Math.abs(c - p) >= 1,
    text: (c, p) => `Compliant devices ${c > p ? "rose" : "fell"} from ${p} to ${c} (${pct(c, p)}).`,
  },
  {
    key: "noncompliant",
    test: (c, p) => c !== p,
    text: (c, p) => c > p
      ? `Non-compliant devices increased from ${p} to ${c} - see the Security Posture section for detail.`
      : `Non-compliant devices decreased from ${p} to ${c}.`,
  },
  {
    key: "securePct",
    test: (c, p) => c != null && p != null && Math.abs(c - p) >= 3,
    text: (c, p) => `Microsoft Secure Score ${c > p ? "rose" : "fell"} from ${p}% to ${c}%.`,
  },
  {
    key: "riskyUsers",
    test: (c, p) => c !== p,
    text: (c, p) => c > p
      ? `${c - p} additional account${c - p > 1 ? "s" : ""} flagged as at-risk this period (${p} to ${c}).`
      : `Risky user accounts decreased from ${p} to ${c}.`,
  },
  {
    key: "totalUsers",
    test: (c, p) => c !== p,
    text: (c, p) => `Licensed user count ${c > p ? "increased" : "decreased"} from ${p} to ${c}.`,
  },
  {
    key: "notSignedIn90",
    test: (c, p) => c !== p,
    text: (c, p) => `Users inactive 90+ days ${c > p ? "increased" : "decreased"} from ${p} to ${c}.`,
  },
  {
    key: "adminRoles",
    test: (c, p) => c !== p,
    text: (c, p) => `Admin role holder count ${c > p ? "increased" : "decreased"} from ${p} to ${c}.`,
  },
  {
    key: "spUsedGB",
    test: (c, p) => p > 0 && Math.abs(c - p) / p >= 0.1,
    text: (c, p) => `SharePoint storage use ${c > p ? "grew" : "shrank"} from ${p} GB to ${c} GB (${pct(c, p)}).`,
  },
];

// Returns null when there's no prior period to compare against (first run
// for this tenant, or the archive lookup failed/timed out) -- callers should
// omit the commentary section entirely in that case, not show an empty one.
function buildCommentary(current, previous) {
  if (!previous) return null;
  const lines = [];
  for (const rule of RULES) {
    const c = current[rule.key], p = previous[rule.key];
    if (c == null || p == null) continue;
    if (rule.test(c, p)) lines.push(rule.text(c, p));
  }
  if (!lines.length) return ["No material change from last month across tracked metrics."];
  return lines;
}

module.exports = { snapshot, buildCommentary };
