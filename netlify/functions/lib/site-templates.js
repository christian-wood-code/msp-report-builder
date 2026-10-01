"use strict";

// Buckets a SharePoint site by the "Root Web Template" value from Microsoft's
// getSharePointSiteUsageDetail report.
//
// Microsoft's docs only say that column is "the template used for creating
// the site" and never list its values, and two naming styles are in use in
// the wild: raw template codes with a version suffix ("GROUP#0", "STS#3",
// "SITEPAGEPUBLISHING#0", "TEAMCHANNEL#1") and friendlier titles ("Group",
// "Communication Site", "Team Site - SharePoint Online configuration"). So
// this matches on prefixes/fragments for BOTH rather than assuming one, and
// intune.js also returns a per-template breakdown (sharepoint.templateBreakdown)
// so the raw values actually seen are visible on the review step rather than
// guessed at.
//
// Returns one of:
//   personal      - OneDrive / My Site personal sites (excluded from counts)
//   group         - created with a Microsoft 365 Group (most Teams)
//   communication - Communication site
//   channel       - the hidden site behind a private/shared Teams channel
//   teams         - modern team site NOT connected to an M365 Group (STS#3)
//   classic       - older-style classic team site (STS#0)
//   other         - tenant system sites (App Catalog, Search Center, My Site
//                   host, admin/hub sites...) or anything unrecognised
function classifySiteTemplate(template) {
  const tl = String(template || "").trim().toLowerCase();

  if (tl.startsWith("msf") || tl.startsWith("personal") || tl.startsWith("spspers") || tl.includes("onedrive")) return "personal";
  if (tl.startsWith("group")) return "group";
  // Spaces vary ("Site Page Publishing" vs "SITEPAGEPUBLISHING#0"), so compare with them removed.
  const compact = tl.replace(/[^a-z0-9#]/g, "");
  if (compact.startsWith("sitepagepublishing") || compact.startsWith("communication")) return "communication";
  // Checked before the team-site rules: a channel site's name also contains "team".
  if (tl.includes("teamchannel") || tl.includes("team channel") || tl.includes("channel site")) return "channel";
  if (tl.startsWith("sts#3") || tl.includes("sharepoint online configuration") || tl.includes("no microsoft 365 group")) return "teams";
  if (tl.startsWith("sts#0") || tl === "team site" || tl.includes("classic")) return "classic";
  return "other";
}

module.exports = { classifySiteTemplate };
