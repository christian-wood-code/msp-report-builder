"use strict";

// Cover disclaimer shown on every report (PDF export, Word export and the on-screen report).
// The retention sentence depends on metricsRetained, which intune.js sets to true ONLY if the
// month-on-month snapshot was actually saved, so the wording never over- or under-states retention.
// index.html carries its own copy of this text (the browser cannot require this file): keep the two in sync.
function disclaimerText(metricsRetained) {
  const retentionText = metricsRetained
    ? 'This report is generated directly from your live Microsoft Graph data at the time of generation. '
      + 'Integricity retains only a small set of summary figures from it (such as device counts and Secure Score - '
      + 'no raw tenant data) for up to 13 months, to show month-on-month changes; this document itself will be '
      + 'securely destroyed within 30 days.'
    : 'This report is generated directly from your live Microsoft Graph data at the time of generation and '
      + 'nothing is retained afterward; this document itself will be securely destroyed within 30 days.';
  return 'This report is a point-in-time, automated summary generated from data available via '
    + 'Microsoft Graph at the time of generation. It is provided as an informational snapshot to support your own '
    + 'evaluation and is not a comprehensive security audit, a compliance certification, or a guarantee of your '
    + "organisation's security posture. Integricity Technology accepts no liability for decisions made on the basis "
    + 'of this report. ' + retentionText
    + ' For a full assessment and remediation plan, contact us to discuss an engagement.';
}

module.exports = { disclaimerText };
