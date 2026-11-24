"use strict";

// Disclaimer shown at the very bottom of one-off (new-client) reports (PDF export, Word export and the on-screen report).
// The retention sentence depends on metricsRetained, which intune.js sets to true ONLY if the
// month-on-month snapshot was actually saved, so the wording never over- or under-states retention.
// index.html carries its own copy of this text (the browser cannot require this file): keep the two in sync.
function disclaimerText(metricsRetained) {
  const retention = metricsRetained
    ? 'The app reads your live Microsoft Graph data to make this report. Integricity keeps only a small set of summary figures, for example device counts and Secure Score, for up to 13 months. The app keeps no raw tenant data. Integricity will securely destroy this document within 30 days.'
    : 'The app reads your live Microsoft Graph data to make this report. The app keeps no data after that. Integricity will securely destroy this document within 30 days.';
  return 'This report is an automated summary of the data that Microsoft Graph gives at the time of the report. '
    + 'The report is information to help your own evaluation. '
    + 'It is not a full security audit, a compliance certification or a guarantee of the security of your organisation. '
    + 'Integricity Technology accepts no liability for decisions that you make from this report. '
    + retention
    + ' For a full assessment and a remediation plan, contact us to discuss an engagement.';
}

module.exports = { disclaimerText };
