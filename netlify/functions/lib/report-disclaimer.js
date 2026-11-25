"use strict";

// Disclaimer shown at the very bottom of one-off (new-client) reports (PDF export, Word export and the on-screen report).
// The retention sentence depends on metricsRetained, which intune.js sets to true ONLY if the
// month-on-month snapshot was actually saved, so the wording never over- or under-states retention.
// index.html carries its own copy of this text (the browser cannot require this file): keep the two in sync.
// Three paragraphs separated by a blank line. Shown only on one-off reports, which never keep any data,
// so the text does not depend on metricsRetained (the argument is kept so existing callers still work).
function disclaimerText(metricsRetained) {
  return [
    'This report is for information, to help your own evaluation. It is not a full security audit, a compliance certification or a guarantee of the security of your organisation. Integricity Technology Limited ("Integricity") accepts no liability for decisions you make based on this report.',
    'This report is an automated summary of the live data that Microsoft Intune provided through the Microsoft Graph API at the time the report was generated. The reporting tool does not keep any of this data afterwards.',
    'Integricity will securely destroy this document within 30 days. For a full assessment and a remediation plan, contact us to discuss an engagement.',
  ].join('\n\n');
}

module.exports = { disclaimerText };
