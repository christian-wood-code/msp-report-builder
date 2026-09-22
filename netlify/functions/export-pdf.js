"use strict";

// PDF export (server-side) -- full mirror of export-docx.js's section
// structure/content (Cover, Executive Summary, Device & Asset Management,
// Security Posture, Patch Status, User Data, SharePoint/MS Teams,
// Ticketing & Support, Recommendations, Security Risk Register), not the
// earlier condensed first pass. Same {client, from, to, preparer, today,
// iData, manual, fullLogoTransparent} input shape as export-docx.js.
//
// No Inter font file exists in this project (unlike the sibling apps) --
// Helvetica throughout, consistent with jsPDF's own "fails open to
// Helvetica" pattern used elsewhere when Inter isn't available.

const { jsPDF } = require('jspdf');
require('jspdf-autotable');

const CORS = { 'Access-Control-Allow-Origin': '*' };

// ── Brand colours -- mirrors export-docx.js's C palette, as RGB arrays ──────
const C = {
  blue: [13, 124, 196], purple: [140, 12, 110], red: [230, 27, 31], orange: [224, 100, 0], teal: [26, 128, 112],
  dark: [31, 41, 55], gray: [107, 114, 128], lgray: [156, 163, 175],
  bgray: [248, 249, 252], border: [229, 231, 235], white: [255, 255, 255],
  success: [15, 122, 60], sucbg: [220, 252, 231],
  warn: [217, 119, 6], warnBg: [255, 247, 237], warnTxt: [146, 64, 14], warnPill: [194, 65, 12],
  err: [192, 21, 42], errBg: [254, 226, 226], errTxt: [153, 27, 27],
  navy: [15, 23, 42],
};
const KPI_COLOR = { good: C.success, warn: C.warn, bad: C.err, info: C.blue, neu: C.dark };
const CALLOUT_CFG = {
  warn: { bg: C.warnBg, accent: C.warn, text: C.warnTxt },
  bad: { bg: C.errBg, accent: C.err, text: C.errTxt },
  good: { bg: C.sucbg, accent: C.success, text: [20, 83, 45] },
  info: { bg: [239, 246, 255], accent: C.blue, text: [30, 58, 138] },
};
const SEVERITY_RGB = { high: C.err, medium: C.warnPill, low: C.success };
const SEVERITY_BG = { high: C.errBg, medium: C.warnBg, low: C.sucbg };
const PILL_CFG = {
  enforced: { bg: C.sucbg, tc: [22, 101, 52], label: 'Enforced' },
  'report-only': { bg: C.warnBg, tc: C.warnTxt, label: 'Report-only' },
  'not found': { bg: C.errBg, tc: C.errTxt, label: 'Not found' },
  disabled: { bg: [243, 244, 246], tc: C.gray, label: 'Disabled' },
  enabled: { bg: C.sucbg, tc: [22, 101, 52], label: 'Enabled' },
  'enabled-warn': { bg: C.warnBg, tc: C.warnTxt, label: 'Enabled - review' },
};

const k = (value, label, cls) => ({ value, label, cls });

function buildPdfDoc({ client, from, to, preparer, today, iData: d, manual = {}, fullLogoTransparent }) {
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });
  const PW = doc.internal.pageSize.getWidth();
  const PH = doc.internal.pageSize.getHeight();
  const M = 40;
  const CW = PW - 2 * M;
  let y = 0;

  const setFill = c => doc.setFillColor(...c);
  const setText = c => doc.setTextColor(...c);
  const setDraw = c => doc.setDrawColor(...c);
  const ensure = need => { if (y + need > PH - M - 24) { doc.addPage(); y = M; } };
  const fmtDate = s => s ? new Date(s).toLocaleDateString('en-NZ', { day: 'numeric', month: 'long', year: 'numeric' }) : '';

  // ── Section header: colour accent bar + bold title + "THIS MONTH" badge,
  // matching sectionHeaderBlock()'s visual in export-docx.js ────────────────
  const MIN_FOLLOW = 90;
  const sectionHeader = (title, color, introText) => {
    ensure(50 + MIN_FOLLOW);
    y += 14;
    setFill(color); doc.rect(M, y - 12, 4, 22, 'F');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(15); setText(C.dark);
    doc.text(title, M + 14, y + 3);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8); setText(C.lgray);
    doc.text('THIS MONTH', PW - M, y + 2, { align: 'right' });
    y += 14;
    setDraw(C.border); doc.setLineWidth(0.7); doc.line(M, y, PW - M, y);
    y += 18;
    if (introText) {
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9); setText([55, 65, 81]);
      const lines = doc.splitTextToSize(introText, CW);
      doc.text(lines, M, y);
      y += lines.length * 11 + 12;
    }
  };

  const subLabel = text => {
    ensure(30);
    setDraw(C.border); doc.setLineWidth(2); doc.line(M, y - 8, M, y + 4);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(9); setText(C.gray);
    doc.text(text.toUpperCase(), M + 10, y);
    y += 16;
  };

  // ── KPI grid -- chunks into rows of up to 4 stat cards ──────────────────
  const kpiGrid = items => {
    const perRow = 4;
    for (let i = 0; i < items.length; i += perRow) {
      const chunk = items.slice(i, i + perRow);
      const gap = 10, n = chunk.length;
      const w = (CW - gap * (n - 1)) / n, h = 56;
      ensure(h + 14);
      chunk.forEach((it, j) => {
        const x = M + j * (w + gap);
        const colour = KPI_COLOR[it.cls] || C.dark;
        setFill(C.white); doc.roundedRect(x, y, w, h, 5, 5, 'F');
        setDraw(C.border); doc.setLineWidth(0.7); doc.roundedRect(x, y, w, h, 5, 5, 'S');
        setFill(colour); doc.roundedRect(x, y, w, 3, 1.5, 1.5, 'F');
        setText(colour); doc.setFont('helvetica', 'bold'); doc.setFontSize(17);
        doc.text(String(it.value), x + w / 2, y + 28, { align: 'center', maxWidth: w - 10 });
        setText(C.gray); doc.setFont('helvetica', 'bold'); doc.setFontSize(7);
        doc.text(it.label.toUpperCase(), x + w / 2, y + 44, { align: 'center', maxWidth: w - 12 });
      });
      y += h + 14;
    }
    y += 4;
  };

  // ── Callout box -- coloured left accent bar + tinted background ────────
  const callout = (text, type) => {
    const cfg = CALLOUT_CFG[type] || { bg: C.bgray, accent: C.border, text: C.dark };
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5);
    const lines = doc.splitTextToSize(text, CW - 26);
    const h = Math.max(28, lines.length * 12 + 14);
    ensure(h + 12);
    setFill(cfg.bg); doc.rect(M, y, CW, h, 'F');
    setFill(cfg.accent); doc.rect(M, y, 3, h, 'F');
    setText(cfg.text); doc.text(lines, M + 14, y + 16);
    y += h + 14;
  };

  const tableStyles = {
    styles: { font: 'helvetica', fontSize: 8.5, cellPadding: 5, lineColor: C.border, lineWidth: 0.5, textColor: [55, 65, 81] },
    headStyles: { font: 'helvetica', fillColor: C.dark, textColor: 255, fontStyle: 'bold', halign: 'left' },
    alternateRowStyles: { fillColor: C.bgray },
  };
  const dataTable = (head, rows, columnStyles) => {
    if (!rows.length) return;
    ensure(45);
    doc.autoTable({ startY: y, margin: { left: M, right: M }, head: [head], body: rows, theme: 'grid', columnStyles: columnStyles || {}, ...tableStyles });
    y = doc.lastAutoTable.finalY + 16;
  };

  // ── Policy/status pill table -- name left, coloured status pill right ──
  const policyTable = rows => {
    ensure(45);
    doc.autoTable({
      startY: y, margin: { left: M, right: M },
      head: [['Policy', 'Status']], body: rows.map(r => [r.label, (PILL_CFG[r.status] || { label: r.status }).label]),
      theme: 'grid', columnStyles: { 1: { halign: 'center', cellWidth: 130 } },
      ...tableStyles,
      didParseCell: data => {
        if (data.section === 'body' && data.column.index === 1) {
          const status = rows[data.row.index].status;
          const pc = PILL_CFG[status] || { bg: [243, 244, 246], tc: C.gray };
          data.cell.styles.textColor = pc.tc;
          data.cell.styles.fillColor = pc.bg;
          data.cell.styles.fontStyle = 'bold';
        }
      },
    });
    y = doc.lastAutoTable.finalY + 16;
  };

  // ── Licence cards -- 3-per-row ───────────────────────────────────────────
  const licenceCards = cards => {
    for (let i = 0; i < cards.length; i += 3) {
      const row = cards.slice(i, i + 3);
      const gap = 10, n = 3;
      const w = (CW - gap * (n - 1)) / n, h = 54;
      ensure(h + 14);
      row.forEach((l, j) => {
        const x = M + j * (w + gap);
        setFill(C.bgray); doc.roundedRect(x, y, w, h, 4, 4, 'F');
        setDraw(C.border); doc.setLineWidth(0.7); doc.roundedRect(x, y, w, h, 4, 4, 'S');
        doc.setFont('helvetica', 'normal'); doc.setFontSize(7.5); setText(C.gray);
        doc.text(l.name, x + 10, y + 14, { maxWidth: w - 20 });
        const label = l.available > 0 ? `${l.count} / ${l.available}` : String(l.count);
        const overCap = l.available > 0 && l.count > l.available;
        doc.setFont('helvetica', 'bold'); doc.setFontSize(16); setText(overCap ? C.err : C.dark);
        doc.text(label, x + 10, y + 34);
        if (l.available > 0) {
          doc.setFont('helvetica', 'normal'); doc.setFontSize(6.5); setText(C.lgray);
          doc.text('licences used', x + 10, y + 46);
        }
      });
      y += h + 12;
    }
    y += 4;
  };

  // ── Risk register block -- severity-coloured left bar, matches the
  // condensed version's already-verified layout ───────────────────────────
  const riskRegister = risks => {
    risks.forEach(r => {
      const descLines = doc.splitTextToSize(r.finding, CW - 10);
      const actLines = doc.splitTextToSize(r.action, CW - 10);
      const h = 20 + descLines.length * 11 + actLines.length * 11 + 16;
      ensure(h);
      const accent = SEVERITY_RGB[r.severity] || C.dark;
      setFill(SEVERITY_BG[r.severity] || C.bgray); doc.rect(M, y - 4, CW, h, 'F');
      setFill(accent); doc.rect(M, y - 4, 3, h, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8); setText(accent);
      doc.text(r.severity.toUpperCase(), M + 12, y + 8);
      setText(C.dark); doc.setFontSize(10.5);
      doc.text(r.area, M + 12 + doc.getTextWidth(r.severity.toUpperCase()) + 10, y + 8);
      y += 20;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9); setText([55, 65, 81]);
      doc.text(descLines, M + 12, y); y += descLines.length * 11 + 6;
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); setText(C.blue);
      doc.text('Action: ', M + 12, y);
      doc.setFont('helvetica', 'normal'); setText([55, 65, 81]);
      doc.text(actLines, M + 12 + doc.getTextWidth('Action: '), y);
      y += Math.max(actLines.length, 1) * 11 + 16;
    });
  };

  // ══════════════════════════════ Cover ═══════════════════════════════════
  const BANDH = 118;
  setFill(C.navy); doc.rect(0, 0, PW, BANDH, 'F');
  const coverLogo = Buffer.from(fullLogoTransparent || '', 'base64');
  if (coverLogo.length > 100) {
    const logoH = 26, logoW = logoH * (960 / 247);
    try { doc.addImage(coverLogo, 'PNG', M, 20, logoW, logoH, undefined, 'SLOW'); } catch (_) { /* ignore malformed logo */ }
  }
  doc.setFont('helvetica', 'bold'); setText(C.white); doc.setFontSize(24);
  doc.text(client, M, 76);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(11); setText([148, 163, 184]);
  doc.text(new Date(from).toLocaleDateString('en-NZ', { month: 'long', year: 'numeric' }), M, 96);
  doc.setFontSize(9);
  doc.text(`Account manager: ${preparer || 'Integricity Technology'}    Generated: ${today}`, M, 110);
  y = BANDH + 22;

  // Disclaimer -- sits directly under the cover band on every report so a
  // reader (client or prospect) sees it before anything else. See
  // msp-report-builder's docs for the full-length version shown to
  // prospects before they submit credentials.
  const disclaimerText = 'This report is a point-in-time, automated summary generated from data available via '
    + 'Microsoft Graph at the time of generation. It is provided as an informational snapshot to support your own '
    + 'evaluation and is not a comprehensive security audit, a compliance certification, or a guarantee of your '
    + "organisation's security posture. Integricity Technology accepts no liability for decisions made on the basis "
    + 'of this report. This report is generated directly from your live Microsoft Graph data at the time of '
    + 'generation and nothing is retained afterward; this document itself will be securely destroyed within 30 '
    + 'days. For a full assessment and remediation plan, contact us to discuss an engagement.';
  doc.setFont('helvetica', 'italic'); doc.setFontSize(7.5); setText(C.lgray);
  const discLines = doc.splitTextToSize(disclaimerText, CW);
  doc.text(discLines, M, y);
  y += discLines.length * 9 + 14;
  setDraw(C.border); doc.setLineWidth(0.7); doc.line(M, y - 6, PW - M, y - 6);
  y += 6;

  // ══════════════════════════ Executive Summary ═══════════════════════════
  if (manual.overview || manual.highlights || manual.concerns || manual.projects) {
    sectionHeader('Executive Summary', C.blue);
    if (manual.overview) { doc.setFont('helvetica', 'normal'); doc.setFontSize(10); setText(C.dark); const lines = doc.splitTextToSize(manual.overview, CW); ensure(lines.length * 12 + 10); doc.text(lines, M, y); y += lines.length * 12 + 12; }
    if (manual.highlights) callout('Highlights: ' + manual.highlights, 'good');
    if (manual.concerns) callout('Concerns: ' + manual.concerns, 'warn');
    if (manual.projects) callout('Projects: ' + manual.projects, 'info');
  }

  // ══════════════════════════ Device & Asset Management ═══════════════════
  sectionHeader('Device & Asset Management', C.blue,
    'Summary of managed devices enrolled in Microsoft Intune, including compliance status, operating system breakdown, and device health indicators.');
  kpiGrid([
    k(d.total ?? 0, 'Total devices', 'info'),
    k(d.comp?.compliant ?? 0, 'Compliant', 'good'),
    k(d.comp?.noncompliant ?? 0, 'Non-Compliant', (d.comp?.noncompliant ?? 0) > 0 ? 'bad' : 'good'),
    k(d.comp?.unknown ?? 0, 'Not evaluated', 'warn'),
    k(d.staleCount ?? 0, 'Last Checked In 90d+', (d.staleCount ?? 0) > 0 ? 'warn' : 'good'),
  ]);

  const osKpis = [
    ...(d.win25h2 > 0 ? [k(d.win25h2, 'Windows 11 25H2', 'good')] : []),
    ...(d.win24h2 > 0 ? [k(d.win24h2, 'Windows 11 24H2', 'good')] : []),
    ...(d.win11 > 0 ? [k(d.win11, 'Windows 11', 'good')] : []),
    ...(d.win10 > 0 ? [k(d.win10, 'Windows 10', 'warn')] : []),
    ...(d.macOS > 0 ? [k(d.macOS, 'macOS', 'neu')] : []),
    ...(d.linux > 0 ? [k(d.linux, 'Linux', 'neu')] : []),
    ...(d.iosCount > 0 ? [k(d.iosCount, 'iOS / iPadOS', 'neu')] : []),
    ...(d.androidCount > 0 ? [k(d.androidCount, 'Android', 'neu')] : []),
  ];
  if (osKpis.length) { subLabel('Operating system'); kpiGrid(osKpis); }

  if ((d.comp?.noncompliant ?? 0) > 0) {
    callout(`${d.comp.noncompliant} device${d.comp.noncompliant > 1 ? 's are' : ' is'} non-compliant with your organisation's policies.`, 'bad');
    if ((d.notCompliantList || []).length) {
      subLabel('Non-compliant devices');
      dataTable(['Device', 'Primary user', 'Last seen'],
        d.notCompliantList.map(dev => [dev.name || '', dev.user || 'Unknown', dev.lastSync ? new Date(dev.lastSync).toLocaleDateString('en-NZ') : 'Never']));
    }
  }

  subLabel('Encryption');
  kpiGrid([
    k(d.encryption?.encrypted ?? 0, 'Encrypted', 'good'),
    k(d.encryption?.notEncrypted ?? 0, 'Not encrypted', (d.encryption?.notEncrypted ?? 0) > 0 ? 'bad' : 'good'),
  ]);
  if ((d.encryption?.notEncrypted ?? 0) > 0) {
    callout(`${d.encryption.notEncrypted} device${d.encryption.notEncrypted > 1 ? 's are' : ' is'} not encrypted. This is a significant data protection risk.`, 'bad');
    if ((d.notEncryptedList || []).length) {
      subLabel('Unencrypted devices');
      dataTable(['Device', 'Primary user', 'OS'], d.notEncryptedList.map(dev => [dev.name || '', dev.user || 'Unknown', dev.os || '']));
    }
  }
  if ((d.lowDisk || []).length) {
    callout(`${d.lowDisk.length} device${d.lowDisk.length > 1 ? 's have' : ' has'} low disk space (less than 15% free).`, 'warn');
    subLabel('Devices with low disk space');
    dataTable(['Device', 'Free space', 'Free (GB)'], d.lowDisk.map(x => [x.name || '', `${x.pct}%`, `${x.gb} GB`]));
  }

  // ══════════════════════════ Security Posture ═════════════════════════════
  sectionHeader('Security Posture', C.purple,
    "Overview of the tenant's security configuration including Microsoft Secure Score, Conditional Access (CA) policy status, authentication methods, and identity protection alerts.");
  {
    const secKpis = [];
    if (d.score) secKpis.push(k(`${d.score.pct}%`, 'Secure Score', d.score.pct >= 70 ? 'good' : d.score.pct >= 50 ? 'warn' : 'bad'));
    if (d.conditionalAccess) {
      secKpis.push(k(d.conditionalAccess.enabled, 'CA Enforced', 'good'));
      secKpis.push(k(d.conditionalAccess.reportOnly, 'CA Report-only', d.conditionalAccess.reportOnly > 0 ? 'warn' : 'good'));
    }
    secKpis.push(k(d.risky ?? 0, 'Risky Users', (d.risky ?? 0) > 0 ? 'bad' : 'good'));
    secKpis.push(k(d.compliancePolicies?.total ?? 0, 'Compliance policies', 'neu'));
    secKpis.push(k(d.appProtection?.total ?? 0, 'App protection', (d.appProtection?.total ?? 0) === 0 ? 'warn' : 'good'));
    if (d.securityDefaults !== null && d.securityDefaults !== undefined) secKpis.push(k(d.securityDefaults ? 'On' : 'Off', 'Security Defaults', d.securityDefaults ? 'good' : 'neu'));
    kpiGrid(secKpis);
    if (d.score) {
      const msg = d.score.pct >= 70
        ? `Microsoft Secure Score is ${d.score.pct}% (${d.score.cur}/${d.score.max}) - a healthy baseline.`
        : d.score.pct >= 50
          ? `Microsoft Secure Score is ${d.score.pct}% (${d.score.cur}/${d.score.max}). There is room to improve.`
          : `Microsoft Secure Score is ${d.score.pct}% (${d.score.cur}/${d.score.max}). This requires attention.`;
      ensure(20); doc.setFont('helvetica', 'normal'); doc.setFontSize(10); setText(C.dark);
      doc.text(msg, M, y); y += 20;
    }
  }
  subLabel('Key policy status');
  {
    const kp = d.keyPolicies || {};
    policyTable([
      { label: 'Legacy Authentication Block', status: kp.legacyAuthBlock?.status || 'not found' },
      { label: 'MFA - All Users', status: kp.mfaAllUsers?.status || 'not found' },
      { label: 'Admin Phishing-Resistant MFA', status: kp.adminMfa?.status || 'not found' },
      { label: 'Geographic Restriction', status: kp.geoBlock?.status || 'not found' },
    ]);
  }
  subLabel('Authentication methods');
  {
    const am = d.authMethods || {};
    policyTable([
      { label: 'Microsoft Authenticator', status: am.authAppEnabled ? 'enabled' : 'disabled' },
      { label: 'FIDO2 / Passkeys', status: am.fido2Enabled ? 'enabled' : 'disabled' },
      { label: 'SMS sign-in', status: am.smsEnabled ? 'enabled-warn' : 'disabled' },
      { label: 'Temporary Access Pass', status: am.tapEnabled ? (am.tapReusable ? 'enabled-warn' : 'enabled') : 'disabled' },
    ]);
  }
  if ((d.risky ?? 0) > 0) callout(`${d.risky} account${d.risky > 1 ? 's are' : ' is'} flagged as at-risk by Entra ID Protection. Reset passwords immediately.`, 'bad');
  if ((d.appProtection?.total ?? 0) === 0) callout('No app protection (MAM) policies found. If BYOD access is permitted, this is a gap.', 'warn');

  // ══════════════════════════ Patch Status ═════════════════════════════════
  const ps = d.patchStatus || {};
  if (ps.current !== undefined) {
    sectionHeader('Patch Status', C.blue,
      'Patch currency is based on the last Intune check-in date. Devices that have not checked in within 30 days may be running unpatched software.');
    kpiGrid([
      k(ps.current ?? 0, 'Current (≤30 days)', 'good'),
      k(ps.over30 ?? 0, 'Over 30 days', (ps.over30 ?? 0) > 0 ? 'warn' : 'good'),
      k(ps.over90 ?? 0, 'Over 90 days', (ps.over90 ?? 0) > 0 ? 'bad' : 'good'),
    ]);
    if ((d.patchOver90 || []).length) {
      callout(`${d.patchOver90.length} device${d.patchOver90.length > 1 ? 's have' : ' has'} not checked in for 90+ days - 3 or more patch cycles behind.`, 'bad');
      subLabel('Devices with patch version older than 90 days');
      dataTable(['Device', 'Primary user', 'Latest patch version'],
        d.patchOver90.map(dev => [dev.name || '', dev.user || 'Unknown', dev.os || 'Unknown']));
    }
  }

  // ══════════════════════════ User Data ════════════════════════════════════
  const u = d.users || {};
  if ((u.total ?? 0) > 0) {
    sectionHeader('User Data', C.purple,
      'Summary of M365 licensed users, guest accounts, administrative role holders, and sign-in activity including accounts not used in the last 90 days and logins from outside Australia and New Zealand.');
    kpiGrid([
      k(u.total ?? 0, 'Total Licensed users', 'info'),
      k(u.sharedMailboxes ?? 0, 'Shared mailboxes', 'neu'),
      k(u.guests ?? 0, 'Guest users', (u.guests ?? 0) > 0 ? 'warn' : 'neu'),
      k(u.notSignedIn90Licensed ?? 0, 'M365 Not Signed in 90d+', (u.notSignedIn90Licensed ?? 0) > 0 ? 'warn' : 'good'),
      k(u.notSignedIn90Guest ?? 0, 'Guests not signed in 90d+', (u.notSignedIn90Guest ?? 0) > 0 ? 'warn' : 'good'),
      k((u.adminRoles || []).length, 'Admin role holders', (u.adminRoles || []).length > 0 ? 'warn' : 'neu'),
      k(u.externalSignIns?.total ?? 0, 'Overseas sign-ins (30d)', (u.externalSignIns?.uniqueUsers ?? 0) > 0 ? 'warn' : 'good'),
    ]);

    if ((u.notSignedIn90Licensed ?? 0) > 0) {
      callout(`${u.notSignedIn90Licensed} M365 licensed user${u.notSignedIn90Licensed > 1 ? 's have' : ' has'} not signed in for 90+ days. Service accounts and unlicensed accounts are excluded.`, 'warn');
      if ((u.notSignedIn90LicensedList || []).length) {
        subLabel('M365 licensed users - not signed in 90+ days');
        dataTable(['Name', 'Email', 'Last sign-in', 'Days'], u.notSignedIn90LicensedList.map(usr => [
          usr.name || '', usr.upn || '', usr.lastSignIn ? new Date(usr.lastSignIn).toLocaleDateString('en-NZ') : 'Never',
          usr.daysSince != null ? String(usr.daysSince) : 'Never',
        ]));
      }
    }
    if ((u.notSignedIn90Guest ?? 0) > 0) {
      callout(`${u.notSignedIn90Guest} guest account${u.notSignedIn90Guest > 1 ? 's have' : ' has'} not signed in for 90+ days. Review for stale guest access and remove if no longer needed.`, 'warn');
      if ((u.notSignedIn90GuestList || []).length) {
        subLabel('Guest accounts - not signed in 90+ days');
        dataTable(['Name', 'Email', 'Last sign-in', 'Days'], u.notSignedIn90GuestList.map(usr => [
          usr.name || '', usr.upn || '', usr.lastSignIn ? new Date(usr.lastSignIn).toLocaleDateString('en-NZ') : 'Never',
          usr.daysSince != null ? String(usr.daysSince) : 'Never',
        ]));
      }
    }
    if ((u.licenceSummary || []).length) { subLabel('Licence assignment'); licenceCards(u.licenceSummary); }
    subLabel('Upcoming licence renewals (next 90 days)');
    if (u.licenceRenewalsError) {
      callout('Licence renewal data unavailable - check that Organization.Read.All permission is granted.', 'warn');
    } else if ((u.licenceRenewals || []).length) {
      callout(`${u.licenceRenewals.length} licence subscription${u.licenceRenewals.length > 1 ? 's are' : ' is'} due for renewal within 90 days. Monthly-billed subscriptions renew automatically each cycle; annual/multi-year ones are worth a closer look.`, 'info');
      dataTable(['Licence', 'Seats', 'Billing', 'Renewal date', 'Days away'], u.licenceRenewals.map(r => [
        r.name || '', String(r.seats ?? ''), r.cycle || '', fmtDate(r.renewalDate), String(r.daysAway ?? ''),
      ]));
    } else {
      callout('No annual licence renewals due in the next 90 days.', 'good');
    }
    if ((u.adminRoles || []).length) {
      subLabel('Users with admin (privileged) roles');
      dataTable(['User', 'Roles'], u.adminRoles.map(a => [`${a.name || ''}\n${a.upn || ''}`, (a.roles || [a.role]).join('\n')]));
    }
    const extUsers = u.externalSignIns?.byUser || [];
    if (extUsers.length) {
      const extTotal = u.externalSignIns?.total ?? 0;
      const extWindow = u.externalSignIns?.windowDays ?? 30;
      subLabel(`Unexpected overseas sign-ins - last ${extWindow} days`);
      callout(`${extTotal} successful login${extTotal > 1 ? 's' : ''} from unexpected locations by ${u.externalSignIns.uniqueUsers} user${u.externalSignIns.uniqueUsers > 1 ? 's' : ''}. Australia, New Zealand and Malaysia are excluded as expected locations.`, 'warn');
      dataTable(['Name', 'Email', 'Country / Territory', 'Logins'], extUsers.map(e => [
        e.name || (e.upn || '').split('@')[0], e.upn || '', (e.countries || []).join(', '), String(e.eventCount ?? 0),
      ]));
    }
  }

  // ══════════════════════════ SharePoint / MS Teams ════════════════════════
  const sp = d.sharepoint || {};
  if ((sp.siteCount ?? 0) > 0 || sp.error) {
    sectionHeader('SharePoint / MS Teams', C.teal,
      'Overview of SharePoint Online sites and storage usage for this tenant. Site counts are sourced from the Microsoft 365 usage reports and exclude personal OneDrive sites.');
    if (sp.error) {
      callout('SharePoint data unavailable - check that Reports.Read.All permission is granted.', 'warn');
    } else {
      const spKpis = [
        k(sp.siteCount ?? 0, 'Total sites', 'info'),
        k(sp.groupCount ?? 0, 'M365 Group sites', 'neu'),
        k(sp.commCount ?? 0, 'Communication sites', 'neu'),
        k(sp.classicCount ?? 0, 'Classic / other', 'neu'),
        ...((sp.m365GroupCount ?? 0) > 0 ? [k(sp.m365GroupCount, 'M365 Groups', 'neu')] : []),
        ...((sp.securityGroupCount ?? 0) > 0 ? [k(sp.securityGroupCount, 'Security groups', 'neu')] : []),
        ...(sp.allocatedGB ? [k(`${sp.allocatedGB} GB`, 'Storage allocated', 'neu')] : []),
        k(`${sp.totalUsedGB ?? 0} GB`, 'Storage used', 'neu'),
        k(sp.inactiveSiteCount ?? 0, 'Inactive 180d+', (sp.inactiveSiteCount ?? 0) > 0 ? 'warn' : 'good'),
      ];
      kpiGrid(spKpis);
      if (sp.allocatedGB && sp.totalUsedGB) {
        const usedPct = Math.round((sp.totalUsedGB / sp.allocatedGB) * 100);
        callout(`Storage: ${sp.totalUsedGB} GB used of ${sp.allocatedGB} GB allocated (${usedPct}% used).`, usedPct > 80 ? 'bad' : usedPct > 60 ? 'warn' : 'good');
      }
    }
  }

  // ══════════════════════════ Ticketing & Support ══════════════════════════
  const hasTickets = manual.ticketsOpened || manual.ticketsClosed || manual.ticketsPending || manual.avgResponse || manual.avgResolution || manual.p1;
  if (hasTickets) {
    sectionHeader('Ticketing & Support', C.teal);
    const tKpis = [];
    if (manual.ticketsOpened) tKpis.push(k(manual.ticketsOpened, 'Opened', 'neu'));
    if (manual.ticketsClosed) tKpis.push(k(manual.ticketsClosed, 'Closed', 'good'));
    if (manual.ticketsPending) tKpis.push(k(manual.ticketsPending, 'Pending', parseInt(manual.ticketsPending, 10) > 5 ? 'warn' : 'neu'));
    if (tKpis.length) kpiGrid(tKpis);
    const slaKpis = [];
    if (manual.avgResponse) slaKpis.push(k(manual.avgResponse, 'Avg response time', 'neu'));
    if (manual.avgResolution) slaKpis.push(k(manual.avgResolution, 'Avg resolution time', 'neu'));
    if (slaKpis.length) kpiGrid(slaKpis);
    callout(manual.p1 ? 'P1 incident: ' + manual.p1 : 'No P1 incidents recorded this period.', manual.p1 ? 'bad' : 'good');
  }

  // ══════════════════════════ Recommendations & Opportunities ══════════════
  if (manual.personal || manual.improvements || manual.cost || manual.roadmap) {
    sectionHeader('Recommendations & Opportunities', C.orange);
    const card = (title, colour, text) => {
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5);
      const lines = doc.splitTextToSize(text, CW - 30);
      const h = lines.length * 12 + 30;
      ensure(h + 10);
      setFill(C.bgray); doc.rect(M, y, CW, h, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8); setText(colour);
      doc.text(title.toUpperCase(), M + 15, y + 16);
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); setText(C.dark);
      doc.text(lines, M + 15, y + 30);
      y += h + 12;
    };
    if (manual.personal) {
      doc.setFont('helvetica', 'italic'); doc.setFontSize(9.5); setText([55, 48, 163]);
      const lines = doc.splitTextToSize(manual.personal, CW - 30);
      const h = lines.length * 12 + 20; ensure(h + 10);
      setFill([239, 246, 255]); doc.rect(M, y, CW, h, 'F'); setFill(C.blue); doc.rect(M, y, 3, h, 'F');
      doc.text(lines, M + 15, y + 16); y += h + 12;
    }
    if (manual.improvements) card('Suggested improvements', C.blue, manual.improvements);
    if (manual.cost) card('Cost-saving opportunities', C.teal, manual.cost);
    if (manual.roadmap) card('Tech roadmap', C.purple, manual.roadmap);
  }

  // ══════════════════════════ Security Risk Register (bottom) ═════════════
  if ((d.risks || []).length) {
    sectionHeader('Security Risk Register', C.red,
      'The following risks have been identified based on the data collected from this tenant. Items are ranked by severity. High-severity items require prompt attention.');
    riskRegister(d.risks);
  }

  // ── Footers ────────────────────────────────────────────────────────────
  const pages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    setDraw(C.border); doc.setLineWidth(0.5); doc.line(M, PH - 30, PW - M, PH - 30);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8); setText(C.lgray);
    doc.text('Integricity Technology  ·  Confidential', PW / 2, PH - 18, { align: 'center' });
    doc.text(`Page ${p} of ${pages}`, PW - M, PH - 18, { align: 'right' });
  }

  return doc;
}

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: CORS, body: '' };
  if (event.httpMethod !== 'POST') return { statusCode: 405, headers: CORS, body: 'Method Not Allowed' };

  let payload;
  try { payload = JSON.parse(event.body); }
  catch { return { statusCode: 400, headers: CORS, body: 'Invalid JSON' }; }

  try {
    const doc = buildPdfDoc(payload);
    const buffer = Buffer.from(doc.output('arraybuffer'));
    const filename = `${payload.client} IT Health Report.pdf`;
    return {
      statusCode: 200,
      headers: { ...CORS, 'Content-Type': 'application/pdf', 'Content-Disposition': `attachment; filename="${filename.replace(/"/g, '')}"` },
      isBase64Encoded: true,
      body: buffer.toString('base64'),
    };
  } catch (err) {
    console.error('export-pdf error:', err);
    return { statusCode: 500, headers: CORS, body: JSON.stringify({ message: err.message }) };
  }
};
