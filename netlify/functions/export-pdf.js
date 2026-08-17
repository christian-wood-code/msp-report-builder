"use strict";

// PDF export (server-side, new -- this app previously had no PDF export at
// all, only "use your browser's print-to-PDF" per the README). Same
// {client, from, to, preparer, today, iData, manual} input shape as
// export-docx.js so report-hub can call both the same way.
//
// First-pass / condensed: covers the highest-value sections (header stats,
// Security Risk Register, Device/Compliance/Encryption/AV, Patch Status,
// Users, SharePoint) rather than an exhaustive mirror of every one of the 52
// metrics in the Word export -- there was no existing client-side jsPDF file
// to port from (unlike the other three report apps), so this was written
// fresh against intune.js's real response shape. Expand section-by-section
// later if the condensed version isn't enough.
//
// No Inter font file exists in this project (unlike the sibling apps) --
// Helvetica throughout, consistent with jsPDF's own "fails open to
// Helvetica" pattern used elsewhere when Inter isn't available.

const { jsPDF } = require('jspdf');
require('jspdf-autotable');

const CORS = { 'Access-Control-Allow-Origin': '*' };

const PDF = {
  navy: [15, 23, 42], navy700: [51, 65, 85],
  blue: [13, 124, 196], green: [15, 122, 60], amber: [217, 119, 6], red: [192, 21, 42],
  slate: [107, 114, 128], slate200: [229, 231, 235], slate50: [248, 249, 252], white: [255, 255, 255],
};
const SEVERITY_RGB = { high: PDF.red, medium: PDF.amber, low: PDF.slate };

function buildPdfDoc({ client, from, to, iData: d, manual = {} }) {
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

  const sectionTitle = (title, rightLabel = '') => {
    ensure(44);
    setFill([239, 246, 255]); doc.roundedRect(M, y - 3, 15, 15, 4, 4, 'F');
    setText(PDF.navy); doc.setFont('helvetica', 'bold'); doc.setFontSize(13);
    doc.text(title, M + 22, y + 8);
    if (rightLabel) {
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8); setText(PDF.slate);
      doc.text(rightLabel.toUpperCase(), PW - M, y + 6, { align: 'right' });
    }
    y += 20;
    setDraw(PDF.slate200); doc.setLineWidth(0.7); doc.line(M, y, PW - M, y);
    y += 16;
  };

  const tableStyles = {
    styles: { font: 'helvetica', fontSize: 9, cellPadding: 5, lineColor: PDF.slate200, lineWidth: 0.5, textColor: PDF.navy700 },
    headStyles: { font: 'helvetica', fillColor: PDF.navy, textColor: 255, fontStyle: 'bold', halign: 'left' },
    alternateRowStyles: { fillColor: PDF.slate50 },
  };
  const table = (head, rows) => {
    if (!rows.length) return;
    doc.autoTable({ startY: y, margin: { left: M, right: M }, head: [head], body: rows, theme: 'grid', ...tableStyles });
    y = doc.lastAutoTable.finalY + 18;
  };

  const statRow = items => {
    const gap = 10, n = items.length;
    const w = (CW - gap * (n - 1)) / n, h = 56;
    ensure(h + 18);
    items.forEach((it, i) => {
      const x = M + i * (w + gap);
      setFill(PDF.white); doc.roundedRect(x, y, w, h, 6, 6, 'F');
      setDraw(PDF.slate200); doc.setLineWidth(0.7); doc.roundedRect(x, y, w, h, 6, 6, 'S');
      setFill(it.colour || PDF.navy); doc.roundedRect(x, y, w, 3, 1.5, 1.5, 'F');
      setText(it.colour || PDF.navy); doc.setFont('helvetica', 'bold'); doc.setFontSize(18);
      doc.text(String(it.value), x + 12, y + 26);
      setText(PDF.slate); doc.setFont('helvetica', 'bold'); doc.setFontSize(7);
      doc.text(it.label.toUpperCase(), x + 12, y + 38, { maxWidth: w - 20 });
      if (it.sub) { doc.setFont('helvetica', 'normal'); doc.setFontSize(6.5); doc.text(it.sub, x + 12, y + 48, { maxWidth: w - 20 }); }
    });
    y += h + 18;
  };

  // ── Header band ────────────────────────────────────────────────────────
  const BANDH = 108;
  setFill(PDF.navy); doc.rect(0, 0, PW, BANDH, 'F');
  doc.setFont('helvetica', 'bold'); setText(PDF.white); doc.setFontSize(19);
  doc.text(`${client} IT Health Report`, M, 50);
  doc.setFont('helvetica', 'normal'); doc.setFontSize(10); setText([148, 163, 184]);
  const fmtDate = s => s ? new Date(s).toLocaleDateString('en-NZ', { day: 'numeric', month: 'long', year: 'numeric' }) : '';
  doc.text(`${fmtDate(from)} – ${fmtDate(to)}`, M, 70);
  doc.setFontSize(9);
  doc.text('Integricity Technology · Confidential', M, 86);

  y = BANDH + 24;

  // ── Top stat band ──────────────────────────────────────────────────────
  statRow([
    { value: d.total ?? 0, label: 'Managed Devices', colour: PDF.blue },
    { value: d.comp?.noncompliant ?? 0, label: 'Non-Compliant', colour: (d.comp?.noncompliant ?? 0) ? PDF.red : PDF.green },
    { value: d.score ? `${d.score.pct}%` : '—', label: 'Secure Score', colour: PDF.navy },
    { value: d.risks?.length ?? 0, label: 'Risk Findings', colour: (d.risks?.length ?? 0) ? PDF.amber : PDF.green },
  ]);

  if (manual.overview) {
    doc.setFont('helvetica', 'normal'); doc.setFontSize(9.5); setText(PDF.navy700);
    const lines = doc.splitTextToSize(manual.overview, CW);
    ensure(lines.length * 12 + 10);
    doc.text(lines, M, y + 9);
    y += lines.length * 12 + 16;
  }

  // ── Security Risk Register (leads -- "what to actually do") ────────────
  sectionTitle('Security Risk Register', d.risks?.length ? `${d.risks.length} findings` : 'None');
  if (!d.risks || !d.risks.length) {
    doc.setFont('helvetica', 'bold'); doc.setFontSize(10); setText(PDF.green);
    doc.text('No findings — nothing crossed a risk threshold this period.', M, y); y += 22;
  } else {
    d.risks.forEach(r => {
      const descLines = doc.splitTextToSize(r.finding, CW - 10);
      const actLines = doc.splitTextToSize(r.action, CW - 10);
      const h = 20 + descLines.length * 11 + actLines.length * 11 + 16;
      ensure(h);
      const accent = SEVERITY_RGB[r.severity] || PDF.navy;
      setFill(accent); doc.rect(M, y, 3, h - 8, 'F');
      doc.setFont('helvetica', 'bold'); doc.setFontSize(8); setText(accent);
      doc.text(r.severity.toUpperCase(), M + 12, y + 10);
      setText(PDF.navy); doc.setFontSize(10.5);
      doc.text(r.area, M + 12 + doc.getTextWidth(r.severity.toUpperCase()) + 10, y + 10);
      y += 20;
      doc.setFont('helvetica', 'normal'); doc.setFontSize(9); setText(PDF.navy700);
      doc.text(descLines, M + 12, y); y += descLines.length * 11 + 6;
      doc.setFont('helvetica', 'bold'); doc.setFontSize(9); setText(PDF.blue);
      doc.text('Action: ', M + 12, y);
      doc.setFont('helvetica', 'normal'); setText(PDF.navy700);
      doc.text(actLines, M + 12 + doc.getTextWidth('Action: '), y);
      y += Math.max(actLines.length, 1) * 11 + 16;
    });
  }

  // ── Device & Compliance ──────────────────────────────────────────────────
  sectionTitle('Device & Asset Management', `${d.total ?? 0} devices`);
  const compTotal = (d.comp?.compliant ?? 0) + (d.comp?.noncompliant ?? 0) + (d.comp?.unknown ?? 0);
  statRow([
    { value: d.comp?.compliant ?? 0, label: 'Compliant', colour: PDF.green },
    { value: d.comp?.noncompliant ?? 0, label: 'Non-Compliant', colour: PDF.red },
    { value: d.staleCount ?? 0, label: 'Stale (90+ days)', colour: PDF.amber },
    { value: `${compTotal ? Math.round(((d.comp?.compliant ?? 0) / compTotal) * 100) : 0}%`, label: 'Compliance Rate', colour: PDF.navy },
  ]);
  if (d.notCompliantList?.length) {
    table(['Device', 'User', 'OS', 'Last Sync'], d.notCompliantList.slice(0, 40).map(x => [x.name, x.user, x.os, x.lastSync ? new Date(x.lastSync).toLocaleDateString() : '—']));
  }

  sectionTitle('Encryption & Antivirus');
  statRow([
    { value: d.encryption?.encrypted ?? 0, label: 'Encrypted', colour: PDF.green },
    { value: d.encryption?.notEncrypted ?? 0, label: 'Not Encrypted', colour: PDF.red },
    { value: d.av?.active ?? 0, label: 'AV Active', colour: PDF.green },
    { value: (d.av?.notActive ?? 0) + (d.av?.outOfDate ?? 0), label: 'AV Issues', colour: (d.av?.notActive || d.av?.outOfDate) ? PDF.red : PDF.green },
  ]);
  if (d.notEncryptedList?.length) table(['Device', 'User', 'OS'], d.notEncryptedList.slice(0, 40).map(x => [x.name, x.user, x.os]));

  sectionTitle('Patch Status');
  statRow([
    { value: d.patchStatus?.current ?? 0, label: 'Current', colour: PDF.green },
    { value: d.patchStatus?.over30 ?? 0, label: '30–90 Days', colour: PDF.amber },
    { value: d.patchStatus?.over90 ?? 0, label: '90+ Days', colour: PDF.red },
    { value: d.win10 ?? 0, label: 'Windows 10 (EOL)', colour: (d.win10 ?? 0) ? PDF.red : PDF.green },
  ]);
  if (d.patchOver90?.length) table(['Device', 'User', 'Last Seen', 'OS'], d.patchOver90.slice(0, 40).map(x => [x.name, x.user, x.lastSeen ? new Date(x.lastSeen).toLocaleDateString() : '—', x.os]));

  // ── Security Posture ─────────────────────────────────────────────────────
  sectionTitle('Security Posture');
  statRow([
    { value: d.score ? `${d.score.pct}%` : '—', label: 'Secure Score', colour: PDF.navy },
    { value: d.conditionalAccess?.enabled ?? 0, label: 'CA Policies Enabled', colour: PDF.blue },
    { value: d.risky ?? 0, label: 'At-Risk Users', colour: (d.risky ?? 0) ? PDF.red : PDF.green },
    { value: d.securityDefaults === true ? 'On' : d.securityDefaults === false ? 'Off' : '—', label: 'Security Defaults', colour: PDF.navy },
  ]);
  const kp = d.keyPolicies || {};
  table(['Policy', 'Status'], [
    ['Legacy Auth Block', kp.legacyAuthBlock?.status || '—'],
    ['MFA — All Users', kp.mfaAllUsers?.status || '—'],
    ['MFA — Admins', kp.adminMfa?.status || '—'],
    ['Geo-Block', kp.geoBlock?.status || '—'],
  ]);

  // ── Users ────────────────────────────────────────────────────────────────
  sectionTitle('User Data', `${d.users?.total ?? 0} licensed`);
  statRow([
    { value: d.users?.total ?? 0, label: 'Licensed Users', colour: PDF.blue },
    { value: d.users?.guests ?? 0, label: 'Guests', colour: PDF.navy },
    { value: d.users?.notSignedIn90Licensed ?? 0, label: 'Inactive 90+ Days', colour: (d.users?.notSignedIn90Licensed ?? 0) ? PDF.amber : PDF.green },
    { value: d.users?.adminRoles?.length ?? 0, label: 'Admin Role Holders', colour: PDF.navy },
  ]);
  if (d.users?.adminRoles?.length) {
    table(['Name', 'Roles'], d.users.adminRoles.map(a => [a.name, a.roles.join(', ')]));
  }
  if (d.users?.licenceSummary?.length) {
    table(['Licence', 'Assigned'], d.users.licenceSummary.map(l => [l.name, l.count]));
  }

  // ── SharePoint ───────────────────────────────────────────────────────────
  const sp = d.sharepoint || {};
  sectionTitle('SharePoint / Teams', `${sp.siteCount ?? 0} sites`);
  statRow([
    { value: sp.siteCount ?? 0, label: 'Sites', colour: PDF.blue },
    { value: `${sp.totalUsedGB ?? 0} GB`, label: 'Storage Used', colour: PDF.navy },
    { value: sp.inactiveSiteCount ?? 0, label: 'Inactive 180+ Days', colour: (sp.inactiveSiteCount ?? 0) ? PDF.amber : PDF.green },
  ]);

  // ── Footers ────────────────────────────────────────────────────────────
  const pages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    setDraw(PDF.slate200); doc.setLineWidth(0.5); doc.line(M, PH - 30, PW - M, PH - 30);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(8); setText(PDF.slate);
    doc.text('Integricity Technology · Confidential', M, PH - 18);
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
