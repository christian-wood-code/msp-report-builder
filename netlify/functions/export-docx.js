"use strict";

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, TabStopType,
  LevelFormat, Footer, PageNumber, ImageRun, HeightRule, TableLayoutType,
  VerticalAlign, LineRuleType,
} = require("docx");

// ── Palette (light theme of the approved design mockup) ──────────────────────
const C = {
  BLUE:"0D7CC4", PURPLE:"8C0C6E", RED:"D6293E", ORANGE:"E06400", TEAL:"1A8070",
  DARK:"0F172A", GRAY:"64748B", LGRAY:"94A3B8",
  SURFACE:"F8FAFC", HEAD_BG:"F1F5F9", BORDER:"E5E9F2", WHITE:"FFFFFF",
  GOOD:"0F9D6B", GOOD_BG:"E7F6EF", WARN:"C77700", WARN_BG:"FFF4E0",
  BAD:"D6293E", BAD_BG:"FDEBEE", INFO:"0D7CC4", INFO_BG:"E6F2FB",
  NEU:"475569", NEU_BG:"EEF1F6",
};
// Semantic tones: fg = text/accent colour, bg = tint
const TONE = {
  good: { fg: C.GOOD, bg: C.GOOD_BG },
  warn: { fg: C.WARN, bg: C.WARN_BG },
  bad:  { fg: C.BAD,  bg: C.BAD_BG },
  info: { fg: C.INFO, bg: C.INFO_BG },
  neu:  { fg: C.NEU,  bg: C.NEU_BG },
};

// A4 (11906) minus 2 x 1134 margins = 9638 DXA of usable width
const PW = 9638;
const GAP = 120; // gutter between side-by-side cards

// ── Border helpers ────────────────────────────────────────────────────────────
const bdr  = (color = C.BORDER, size = 4) => ({ style: BorderStyle.SINGLE, size, color });
const none = () => ({ style: BorderStyle.NONE, size: 0, color: C.WHITE });
const allB = (color = C.BORDER, size = 4) => ({ top: bdr(color,size), bottom: bdr(color,size), left: bdr(color,size), right: bdr(color,size) });
const noB  = () => ({ top: none(), bottom: none(), left: none(), right: none() });
const rowB = (color = C.BORDER, size = 4) => ({ top: none(), bottom: bdr(color,size), left: none(), right: none() });
const NO_TABLE_BORDERS = { top: none(), bottom: none(), left: none(), right: none(), insideHorizontal: none(), insideVertical: none() };

// ── Text & paragraph helpers ─────────────────────────────────────────────────
const run = (text, { size=20, bold=false, color="333333", font="Arial", italics=false, shade } = {}) =>
  new TextRun({ text: String(text ?? ""), font, size, bold, color, italics,
    ...(shade ? { shading: { type: ShadingType.CLEAR, fill: shade, color: "auto" } } : {}) });

// Shaded "pill" text (Word cannot round the corners, but shading + padding reads as a tag)
const tag = (text, tone = "neu", size = 16) => {
  const t = TONE[tone] || TONE.neu;
  return run(` ${text} `, { size, bold: true, color: t.fg, shade: t.bg });
};

const para = (children, { before=0, after=0, align=AlignmentType.LEFT, border, keepNext=false, tabStops, shading, indent } = {}) => {
  const opts = { children: Array.isArray(children) ? children : [children], alignment: align, spacing: { before, after }, keepNext };
  if (border) opts.border = border;
  if (tabStops) opts.tabStops = tabStops;
  if (shading) opts.shading = shading;
  if (indent) opts.indent = indent;
  return new Paragraph(opts);
};

// Fixed-height gap paragraph (exact line height so it never grows with font size)
const gap = (h = 120) => new Paragraph({ children: [], spacing: { before: 0, after: 0, line: h, lineRule: LineRuleType.EXACT } });
// Tiny trailing paragraph for cells that end in a nested table
const tinyPara = () => new Paragraph({ children: [], spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT } });

const LEAD = { size: 19, color: "475569" };
const lead = text => para([run(text, LEAD)], { after: 140, keepNext: true, align: AlignmentType.RIGHT, indent: { left: 3400 } });

// Section heading: accent bar (left border) + title + right-aligned THIS MONTH tag.
// One top-level paragraph with keepNext so the heading always travels with the
// block that follows it (paragraphs nested in table cells cannot chain).
const sectionHeader = (title, color) => new Paragraph({
  tabStops: [{ type: TabStopType.RIGHT, position: PW }],
  border: { left: bdr(color, 72), bottom: bdr(C.BORDER, 4), top: none(), right: none() },
  indent: { left: 200 },
  spacing: { before: 360, after: 160 },
  keepNext: true,
  children: [
    run(title, { size: 32, bold: true, color: C.DARK }),
    run("\tTHIS MONTH", { size: 15, bold: true, color: C.GRAY }),
  ],
});
const sectionHeaderBlock = (title, color) => [sectionHeader(title, color)];

const subLabel = text => new Paragraph({
  children: [run(text.toUpperCase(), { size: 16, bold: true, color: C.GRAY })],
  spacing: { before: 200, after: 100 },
  keepNext: true,
});

// Small label inside a card
const cardLabel = text => para([run(text.toUpperCase(), { size: 16, bold: true, color: C.GRAY })], { after: 100, keepNext: true });

// ── Dates ────────────────────────────────────────────────────────────────────
const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const dShort = s => {
  if (!s) return null;
  const t = new Date(s);
  if (isNaN(t)) return null;
  return `${t.getUTCDate()} ${MON[t.getUTCMonth()]} ${t.getUTCFullYear()}`;
};
const dOrNever = s => dShort(s) || "Never";

// ── Cells / tables ───────────────────────────────────────────────────────────
const cell = (children, { width, bg = C.WHITE, borders = noB(), margins = { top:80, bottom:80, left:120, right:120 }, vAlign, span } = {}) =>
  new TableCell({
    borders, shading: { fill: bg, type: ShadingType.CLEAR, color: "auto" },
    width: { size: width, type: WidthType.DXA },
    margins,
    verticalAlign: vAlign,
    columnSpan: span,
    children: Array.isArray(children) ? children : [children],
  });

const mkTable = (cols, rows, { borders = NO_TABLE_BORDERS } = {}) => new Table({
  width: { size: cols.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  columnWidths: cols,
  layout: TableLayoutType.FIXED,
  borders,
  rows,
});

// Column widths from fractions; remainder goes to the last column so they sum exactly
const colsFrom = (fracs, total = PW) => {
  const w = fracs.map(f => Math.floor(total * f));
  w[w.length - 1] += total - w.reduce((a, b) => a + b, 0);
  return w;
};

// n equal columns separated by gutters. Returns { cols, cellW: [widths of content cols] }
const gridCols = (n, total = PW, gutter = GAP) => {
  const cw = Math.floor((total - gutter * (n - 1)) / n);
  const cellW = Array(n).fill(cw);
  cellW[n - 1] += total - (cw * n + gutter * (n - 1));
  const cols = [];
  cellW.forEach((w, i) => { cols.push(w); if (i < n - 1) cols.push(gutter); });
  return { cols, cellW };
};

// ── KPI card grid (coloured top border, big coloured number) ─────────────────
const KPI_ACCENT = { good: C.GOOD, warn: C.WARN, bad: C.BAD, info: C.INFO, neu: "CBD5E1" };
const KPI_VALUE  = { good: C.GOOD, warn: C.WARN, bad: C.BAD, info: C.INFO, neu: C.DARK };

const k = (value, label, cls) => ({ value, label, cls });

function kpiGrid(items, cols, total = PW) {
  cols = cols || Math.min(items.length, 4);
  const { cols: colW, cellW } = gridCols(cols, total);
  const rows = [];
  for (let i = 0; i < items.length; i += cols) {
    const chunk = items.slice(i, i + cols);
    if (rows.length) rows.push(new TableRow({
      height: { value: 100, rule: HeightRule.EXACT },
      children: colW.map(w => cell(gap(20), { width: w, margins: { top:0, bottom:0, left:0, right:0 } })),
    }));
    const cells = [];
    for (let j = 0; j < cols; j++) {
      const it = chunk[j];
      if (!it) {
        cells.push(cell(para([]), { width: cellW[j] }));
      } else {
        const cls = KPI_ACCENT[it.cls] ? it.cls : "neu";
        cells.push(cell([
          para([run(String(it.value ?? ""), { size: 44, bold: true, color: KPI_VALUE[cls] })], { after: 20, keepNext: true }),
          para([run(it.label || "", { size: 17, color: C.GRAY })]),
        ], { width: cellW[j], bg: C.SURFACE,
             borders: { top: bdr(KPI_ACCENT[cls], 18), bottom: bdr(C.BORDER), left: bdr(C.BORDER), right: bdr(C.BORDER) },
             margins: { top: 100, bottom: 100, left: 160, right: 120 } }));
      }
      if (j < cols - 1) cells.push(cell(para([]), { width: GAP, margins: { top:0, bottom:0, left:0, right:0 } }));
    }
    rows.push(new TableRow({ cantSplit: true, children: cells }));
  }
  return mkTable(colW, rows);
}

// ── Callout (accent bar + tinted box) ────────────────────────────────────────
const CALLOUT = {
  warn: { bg: C.WARN_BG, accent: C.WARN, text: "78350F" },
  bad:  { bg: C.BAD_BG,  accent: C.BAD,  text: "7F1D1D" },
  good: { bg: C.GOOD_BG, accent: C.GOOD, text: "14532D" },
  info: { bg: C.INFO_BG, accent: C.INFO, text: "1E3A8A" },
};

// `lead` = optional bold first sentence rendered before `text`
function callout(text, type, leadText) {
  const cfg = CALLOUT[type] || { bg: C.SURFACE, accent: C.BORDER, text: C.DARK };
  const runs = [];
  if (leadText) runs.push(run(leadText + " ", { size: 19, bold: true, color: cfg.text }));
  runs.push(run(text, { size: 19, color: cfg.text }));
  return mkTable([120, PW - 120], [new TableRow({ cantSplit: true, children: [
    cell(para([]), { width: 120, bg: cfg.accent, margins: { top:0, bottom:0, left:0, right:0 } }),
    cell(para(runs), { width: PW - 120, bg: cfg.bg,
      borders: { top: bdr(cfg.bg), bottom: bdr(cfg.bg), left: none(), right: bdr(cfg.bg) },
      margins: { top: 100, bottom: 100, left: 160, right: 140 } }),
  ]})]);
}

// Callout with a bold headline in the left cell and a note in the right cell
function splitCallout(headline, note, type) {
  const cfg = CALLOUT[type] || CALLOUT.info;
  const [a, b] = colsFrom([0.4, 0.6]);
  const bd = allB(cfg.accent, 4);
  return mkTable([a, b], [new TableRow({ cantSplit: true, children: [
    cell(para([run(headline, { size: 19, bold: true, color: cfg.text })]), { width: a, bg: cfg.bg, borders: { ...bd, right: none() }, margins: { top: 100, bottom: 100, left: 160, right: 100 }, vAlign: VerticalAlign.CENTER }),
    cell(para([run(note, { size: 18, color: cfg.text })]), { width: b, bg: cfg.bg, borders: { ...bd, left: none() }, margins: { top: 100, bottom: 100, left: 100, right: 160 }, vAlign: VerticalAlign.CENTER }),
  ]})]);
}

// ── Generic data table ───────────────────────────────────────────────────────
// headers: string[]; rows: array of arrays of cell values; widths: DXA[] summing to PW.
// A cell value is a string, or { t, b, c, i, size } for styled text, or { runs: TextRun[] },
// or { paras: Paragraph[] }, optionally with { bg }.
function dataTable(headers, rows, widths, { num = [], total = PW } = {}) {
  const alignOf = i => num.includes(i) ? AlignmentType.RIGHT : AlignmentType.LEFT;
  const hdr = new TableRow({ tableHeader: true, cantSplit: true, children: headers.map((h, i) =>
    cell(para([run(h.toUpperCase(), { size: 15, bold: true, color: C.GRAY })], { align: alignOf(i), keepNext: true }),
      { width: widths[i], bg: C.HEAD_BG, borders: { top: bdr(C.BORDER), bottom: bdr(C.BORDER, 8), left: none(), right: none() }, margins: { top: 70, bottom: 70, left: 110, right: 110 } })) });
  const body = rows.map((r, ri) => new TableRow({ cantSplit: true, children: r.map((v, i) => {
    const o = (v && typeof v === "object" && !Array.isArray(v)) ? v : { t: v };
    const keep = ri === 0; // keep header + first data row together
    let content;
    if (o.paras) content = o.paras;
    else if (o.runs) content = [para(o.runs, { align: alignOf(i), keepNext: keep })];
    else content = [para([run(o.t ?? "", { size: o.size || 18, bold: !!o.b, color: o.c || (o.b ? C.DARK : "334155"), italics: !!o.i })], { align: alignOf(i), keepNext: keep })];
    return cell(content, { width: widths[i], bg: o.bg || C.WHITE,
      borders: { top: none(), bottom: bdr(C.BORDER), left: none(), right: none() },
      margins: { top: 80, bottom: 80, left: 110, right: 110 }, vAlign: VerticalAlign.CENTER });
  })}));
  return mkTable(widths, [hdr, ...body]);
}

// ── Card helpers (bordered cells in a side-by-side layout) ───────────────────
const CARD_M = { top: 140, bottom: 140, left: 180, right: 180 };
const cardCell = (children, width, extra = {}) => cell(children, { width, bg: C.WHITE, borders: allB(C.BORDER), margins: CARD_M, ...extra });
const gapCell = () => cell(para([]), { width: GAP, margins: { top:0, bottom:0, left:0, right:0 } });

// Label ........ value rows (right tab), hairline between rows
function labelValueRows(rows, contentW) {
  return rows.map((r, i) => para(
    [run(r.label, { size: 19, color: C.DARK }), run("\t"), ...r.value],
    { tabStops: [{ type: TabStopType.RIGHT, position: contentW }], before: 70, after: 70,
      border: i < rows.length - 1 ? { bottom: bdr(C.BORDER, 4), top: none(), left: none(), right: none() } : undefined }));
}

// ── Stacked horizontal bar (single-row table, cell widths proportional) ──────
function stackedBar(parts, total = PW - 2 * 180, height = 220) {
  parts = parts.filter(p => p.n > 0);
  const sum = parts.reduce((a, p) => a + p.n, 0);
  if (!sum) return null;
  const MIN = 50;
  let w = parts.map(p => Math.max(MIN, Math.round(total * p.n / sum)));
  const diff = total - w.reduce((a, b) => a + b, 0);
  // absorb rounding / min-width drift in the widest segment
  const widest = w.indexOf(Math.max(...w));
  w[widest] += diff;
  return mkTable(w, [new TableRow({ height: { value: height, rule: HeightRule.EXACT }, cantSplit: true,
    children: parts.map((p, i) => cell(new Paragraph({ children: [run("", { size: 2 })], spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT } }),
      { width: w[i], bg: p.color, margins: { top:0, bottom:0, left:0, right:0 } })) })]);
}
function legendPara(items) {
  const runs = [];
  items.forEach((it, i) => {
    runs.push(run("■ ", { size: 18, color: it.color }));
    runs.push(run(it.label + (i < items.length - 1 ? "      " : ""), { size: 18, color: C.GRAY }));
  });
  return para(runs, { before: 100 });
}

// ── Rec card & personal message ───────────────────────────────────────────────
function recCard(title, titleColor, text) {
  return [
    mkTable([PW], [new TableRow({ children: [cell(
      [
        para([run(title.toUpperCase(), { size: 17, bold: true, color: titleColor })], { after: 60, keepNext: true }),
        para([run(text, { size: 19, color: C.DARK })]),
      ],
      { width: PW, bg: C.SURFACE, borders: { ...allB(C.BORDER), left: bdr(titleColor, 24) }, margins: { top:120, bottom:120, left:180, right:160 } }
    )]})]),
    gap(100),
  ];
}

function personalCard(text) {
  return mkTable([PW], [new TableRow({ children: [cell(
    para([run(text, { size: 20, italics: true, color: "3730A3" })]),
    { width: PW, bg: C.INFO_BG, borders: { top: bdr(C.BLUE, 6), bottom: bdr(C.PURPLE, 6), left: bdr(C.BLUE, 12), right: bdr(C.BORDER) }, margins: { top:140, bottom:140, left:200, right:160 } }
  )]})]);
}

// ── CORS / response ───────────────────────────────────────────────────────────
const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "Content-Type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Content-Type": "application/json",
  "X-Content-Type-Options": "nosniff",
};

// Rate limiter - 20 exports per IP per 5 minutes
const exportRateMap = new Map();
function checkExportRate(ip) {
  const now = Date.now(), WIN = 5*60*1000, LIMIT = 20;
  const e = exportRateMap.get(ip) || { count:0, windowStart:now };
  if (now - e.windowStart > WIN) { e.count = 0; e.windowStart = now; }
  e.count++; exportRateMap.set(ip, e);
  if (exportRateMap.size > 500) for (const [k,v] of exportRateMap) { if (now - v.windowStart > WIN) exportRateMap.delete(k); }
  return e.count <= LIMIT;
}

exports.handler = async (event) => {
  if (event.httpMethod === "OPTIONS") return { statusCode: 204, headers: CORS, body: "" };
  if (event.httpMethod !== "POST") return { statusCode: 405, headers: CORS, body: JSON.stringify({ error: "Method not allowed" }) };
  const clientIp = event.headers?.["x-forwarded-for"]?.split(",")[0]?.trim() || event.headers?.["x-nf-client-connection-ip"] || "unknown";
  if (!checkExportRate(clientIp)) return { statusCode: 429, headers: CORS, body: JSON.stringify({ error: "Too many requests - please wait a few minutes." }) };

  // Guard against oversized payloads (logos + data can be ~200KB, 2MB is generous)
  if (event.body && event.body.length > 2 * 1024 * 1024) {
    return { statusCode: 413, headers: CORS, body: JSON.stringify({ error: "Payload too large" }) };
  }
  let payload;
  try { payload = JSON.parse(event.body); }
  catch { return { statusCode: 400, headers: CORS, body: JSON.stringify({ error: "Invalid JSON" }) }; }

  try {
    const { client, from, to, preparer, today, iData: d } = payload;
    const manual = payload.manual || {};
    if (!client || !from || !to || !d) throw new Error("Missing required report data");
    const fmt = dt => new Date(dt).toLocaleDateString("en-NZ", { day: "numeric", month: "long", year: "numeric" });

    const coverLogo  = Buffer.from(payload.fullLogoTransparent || "", "base64");
    const headerLogo = Buffer.from(payload.fullLogoHeader || "", "base64");
    const hasLogo    = coverLogo.length > 100;
    const hasHeader  = headerLogo.length > 100;

    const children = [];

    // Defensive aliases so a sparse payload never throws mid-build
    const comp  = d.comp || { compliant: 0, noncompliant: 0, unknown: 0 };
    const total = d.total || 0;
    const u     = d.users || {};
    const renew = Array.isArray(u.licenceRenewals) ? u.licenceRenewals : [];
    const renewErr = !!u.licenceRenewalsError;
    const risks = d.risks || [];
    const pctComp = total > 0 ? Math.round((comp.compliant || 0) / total * 100) : 0;
    const scoreTone = p => p >= 70 ? "good" : p >= 50 ? "warn" : "bad";

    // ── Cover (navy hero) ────────────────────────────────────────────────────
    const monthLabel = new Date(from).toLocaleDateString("en-NZ", { month: "long", year: "numeric" });
    const NAVY = "0F1A3C", CHIP = "27345E", MUTED = "B9C6E6";
    const highRisks = risks.filter(r => r.severity === "high").length;
    const chip = (label, val) => [
      run(` ${label} `, { size: 17, color: MUTED, shade: CHIP }),
      ...(val ? [run(`${val} `, { size: 17, bold: true, color: C.WHITE, shade: CHIP })] : [run(" ", { size: 17, shade: CHIP })]),
      run("   ", { size: 17 }),
    ];
    const heroTop = [];
    if (hasLogo) heroTop.push(new ImageRun({ data: coverLogo, transformation: { width: 186, height: 48 }, type: "png" }));
    else heroTop.push(run("Integricity Technology", { size: 28, bold: true, color: C.WHITE }));
    heroTop.push(run("\t"), run(` ● Monthly IT report · ${monthLabel} `, { size: 17, bold: true, color: "E8F0FF", shade: CHIP }));
    children.push(mkTable([PW], [new TableRow({ cantSplit: true, children: [cell([
      para(heroTop, { after: 360, tabStops: [{ type: TabStopType.RIGHT, position: PW - 480 }] }),
      para([run(client, { size: 72, bold: true, color: C.WHITE })], { after: 80 }),
      para([run(`Report period: ${fmt(from)} – ${fmt(to)}`, { size: 22, color: MUTED })], { after: 240 }),
      para([
        ...chip("Prepared", today),
        ...chip("Source", "Microsoft Intune & Graph API"),
        ...(preparer ? chip("Account manager", preparer) : []),
        ...(highRisks > 0 ? chip(`${highRisks} high-priority risk${highRisks > 1 ? "s" : ""}`) : []),
      ], { after: 0 }),
    ], { width: PW, bg: NAVY, margins: { top: 300, bottom: 300, left: 240, right: 240 } })] })]));
    children.push(gap(200));

    // ── Headline strip (4 tiles on white, below the cover) ───────────────────
    {
      const nxt = renew.length ? [...renew].filter(r => r.renewalDate).sort((a, b) => String(a.renewalDate).localeCompare(String(b.renewalDate)))[0] : null;
      const tiles = [
        { label: "Devices compliant", val: total > 0 ? `${pctComp}%` : "N/A", tone: total > 0 ? (pctComp >= 95 ? "good" : pctComp >= 80 ? "warn" : "bad") : "neu", sub: total > 0 ? `${comp.compliant || 0} of ${total} devices` : "No devices" },
        { label: "Secure Score", val: d.score ? `${d.score.pct}%` : "N/A", tone: d.score ? scoreTone(d.score.pct) : "neu", sub: d.score ? `${d.score.cur} of ${d.score.max} points` : "Unavailable" },
        { label: "Risky users", val: String(d.risky || 0), tone: (d.risky || 0) > 0 ? "bad" : "good", sub: (d.risky || 0) > 0 ? "Accounts flagged at risk" : "No accounts at risk" },
        { label: "Renewals due (90 days)", val: renewErr ? "N/A" : String(renew.length), tone: renewErr ? "neu" : "info", sub: renewErr ? "Data unavailable" : `Next: ${nxt ? dShort(nxt.renewalDate) : "none"}` },
      ];
      const { cols, cellW } = gridCols(4);
      const cells = [];
      tiles.forEach((t, i) => {
        cells.push(cell([
          para([run(t.label.toUpperCase(), { size: 15, bold: true, color: C.GRAY })], { after: 40, keepNext: true }),
          para([run(t.val, { size: 64, bold: true, color: KPI_VALUE[t.tone] })], { after: 20, keepNext: true }),
          para([run(t.sub, { size: 17, color: C.GRAY })]),
        ], { width: cellW[i], bg: C.WHITE,
             borders: { top: bdr(KPI_ACCENT[t.tone], 24), bottom: bdr(C.BORDER), left: bdr(C.BORDER), right: bdr(C.BORDER) },
             margins: { top: 120, bottom: 120, left: 160, right: 120 } }));
        if (i < 3) cells.push(gapCell());
      });
      children.push(mkTable(cols, [new TableRow({ cantSplit: true, children: cells })]));
    }

    // ── Executive Summary ────────────────────────────────────────────────────
    if (manual.overview || manual.highlights || manual.concerns || manual.projects) {
      children.push(...sectionHeaderBlock("Executive Summary", C.BLUE));
      if (manual.overview) children.push(para([run(manual.overview, { size: 20 })], { after: 120 }));
      if (manual.highlights) { children.push(callout(manual.highlights, "good", "Highlights:")); children.push(gap(120)); }
      if (manual.concerns)   { children.push(callout(manual.concerns,   "warn", "Concerns:"));   children.push(gap(120)); }
      if (manual.projects)   { children.push(callout(manual.projects,   "info", "Projects:"));   children.push(gap(120)); }
    }

    // ── Month-on-Month Changes ────────────────────────────────────────────────
    // Deterministic template commentary (lib/metrics-commentary.js); omitted when null.
    if (d.commentary && d.commentary.length) {
      children.push(...sectionHeaderBlock("Month-on-Month Changes", C.ORANGE));
      children.push(lead("Notable changes since last month's report, based on the summary metrics from each run."));
      d.commentary.forEach(line => { children.push(callout(line, "info")); children.push(gap(100)); });
    }

    // ── Devices ──────────────────────────────────────────────────────────────
    const lowDisk = d.lowDisk || [];
    const enc = d.encryption || { encrypted: 0, notEncrypted: 0 };
    children.push(...sectionHeaderBlock("Device & Asset Management", C.BLUE));
    children.push(lead("Summary of managed devices enrolled in Microsoft Intune, including compliance status, operating system breakdown, and device health indicators."));
    children.push(kpiGrid([
      k(total, "Total devices", "info"),
      k(comp.compliant || 0, "Compliant", "good"),
      k(comp.noncompliant || 0, "Non-compliant", (comp.noncompliant || 0) > 0 ? "bad" : "good"),
      k(lowDisk.length, "Low disk (<15%)", lowDisk.length > 0 ? "warn" : "good"),
    ], 4));
    children.push(gap(140));

    // Compliance card | Encryption & antivirus card
    {
      const LW = Math.floor(PW * 0.40), RW = PW - LW - GAP;
      const legendLine = (color, text) => para([run("■ ", { size: 18, color }), run(text, { size: 18, color: C.GRAY })], { after: 30 });
      const left = [
        cardLabel("Compliance"),
        para([run(total > 0 ? `${pctComp}%` : "N/A", { size: 80, bold: true, color: total > 0 ? KPI_VALUE[pctComp >= 95 ? "good" : pctComp >= 80 ? "warn" : "bad"] : C.DARK })], { keepNext: true }),
        para([run("compliant", { size: 17, bold: true, color: C.GRAY })], { after: 120, keepNext: true }),
        legendLine(C.GOOD, `${comp.compliant || 0} compliant`),
        legendLine(C.BAD, `${comp.noncompliant || 0} non-compliant`),
        legendLine(C.NEU, `${comp.unknown || 0} unknown`),
        legendLine(C.WARN, `${d.staleCount || 0} not checked in for 90+ days`),
      ];
      const av = d.av;
      const avBad = av ? (av.outOfDate || 0) + (av.notActive || 0) : 0;
      const encItems = [
        k(enc.encrypted, "Encrypted", "good"),
        k(enc.notEncrypted, "Not encrypted", enc.notEncrypted > 0 ? "bad" : "good"),
        ...(av ? [k(av.active || 0, "AV active", "good"), k(avBad, "AV inactive / out of date", avBad > 0 ? "bad" : "good")] : []),
      ];
      const right = [cardLabel("Encryption & antivirus"), kpiGrid(encItems, 2, RW - 2 * 180 - 20), tinyPara()];
      children.push(mkTable([LW, GAP, RW], [new TableRow({ cantSplit: true, children: [
        cardCell(left, LW), gapCell(), cardCell(right, RW),
      ]})]));
      children.push(gap(140));
    }

    // Operating systems stacked bar
    {
      const OSP = [
        ["Windows 11 25H2", d.win25h2, "0D7CC4"], ["Windows 11 24H2", d.win24h2, "4AA3DF"], ["Windows 11", d.win11, "9CC9EA"],
        ["Windows 10", d.win10, "64748B"], ["macOS", d.macOS, "8C0C6E"], ["Linux", d.linux, "334155"],
        ["iOS / iPad", d.iosCount, "F5A01A"], ["Android", d.androidCount, "0F9D6B"],
      ].map(([label, n, color]) => ({ label, n: n || 0, color })).filter(p => p.n > 0);
      if (OSP.length) {
        const sumN = OSP.reduce((a, p) => a + p.n, 0);
        const barParts = OSP.map(p => ({ ...p }));
        if (total > sumN) barParts.push({ n: total - sumN, color: "E2E8F0" });
        children.push(mkTable([PW], [new TableRow({ cantSplit: true, children: [cardCell([
          cardLabel("Operating systems"),
          stackedBar(barParts),
          legendPara(OSP.map(p => ({ color: p.color, label: `${p.label} · ${p.n}` }))),
        ], PW)] })]));
        children.push(gap(100));
      }
    }

    if ((d.comp?.noncompliant || 0) > 0) {
      const n = d.comp.noncompliant;
      children.push(callout(`${n} device${n > 1 ? "s are" : " is"} non-compliant with your organisation's policies.`, "bad")); children.push(gap(100));
    }
    if ((d.notCompliantList || []).length > 0) {
      children.push(subLabel("Non-compliant devices"));
      children.push(dataTable(["Device", "Primary user", "OS", "Last seen"],
        d.notCompliantList.map(x => [{ t: x.name, b: true }, x.user || "Unknown", { t: x.os || "", size: 16 }, dOrNever(x.lastSync)]),
        colsFrom([0.25, 0.32, 0.25, 0.18])));
      children.push(gap(100));
    }
    if ((d.encryption?.notEncrypted || 0) > 0) {
      const n = d.encryption.notEncrypted;
      children.push(callout("This is a significant data protection risk.", "bad", `${n} device${n > 1 ? "s are" : " is"} not encrypted.`)); children.push(gap(100));
    }
    if ((d.notEncryptedList || []).length > 0) {
      children.push(subLabel("Devices not encrypted"));
      children.push(dataTable(["Device", "Primary user", "OS"],
        d.notEncryptedList.map(x => [{ t: x.name, b: true }, x.user || "Unknown", { t: x.os || "", size: 16 }]),
        colsFrom([0.28, 0.37, 0.35])));
      children.push(gap(100));
    }
    if (lowDisk.length > 0) {
      children.push(callout(`${lowDisk.length} device${lowDisk.length > 1 ? "s have" : " has"} low disk space (less than 15% free).`, "warn")); children.push(gap(100));
      children.push(subLabel("Low disk space"));
      children.push(dataTable(["Device", "Primary user", "Free", "Free %"],
        lowDisk.map(x => [{ t: x.name, b: true }, x.user || "Unknown", `${x.gb} GB`, { runs: [tag(`${x.pct}%`, x.pct < 5 ? "bad" : "warn")] }]),
        colsFrom([0.30, 0.42, 0.14, 0.14]), { num: [2, 3] }));
      children.push(gap(100));
    }

    // ── Security ──────────────────────────────────────────────────────────────
    children.push(...sectionHeaderBlock("Security Posture", C.PURPLE));
    children.push(lead("Overview of the tenant's security configuration including Microsoft Secure Score, Conditional Access (CA) policy status, authentication methods, and identity protection alerts."));
    {
      // Secure Score card (narrower) | Identity & access card (wider, 5:7)
      const LW = Math.floor((PW - GAP) * 5 / 12), RW = PW - GAP - LW;
      const sc = d.score;
      const left = [
        cardLabel("Microsoft Secure Score"),
        para([run(sc ? `${sc.pct}%` : "N/A", { size: 80, bold: true, color: sc ? KPI_VALUE[scoreTone(sc.pct)] : C.DARK })], { keepNext: true }),
        para([run(sc ? `${sc.cur} / ${sc.max} points` : "Score unavailable", { size: 17, bold: true, color: C.GRAY })], { after: 100, keepNext: true }),
        ...(sc ? [
          para([run(sc.pct >= 70 ? "Strong position." : "Room to improve.", { size: 19, bold: true, color: C.DARK })], { after: 20 }),
          para([run("Above 70% is a good target.", { size: 17, color: C.GRAY })]),
        ] : []),
      ];
      const ca = d.conditionalAccess;
      const idRows = [];
      if (d.securityDefaults !== null && d.securityDefaults !== undefined) idRows.push({ label: "Security defaults", value: [tag(d.securityDefaults ? "On" : "Off", "neu")] });
      if (ca) idRows.push({ label: "Conditional Access policies", value: [tag(`${ca.enabled} enforced`, "good"), run(" "), tag(`${ca.reportOnly} report-only`, ca.reportOnly > 0 ? "warn" : "neu")] });
      idRows.push({ label: "Risky users", value: [tag(String(d.risky || 0), (d.risky || 0) > 0 ? "bad" : "good")] });
      idRows.push({ label: "Compliance policies", value: [tag(String((d.compliancePolicies || {}).total || 0), "neu")] });
      idRows.push({ label: "App protection policies", value: [tag(String((d.appProtection || {}).total || 0), ((d.appProtection || {}).total || 0) === 0 ? "warn" : "neu")] });
      const right = [cardLabel("Identity & access"), ...labelValueRows(idRows, RW - 2 * 180)];
      children.push(mkTable([LW, GAP, RW], [new TableRow({ cantSplit: true, children: [
        cardCell(left, LW), gapCell(), cardCell(right, RW),
      ]})]));
      children.push(gap(140));
    }
    {
      const HW = Math.floor((PW - GAP) / 2), HW2 = PW - GAP - HW;
      const kp = d.keyPolicies || {};
      const stTag = s => ({ "enforced": ["Enforced", "good"], "report-only": ["Report-only", "warn"], "not found": ["Not found", "warn"], "disabled": ["Disabled", "bad"] }[s] || [String(s), "neu"]);
      const kpRows = [
        ["Block legacy authentication", kp.legacyAuthBlock], ["MFA for all users", kp.mfaAllUsers],
        ["MFA for admin portals", kp.adminMfa], ["Geographic restriction", kp.geoBlock],
      ].map(([label, v]) => { const [t, tone] = stTag((v || {}).status || "not found"); return { label, value: [tag(t, tone)] }; });
      const am = d.authMethods || {};
      const amRows = [
        { label: "Authenticator app", value: [am.authAppEnabled ? tag("Enabled", "good") : tag("Off", "neu")] },
        { label: "FIDO2 security keys", value: [am.fido2Enabled ? tag("Enabled", "good") : tag("Off", "neu")] },
        { label: "Temporary Access Pass", value: [am.tapEnabled ? tag(am.tapReusable ? "Enabled - review" : "Enabled", am.tapReusable ? "warn" : "good") : tag("Off", "neu")] },
        { label: "SMS", value: [am.smsEnabled ? tag("Enabled - review", "warn") : tag("Off", "neu")] },
      ];
      children.push(mkTable([HW, GAP, HW2], [new TableRow({ cantSplit: true, children: [
        cardCell([cardLabel("Key policies"), ...labelValueRows(kpRows, HW - 2 * 180)], HW),
        gapCell(),
        cardCell([cardLabel("Authentication methods"), ...labelValueRows(amRows, HW2 - 2 * 180)], HW2),
      ]})]));
      children.push(gap(100));
      if ((d.risky || 0) > 0) { children.push(callout("Reset passwords immediately.", "bad", `${d.risky} account${d.risky > 1 ? "s are" : " is"} flagged as at-risk by Entra ID Protection.`)); children.push(gap(100)); }
      if ((d.appProtection?.total || 0) === 0) { children.push(callout("If BYOD access is permitted, this is a gap.", "warn", "No app protection (MAM) policies found.")); children.push(gap(100)); }
    }

    // ── Patch Status ─────────────────────────────────────────────────────────
    const ps = d.patchStatus || {};
    if (ps.current !== undefined) {
      children.push(...sectionHeaderBlock("Patch Status", C.BLUE));
      children.push(lead("Patch currency is based on the last Intune check-in date. Devices that have not checked in within 30 days may be running unpatched software."));
      const parts = [
        { n: ps.current || 0, color: C.GOOD, label: `Current · ${ps.current || 0}` },
        { n: ps.over30 || 0,  color: C.WARN, label: `30+ days behind · ${ps.over30 || 0}` },
        { n: ps.over90 || 0,  color: C.BAD,  label: `90+ days behind · ${ps.over90 || 0}` },
      ];
      children.push(mkTable([PW], [new TableRow({ cantSplit: true, children: [cardCell([
        ...(stackedBar(parts) ? [stackedBar(parts)] : []),
        legendPara(parts.map(p => ({ color: p.color, label: p.label }))),
      ], PW)] })]));
      children.push(gap(100));
      if ((d.patchOver90 || []).length > 0) {
        children.push(callout(`${d.patchOver90.length} device${d.patchOver90.length > 1 ? "s have" : " has"} not checked in for 90+ days - 3 or more patch cycles behind.`, "bad")); children.push(gap(100));
        children.push(subLabel("Devices over 90 days without a check-in"));
        children.push(dataTable(["Device", "Primary user", "Last seen", "OS"],
          d.patchOver90.map(x => [{ t: x.name, b: true }, x.user || "Unknown", dOrNever(x.lastSeen), { t: x.os || "Unknown", size: 16 }]),
          colsFrom([0.22, 0.28, 0.14, 0.36])));
        children.push(gap(100));
      }
    }

    // ── User Data ─────────────────────────────────────────────────────────────
    if (u.total > 0) {
      children.push(...sectionHeaderBlock("User Data", C.PURPLE));
      children.push(lead("Summary of M365 licensed users, guest accounts, administrative role holders, and sign-in activity including accounts not used in the last 90 days and logins from outside Australia and New Zealand."));
      children.push(kpiGrid([
        k(u.total || 0, "Licensed users", "info"),
        k(u.sharedMailboxes || 0, "Shared mailboxes", "neu"),
        k(u.guests || 0, "Guest accounts", "neu"),
        k((u.notSignedIn90Licensed || 0) + (u.notSignedIn90Guest || 0), "Inactive 90+ days", ((u.notSignedIn90Licensed || 0) + (u.notSignedIn90Guest || 0)) > 0 ? "warn" : "good"),
      ], 4));
      children.push(gap(120));
      if ((u.externalSignIns || {}).timedOut) {
        children.push(callout("The sign-in log query timed out, so unexpected overseas sign-ins are not reported. Re-pulling usually resolves it.", "info", "Overseas sign-in data unavailable this run."));
        children.push(gap(100));
      }

      // Licence assignment: name | usage bar | "n / available" (fixed-width right column)
      if ((u.licenceSummary || []).length > 0) {
        const LN = Math.floor(PW * 0.36), LC = 2700, LB = PW - LN - LC;
        const BARW = LB - 2 * 100;
        const licRow = (l, last) => {
          const n = l.count || 0, a = l.available || 0;
          let bar, count;
          if (a === 0) {
            bar = stackedBar([{ n: 1, color: "CBD5E1" }], BARW, 140);
            count = [run(String(n), { size: 19, bold: true, color: C.DARK }), run(" assigned ", { size: 17, color: C.GRAY }), tag("none purchased", "neu", 14)];
          } else {
            const p = Math.min(100, n / a * 100);
            const col = p >= 100 ? C.BAD : p >= 90 ? C.WARN : l.free ? C.INFO : C.GOOD;
            const usedW = Math.min(BARW - 40, Math.max(50, Math.round(BARW * p / 100)));
            bar = p >= 100 ? stackedBar([{ n: 1, color: col }], BARW, 140)
              : mkTable([usedW, BARW - usedW], [new TableRow({ height: { value: 140, rule: HeightRule.EXACT }, cantSplit: true, children: [
                  cell(new Paragraph({ children: [run("", { size: 2 })], spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT } }), { width: usedW, bg: col, margins: { top:0, bottom:0, left:0, right:0 } }),
                  cell(new Paragraph({ children: [run("", { size: 2 })], spacing: { before: 0, after: 0, line: 20, lineRule: LineRuleType.EXACT } }), { width: BARW - usedW, bg: "E2E8F0", margins: { top:0, bottom:0, left:0, right:0 } }),
                ] })]);
            count = [run(String(n), { size: 19, bold: true, color: n > a ? C.BAD : C.DARK }), run(` / ${a.toLocaleString("en-NZ")}`, { size: 19, color: C.GRAY })];
          }
          const bd = { top: none(), bottom: bdr(C.BORDER), left: none(), right: none() };
          return new TableRow({ cantSplit: true, children: [
            cell(para([run(l.name || "", { size: 19, color: C.DARK })]), { width: LN, borders: bd, vAlign: VerticalAlign.CENTER, margins: { top: 70, bottom: 70, left: 100, right: 100 } }),
            cell([bar, tinyPara()], { width: LB, borders: bd, vAlign: VerticalAlign.CENTER, margins: { top: 70, bottom: 70, left: 100, right: 100 } }),
            cell(para(count, { align: AlignmentType.RIGHT }), { width: LC, borders: bd, vAlign: VerticalAlign.CENTER, margins: { top: 70, bottom: 70, left: 100, right: 100 } }),
          ]});
        };
        // Server already orders paid before free; keep order, only split for the group labels
        const paid = u.licenceSummary.filter(l => !l.free), free = u.licenceSummary.filter(l => l.free);
        children.push(subLabel("Licence assignment"));
        if (paid.length) {
          children.push(para([run("PAID LICENCES", { size: 15, bold: true, color: C.LGRAY })], { after: 40, keepNext: true }));
          children.push(mkTable([LN, LB, LC], paid.map(l => licRow(l))));
          children.push(gap(100));
        }
        if (free.length) {
          children.push(para([run("FREE LICENCES", { size: 15, bold: true, color: C.LGRAY })], { before: 60, after: 40, keepNext: true }));
          children.push(mkTable([LN, LB, LC], free.map(l => licRow(l))));
          children.push(gap(100));
        }
      }

      // Upcoming licence renewals -- always shown so "nothing due" reads as a completed check
      {
        children.push(subLabel("Upcoming licence renewals · next 90 days"));
        if (renewErr) {
          children.push(callout("Licence renewal data unavailable - check that Organization.Read.All permission is granted.", "warn"));
          children.push(gap(100));
        } else if (renew.length > 0) {
          children.push(splitCallout(
            `${renew.length} licence subscription${renew.length > 1 ? "s are" : " is"} due for renewal within 90 days.`,
            "Monthly-billed subscriptions renew automatically each cycle; annual/multi-year ones are worth a closer look.", "info"));
          children.push(gap(100));
          const sorted = [...renew].sort((a, b) => String(a.renewalDate || "").localeCompare(String(b.renewalDate || "")));
          children.push(dataTable(["Licence", "Seats", "Billing", "Renewal date", "Days away"],
            sorted.map(r => [
              { t: r.name || "", b: true },
              r.seats !== undefined && r.seats !== null ? Number(r.seats).toLocaleString("en-NZ") : "",
              { runs: [tag(String(r.cycle || "").replace("Annual/Multi-year", "Annual / multi-year"), r.cycle === "Monthly" ? "neu" : "info")] },
              dShort(r.renewalDate) || "",
              { runs: [tag(`${r.daysAway ?? ""} days`, (r.daysAway ?? 99) <= 14 ? "warn" : "neu")] },
            ]),
            colsFrom([0.31, 0.10, 0.22, 0.20, 0.17]), { num: [1, 4] }));
          children.push(gap(100));
        } else {
          children.push(callout("No licence renewals due in the next 90 days.", "good"));
          children.push(gap(100));
        }
      }

      if ((u.notSignedIn90Licensed || 0) > 0) {
        const n = u.notSignedIn90Licensed;
        children.push(callout("Service accounts and unlicensed accounts are excluded.", "warn", `${n} M365 licensed user${n > 1 ? "s have" : " has"} not signed in for 90+ days.`)); children.push(gap(100));
      }
      if ((u.notSignedIn90LicensedList || []).length > 0) {
        children.push(subLabel("Licensed users inactive for 90+ days"));
        children.push(dataTable(["User", "Account", "Last sign-in", "Inactive"],
          u.notSignedIn90LicensedList.map(x => [{ t: x.name || "", b: true }, { t: x.upn || "", size: 16 }, dOrNever(x.lastSignIn), x.daysSince ? `${x.daysSince} days` : "Never signed in"]),
          colsFrom([0.25, 0.33, 0.20, 0.22])));
        children.push(gap(100));
      }
      if ((u.notSignedIn90Guest || 0) > 0) {
        const n = u.notSignedIn90Guest;
        children.push(callout("Review for stale guest access and remove if no longer needed.", "warn", `${n} guest account${n > 1 ? "s have" : " has"} not signed in for 90+ days.`)); children.push(gap(100));
      }
      if ((u.notSignedIn90GuestList || []).length > 0) {
        children.push(subLabel("Guest accounts inactive for 90+ days"));
        children.push(dataTable(["Guest", "Account", "Last sign-in", "Inactive"],
          u.notSignedIn90GuestList.map(x => [{ t: x.name || "", b: true }, { t: String(x.upn || "").split("#EXT#")[0], size: 16 }, dOrNever(x.lastSignIn), x.daysSince ? `${x.daysSince} days` : "Never signed in"]),
          colsFrom([0.25, 0.33, 0.20, 0.22])));
        children.push(gap(100));
      }

      if ((u.adminRoles || []).length > 0) {
        children.push(subLabel("Admin role holders"));
        children.push(dataTable(["Name", "Account", "Roles"],
          u.adminRoles.map(a => {
            const roles = a.roles || (a.role ? [a.role] : []);
            const runs = [];
            roles.forEach((r, i) => { runs.push(tag(r, /Global Administrator/.test(r) ? "warn" : "neu", 15)); if (i < roles.length - 1) runs.push(run(" ", { size: 15 })); });
            return [{ t: a.name || "", b: true }, { t: a.upn || "-", size: 16 }, { runs }];
          }),
          colsFrom([0.24, 0.31, 0.45])));
        children.push(gap(100));
      }

      // Overseas sign-ins (shown only when there is something to report)
      const ext = u.externalSignIns || {};
      const extUsers = ext.byUser || [];
      if (extUsers.length > 0) {
        const extTotal = ext.total || 0, extWindow = ext.windowDays || 30;
        children.push(subLabel(`Unexpected overseas sign-ins - last ${extWindow} days`));
        children.push(callout(`${extTotal} successful login${extTotal > 1 ? "s" : ""} from unexpected locations by ${ext.uniqueUsers} user${ext.uniqueUsers > 1 ? "s" : ""}. Australia, New Zealand and Malaysia are excluded as expected locations.`, "warn"));
        children.push(gap(100));
        children.push(dataTable(["Name", "Email", "Country / Territory", "Logins"],
          extUsers.map(e => [{ t: e.name || String(e.upn || "").split("@")[0], b: true }, { t: e.upn || "", size: 16 }, (e.countries || []).join(", "), String(e.eventCount || 0)]),
          colsFrom([0.24, 0.30, 0.30, 0.16]), { num: [3] }));
        children.push(gap(100));
      }
    }

    // ── SharePoint ────────────────────────────────────────────────────────────
    const sp = d.sharepoint || {};
    if ((sp.siteCount || 0) > 0 || sp.error) {
      children.push(...sectionHeaderBlock("SharePoint / MS Teams", C.TEAL));
      children.push(lead("Overview of SharePoint Online sites and storage usage for this tenant. Site counts are sourced from the Microsoft 365 usage reports and exclude personal OneDrive sites."));
      if (sp.error) {
        children.push(callout("SharePoint data unavailable - check that Reports.Read.All permission is granted.", "warn"));
        children.push(gap(100));
      } else {
        const spKpis = [
          k(sp.siteCount || 0,    "Total sites",         "info"),
          k(sp.groupCount || 0,   "M365 Group sites",    "info"),
          k(sp.commCount || 0,    "Communication sites", "neu"),
          k(sp.channelCount || 0, "Teams channels",      "neu"),
          ...((sp.classicCount || 0) > 0 ? [k(sp.classicCount, "Classic sites", "neu")] : []),
          ...((sp.otherCount || 0) > 0 ? [k(sp.otherCount, "Other (system sites)", "neu")] : []),
          k(`${(sp.totalUsedGB || 0).toLocaleString("en-NZ")} GB`, "Storage used", (sp.totalUsedGB || 0) > 3000 ? "bad" : "neu"),
          k(sp.inactiveSiteCount || 0, "Inactive Sites and Channels 180 days +", (sp.inactiveSiteCount || 0) > 0 ? "warn" : "good"),
          k(sp.securityGroupCount ?? "N/A", "Security groups", "info"),
        ];
        children.push(kpiGrid(spKpis, 3));
        children.push(gap(120));
        if ((sp.inactiveSiteCount || 0) > 0) {
          const n = sp.inactiveSiteCount;
          children.push(callout("Review for archiving or deletion. Note: archived sites may appear in this count.", "warn",
            `${n} site${n > 1 ? "s have" : " has"} had no recorded file activity in 180+ days.`));
          children.push(gap(100));
        }
      }
    }

    // ── Ticketing ─────────────────────────────────────────────────────────────
    const hasTickets = manual.ticketsOpened || manual.ticketsClosed || manual.ticketsPending || manual.avgResponse || manual.avgResolution || manual.p1;
    if (hasTickets) {
      children.push(...sectionHeaderBlock("Ticketing & Support", C.TEAL));
      const tKpis = [];
      if (manual.ticketsOpened)  tKpis.push(k(manual.ticketsOpened,  "Opened",  "neu"));
      if (manual.ticketsClosed)  tKpis.push(k(manual.ticketsClosed,  "Closed",  "good"));
      if (manual.ticketsPending) tKpis.push(k(manual.ticketsPending, "Pending", parseInt(manual.ticketsPending) > 5 ? "warn" : "neu"));
      if (manual.avgResponse)    tKpis.push(k(manual.avgResponse,    "Avg response time",   "neu"));
      if (manual.avgResolution)  tKpis.push(k(manual.avgResolution,  "Avg resolution time", "neu"));
      if (tKpis.length > 0) { children.push(kpiGrid(tKpis, Math.min(tKpis.length, 3))); children.push(gap(140)); }
      children.push(callout(manual.p1 ? manual.p1 : "No P1 incidents recorded this period.", manual.p1 ? "bad" : "good", manual.p1 ? "P1 incident:" : undefined));
      children.push(gap(100));
    }

    // ── Recommendations ───────────────────────────────────────────────────────
    if (manual.personal || manual.improvements || manual.cost || manual.roadmap) {
      children.push(...sectionHeaderBlock("Recommendations & Opportunities", C.ORANGE));
      if (manual.personal)     { children.push(personalCard(manual.personal)); children.push(gap(140)); }
      if (manual.improvements) children.push(...recCard("Suggested improvements", C.BLUE,   manual.improvements));
      if (manual.cost)         children.push(...recCard("Cost-saving opportunities", C.TEAL, manual.cost));
      if (manual.roadmap)      children.push(...recCard("Tech roadmap", C.PURPLE,            manual.roadmap));
    }

    // ── Security Risk Register (bottom of report) ──────────────────────────────
    if (risks.length > 0) {
      children.push(...sectionHeaderBlock("Security Risk Register", C.RED));
      children.push(lead("The following risks have been identified based on the data collected from this tenant. Items are ranked by severity. High-severity items require prompt attention."));
      const SEV = { high: ["High", "bad"], medium: ["Medium", "warn"], low: ["Low", "info"] };
      children.push(dataTable(["Risk", "Area", "Rating", "Recommended action"],
        risks.map(r => {
          const [lbl, tone] = SEV[r.severity] || [String(r.severity || ""), "neu"];
          return [{ t: r.finding || "", b: true }, { t: r.area || "" }, { runs: [tag(lbl, tone, 17)] }, { t: r.action || "" }];
        }),
        colsFrom([0.34, 0.16, 0.10, 0.40])));
      children.push(gap(100));
    }

    // ── Footer note ───────────────────────────────────────────────────────────
    const footerImg = hasHeader ? [new ImageRun({ data: headerLogo, transformation: { width: 62, height: 16 }, type: "png" })] : [];
    children.push(new Paragraph({
      children: [...footerImg, run(`${footerImg.length ? "    " : ""}Prepared: ${today}  ·  Microsoft Intune & Graph API  ·  Confidential`, { size: 16, color: C.LGRAY })],
      spacing: { before: 320, after: 0 },
      border: { top: bdr(C.BORDER, 6), bottom: none(), left: none(), right: none() },
    }));

    // ── Document ──────────────────────────────────────────────────────────────
    const doc = new Document({
      numbering: { config: [{ reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] }] },
      styles: {
        default: { document: { run: { font: "Arial", size: 20 } } },
        paragraphStyles: [
          { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 36, bold: true, font: "Arial", color: C.BLUE }, paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 } },
          { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: "Arial", color: C.DARK }, paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 1 } },
        ],
      },
      sections: [{
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 } } },
        // No running page header -- the cover already carries the logo and client name.
        footers: { default: new Footer({ children: [new Paragraph({
          children: [run("Integricity Technology  ·  Confidential    ", { size: 16, color: C.LGRAY }), new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: C.LGRAY })],
          alignment: AlignmentType.CENTER,
          border: { top: bdr(C.BORDER, 4), bottom: none(), left: none(), right: none() },
        })] }) },
        children,
      }],
    });

    const buffer = await Packer.toBuffer(doc);
    return {
      statusCode: 200,
      headers: CORS,
      body: JSON.stringify({
        base64: buffer.toString("base64"),
        filename: client.replace(/[^a-zA-Z0-9]/g, "_") + "_" + new Date(from).toLocaleDateString("en-NZ", {month:"long",year:"numeric"}).replace(/ /g,"_") + "_IT_Report.docx",
      }),
    };
  } catch (e) {
    return { statusCode: 500, headers: CORS, body: JSON.stringify({ error: "Export error: " + e.message }) };
  }
};
