"use strict";

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, TabStopType,
  LevelFormat, Footer, PageNumber, ImageRun,
} = require("docx");

// ── Brand colours ────────────────────────────────────────────────────────────
const C = {
  BLUE:"0D7CC4", PURPLE:"8C0C6E", RED:"E61B1F", ORANGE:"E06400", TEAL:"1A8070",
  DARK:"1F2937", GRAY:"6B7280", LGRAY:"9CA3AF",
  BGRAY:"F8F9FC", BGRAY2:"F3F4F6", FAFAFA:"FAFAFA", BORDER:"E5E7EB", WHITE:"FFFFFF",
  SUCCESS:"0F7A3C", SUCBG:"DCFCE7",
  WARN:"D97706", WARN_BG:"FFF7ED", WARN_TXT:"92400E", WARN_PILL:"C2410C",
  ERR:"C0152A", ERR_BG:"FEE2E2", ERR_TXT:"991B1B",
  INFO:"2563EB", INFO_BG:"EFF6FF", INFO_TXT:"1E3A8A",
};

const PW = 9360; // page content width in DXA

// ── Border helpers ────────────────────────────────────────────────────────────
const bdr  = (color = C.BORDER, size = 4) => ({ style: BorderStyle.SINGLE, size, color });
const none = () => ({ style: BorderStyle.NONE, size: 0, color: C.WHITE });
const allB = (color = C.BORDER, size = 4) => ({ top: bdr(color,size), bottom: bdr(color,size), left: bdr(color,size), right: bdr(color,size) });
const noB  = () => ({ top: none(), bottom: none(), left: none(), right: none() });
const botB = (color, size = 6) => ({ top: none(), bottom: bdr(color,size), left: none(), right: none() });

// ── Text & paragraph helpers ─────────────────────────────────────────────────
const run = (text, { size=20, bold=false, color="333333", font="Arial", italics=false } = {}) =>
  new TextRun({ text: String(text), font, size, bold, color, italics });

const para = (children, { before=0, after=0, align=AlignmentType.LEFT, border } = {}) => {
  const opts = { children: Array.isArray(children) ? children : [children], alignment: align, spacing: { before, after } };
  if (border) opts.border = border;
  return new Paragraph(opts);
};

const spacer = (h = 120) => new Paragraph({ children: [], spacing: { before: h, after: 0 } });

// Coloured accent bar + bold dark title + THIS MONTH badge — matches v11 PDF.
// This used to be a Table (a coloured-square cell + title cell + badge cell),
// with `keepNext` set on the title cell's own paragraph. That keepNext was a
// no-op for our purposes: Word only chains "keep with next" between sibling
// paragraphs, so a paragraph nested inside a table cell can't pull the WHOLE
// TABLE forward to stay with whatever follows it in the document body —
// confirmed by an isolated repro where the header+rule were left orphaned at
// the page foot while the section's content jumped to the next page alone.
// A single top-level paragraph *can* chain via keepNext, so the coloured
// square is now a left-border accent (same visual effect) on one paragraph
// that also carries the title and the right-aligned badge via a tab stop.
const sectionHeader = (title, color) => new Paragraph({
  tabStops: [{ type: TabStopType.RIGHT, position: PW }],
  border: { left: bdr(color, 64), top: none(), bottom: none(), right: none() },
  indent: { left: 200 },
  spacing: { before: 40, after: 40 },
  keepNext: true,
  children: [
    run(title, { size: 32, bold: true, color: C.DARK }),
    run('\tTHIS MONTH', { size: 14, bold: true, color: C.LGRAY }),
  ],
});
const sectionHeaderBlock = (title, color) => [
  new Paragraph({ children: [], spacing: { before: 400, after: 0 } }),
  sectionHeader(title, color),
  new Paragraph({
    children: [],
    border: { top: none(), bottom: bdr(C.BORDER, 4), left: none(), right: none() },
    spacing: { before: 60, after: 120 },
    keepNext: true,
  }),
];

const subLabel = text => new Paragraph({
  children: [run(text.toUpperCase(), { size: 15, bold: true, color: C.GRAY })],
  spacing: { before: 200, after: 80 },
  border: { top: none(), bottom: none(), left: bdr(C.BORDER, 20), right: none() },
  indent: { left: 160 },
  keepNext: true,
});

// ── Shared cell builder ───────────────────────────────────────────────────────
const cell = (children, { width, bg = C.WHITE, borders, margins = { top:80, bottom:80, left:120, right:120 } } = {}) =>
  new TableCell({
    borders, shading: { fill: bg, type: ShadingType.CLEAR },
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    margins,
    children: Array.isArray(children) ? children : [children],
  });

// ── KPI grid ─────────────────────────────────────────────────────────────────
const KPI_COLORS = {
  good: { val: C.SUCCESS, accent: C.SUCCESS },
  warn: { val: C.WARN,    accent: C.WARN },
  bad:  { val: C.ERR,     accent: C.ERR },
  info: { val: C.BLUE,    accent: C.BLUE },
  neu:  { val: C.DARK,    accent: "CCCCCC" },
};

function kpiGrid(items) {
  const cols = Math.min(items.length, 5);
  // cantSplit prevents rows breaking across pages
  const colW = Math.floor(PW / cols);
  const rows = [];

  for (let i = 0; i < items.length; i += cols) {
    const chunk = [...items.slice(i, i + cols)];
    while (chunk.length < cols) chunk.push({ label: "", value: "" });
    const clrs = chunk.map(it => KPI_COLORS[it.cls] || KPI_COLORS.neu);

    // Accent strip row
    rows.push(new TableRow({ cantSplit: true, children: chunk.map((_, j) => cell(
      para([]),
      { width: colW, bg: C.BGRAY, borders: { top: bdr(clrs[j].accent, 12), bottom: none(), left: none(), right: none() }, margins: { top:0, bottom:0, left:0, right:0 } }
    ))}));

    // Value + label row
    rows.push(new TableRow({ cantSplit: true, children: chunk.map((it, j) => cell(
      [
        para([run(String(it.value || ""), { size: 40, bold: true, color: clrs[j].val })], { align: AlignmentType.CENTER, after: 40 }),
        para([run((it.label || "").toUpperCase(), { size: 14, color: C.GRAY })], { align: AlignmentType.CENTER }),
      ],
      { width: colW, bg: C.BGRAY, borders: { top: none(), bottom: bdr(C.BORDER), left: bdr(C.BORDER), right: bdr(C.BORDER) }, margins: { top:120, bottom:120, left:100, right:100 } }
    ))}));
  }
  return new Table({ width: { size: PW, type: WidthType.DXA }, columnWidths: Array(cols).fill(colW), rows });
}

const k = (value, label, cls) => ({ value, label, cls });

// ── Callout box ───────────────────────────────────────────────────────────────
const CALLOUT = {
  warn: { bg: C.WARN_BG, accent: C.WARN,    text: C.WARN_TXT },
  bad:  { bg: C.ERR_BG,  accent: C.ERR,     text: C.ERR_TXT },
  good: { bg: C.SUCBG,   accent: C.SUCCESS,  text: "14532D" },
  info: { bg: C.INFO_BG, accent: C.INFO,    text: C.INFO_TXT },
};

function callout(text, type) {
  const cfg = CALLOUT[type] || { bg: C.BGRAY, accent: C.BORDER, text: C.DARK };
  return new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: [160, PW - 160],
    rows: [new TableRow({ children: [
      cell(para([]), { width: 160, bg: cfg.accent, borders: noB(), margins: { top:0, bottom:0, left:0, right:0 } }),
      cell(para([run(text, { size: 19, color: cfg.text })]), { width: PW-160, bg: cfg.bg, borders: { top: none(), bottom: none(), left: none(), right: bdr(C.BORDER) }, margins: { top:80, bottom:80, left:160, right:120 } }),
    ]})]
  });
}

// ── Policy table ──────────────────────────────────────────────────────────────
const STATUS_CFG = {
  enforced:      { color: C.SUCCESS, text: "Enforced" },
  "report-only": { color: C.WARN_PILL, text: "Report-only" },
  "not found":   { color: C.ERR,     text: "Not found" },
  disabled:      { color: C.GRAY,    text: "Disabled" },
  enabled:       { color: C.SUCCESS, text: "Enabled" },
  "enabled-warn":{ color: C.WARN_PILL,text: "Enabled - review" },
};

function policyTable(rows) {
  // Matches mockup: name left, coloured status pill right, alternating row bg
  const nameW = Math.floor(PW * 0.68);
  const pillW = PW - nameW;
  const PILL_CFG = {
    "enforced":      { bg: "DCFCE7", tc: "166534", label: "Enforced" },
    "report-only":   { bg: "FFF7ED", tc: "92400E", label: "Report-only" },
    "not found":     { bg: "FEE2E2", tc: "991B1B", label: "Not found" },
    "disabled":      { bg: "F3F4F6", tc: "6B7280", label: "Disabled" },
    "enabled":       { bg: "DCFCE7", tc: "166534", label: "Enabled" },
    "enabled-warn":  { bg: "FFF7ED", tc: "92400E", label: "Enabled - review" },
  };
  return new Table({ width: { size: PW, type: WidthType.DXA }, columnWidths: [nameW, pillW],
    rows: rows.map((r, i) => {
      const bg = i % 2 === 0 ? C.BGRAY2 : C.WHITE;
      const pc = PILL_CFG[r.status] || { bg: "F3F4F6", tc: "6B7280", label: r.status };
      return new TableRow({ children: [
        cell(para([run(r.label, { size: 19, color: C.DARK })]),
          { width: nameW, bg, borders: { top: none(), bottom: bdr(C.BORDER), left: bdr(C.BORDER), right: none() },
            margins: { top: 100, bottom: 100, left: 140, right: 100 } }),
        cell(para([run(pc.label, { size: 18, bold: true, color: pc.tc })], { align: AlignmentType.CENTER }),
          { width: pillW, bg: pc.bg, borders: allB(pc.tc, 2),
            margins: { top: 100, bottom: 100, left: 80, right: 80 } }),
      ]});
    })
  });
}

// ── Risk table ────────────────────────────────────────────────────────────────
function riskTable(risks) {
  const cols = [1100, 1800, 3300, 3160];
  const SEV = { high: { color: C.ERR, bg: C.ERR_BG }, medium: { color: C.WARN_PILL, bg: C.WARN_BG }, low: { color: C.SUCCESS, bg: C.SUCBG } };
  const hdr = new TableRow({ children:
    ["Severity","Area","Finding","Recommended action"].map((t, i) =>
      cell(para([run(t, { size: 17, bold: true, color: C.WHITE })]), { width: cols[i], bg: C.DARK, borders: allB(C.DARK) })
    )
  });
  const dataRows = risks.map((r, i) => {
    const bg = i % 2 === 0 ? C.WHITE : C.FAFAFA;
    const sc = SEV[r.severity] || { color: C.GRAY, bg };
    return new TableRow({ children: [
      cell(para([run(r.severity.charAt(0).toUpperCase() + r.severity.slice(1), { size: 18, bold: true, color: sc.color })]), { width: cols[0], bg: sc.bg, borders: allB(C.BORDER) }),
      cell(para([run(r.area,    { size: 18, color: C.GRAY })]), { width: cols[1], bg, borders: allB(C.BORDER) }),
      cell(para([run(r.finding, { size: 18 })]),                { width: cols[2], bg, borders: allB(C.BORDER) }),
      cell(para([run(r.action,  { size: 18, color: C.GRAY })]), { width: cols[3], bg, borders: allB(C.BORDER) }),
    ]});
  });
  return new Table({ width: { size: PW, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...dataRows] });
}

// ── Rec card & personal message ───────────────────────────────────────────────
function recCard(title, titleColor, text) {
  return [
    new Table({
      width: { size: PW, type: WidthType.DXA }, columnWidths: [PW],
      rows: [new TableRow({ children: [cell(
        [
          para([run(title.toUpperCase(), { size: 17, bold: true, color: titleColor })], { after: 60 }),
          para([run(text, { size: 19, color: C.DARK })]),
        ],
        { width: PW, bg: C.BGRAY, borders: allB(C.BORDER), margins: { top:120, bottom:120, left:160, right:160 } }
      )]})]
    }),
    spacer(80),
  ];
}

function personalCard(text) {
  return new Table({
    width: { size: PW, type: WidthType.DXA }, columnWidths: [PW],
    rows: [new TableRow({ children: [cell(
      para([run(text, { size: 20, italics: true, color: "3730A3" })]),
      { width: PW, bg: C.INFO_BG, borders: { top: bdr(C.BLUE, 6), bottom: bdr(C.PURPLE, 6), left: bdr(C.BLUE, 12), right: bdr(C.BORDER) }, margins: { top:140, bottom:140, left:200, right:160 } }
    )]})]
  });
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
    const { client, from, to, preparer, today, iData: d, manual } = payload;
    if (!client || !from || !to || !d) throw new Error("Missing required report data");
    const fmt = dt => new Date(dt).toLocaleDateString("en-NZ", { day: "numeric", month: "long", year: "numeric" });

    const coverLogo  = Buffer.from(payload.fullLogoTransparent || "", "base64");
    const headerLogo = Buffer.from(payload.fullLogoHeader || "", "base64");
    const hasLogo    = coverLogo.length > 100;
    const hasHeader  = headerLogo.length > 100;

    const children = [];

    // ── Cover ────────────────────────────────────────────────────────────────
    // Navy cover matching PDF v11 style
    const monthLabel = new Date(from).toLocaleDateString("en-NZ",{month:"long",year:"numeric"});
    const NAVY = "0F172A";
    const NAVY2 = "1A2744";
    const MUTED = "94A3B8";
    if (hasLogo) {
      children.push(new Table({
        width: { size: PW, type: WidthType.DXA }, columnWidths: [PW],
        rows: [
          // Logo + badge row
          new TableRow({ children: [cell(
            [
              para([new ImageRun({ data: coverLogo, transformation: { width: 186, height: 48 }, type: "png" })], { after: 240 }),
              para([run(client, { size: 56, bold: true, color: C.WHITE })], { after: 80 }),
              para([run(monthLabel, { size: 22, color: MUTED })], { after: 60 }),
            ],
            { width: PW, bg: NAVY, borders: noB(), margins: { top:240, bottom:0, left:200, right:200 } }
          )]}),
          // Separator + meta row
          new TableRow({ children: [cell(
            [
              para([run("", { size: 8 })], { after: 0, border: botB(C.BORDER, 4) }),
              para([
                run(`Account manager: `, { size: 17, color: MUTED }),
                run(`${preparer || "Integricity Technology"}`, { size: 17, bold: true, color: C.WHITE }),
                run(`    Generated: `, { size: 17, color: MUTED }),
                run(today, { size: 17, bold: true, color: C.WHITE }),
              ], { after: 0 }),
            ],
            { width: PW, bg: NAVY, borders: noB(), margins: { top:80, bottom:200, left:200, right:200 } }
          )]}),
        ]
      }));
    } else {
      children.push(para([run(`Monthly IT Report - ${client}`, { size: 52, bold: true, color: C.BLUE })], { after: 80 }));
      children.push(para([run(`${fmt(from)} - ${fmt(to)}  ·  Prepared: ${today}${preparer ? " · " + preparer : ""}`, { size: 20, color: C.GRAY })], { after: 200 }));
    }
    // No extra spacer here -- sectionHeaderBlock() (used by every section
    // below, including Executive Summary/Device & Asset Management) already
    // adds its own 400-twip leading gap, and the cover table above already
    // has its own bottom margin. Stacking a third gap on top produced a
    // visibly oversized blank band between the cover and the first section.

    // ── Executive Summary ────────────────────────────────────────────────────
    if (manual.overview || manual.highlights || manual.concerns || manual.projects) {
      children.push(...sectionHeaderBlock("Executive Summary", C.BLUE));
      if (manual.overview) children.push(para([run(manual.overview, { size: 20 })], { after: 100 }));
      if (manual.highlights) { children.push(callout("Highlights: " + manual.highlights, "good")); children.push(spacer(140)); }
      if (manual.concerns)   { children.push(callout("Concerns: "   + manual.concerns,   "warn")); children.push(spacer(140)); }
      if (manual.projects)   { children.push(callout("Projects: "   + manual.projects,   "info")); children.push(spacer(140)); }
    }

    // ── Devices ──────────────────────────────────────────────────────────────
    children.push(...sectionHeaderBlock("Device & Asset Management", C.BLUE));
    children.push(para([run("Summary of managed devices enrolled in Microsoft Intune, including compliance status, operating system breakdown, and device health indicators.", {size:19,color:"374151"})], {after:120}));
    const devKpis = [
      k(d.total, "Total devices", "info"),
      k(d.comp.compliant, "Compliant", "good"),
      k(d.comp.noncompliant, "Non-Compliant", d.comp.noncompliant > 0 ? "bad" : "good"),
      k(d.comp.unknown||0, "Not evaluated", "warn"),
      k(d.staleCount||0, "Last Checked In 90d+", (d.staleCount||0) > 0 ? "warn" : "good"),
    ];
    children.push(kpiGrid(devKpis)); children.push(spacer(140));
    // OS breakdown - only show non-zero counts
    const osKpis = [
      ...(d.win25h2 > 0 ? [k(d.win25h2, "Windows 11 25H2", "good")] : []),
      ...(d.win24h2 > 0 ? [k(d.win24h2, "Windows 11 24H2", "good")] : []),
      ...(d.win11   > 0 ? [k(d.win11,   "Windows 11",      "good")] : []),
      ...(d.win10   > 0 ? [k(d.win10,   "Windows 10",      "warn")] : []),
      ...(d.macOS   > 0 ? [k(d.macOS,   "macOS",           "neu")]  : []),
      ...(d.linux        > 0 ? [k(d.linux,        "Linux",        "neu")] : []),
      ...(d.iosCount    > 0 ? [k(d.iosCount,    "iOS / iPadOS", "neu")] : []),
      ...(d.androidCount > 0 ? [k(d.androidCount, "Android",      "neu")] : []),
    ];
    if (osKpis.length > 0) {
      children.push(subLabel("Operating system"));
      // Add platform prefix symbols to labels for Word (SVG icons not supported in docx)
      const osKpisLabelled = osKpis.map(o => {
        const l = o.label.toLowerCase();
        const prefix = l.includes("windows") ? "⊞ " : l.includes("macos") || l.includes("ios") || l.includes("ipad") ? "⌘ " : l.includes("android") ? "⬡ " : l.includes("linux") ? "🐧 " : "";
        return { ...o, label: prefix + o.label };
      });
      children.push(kpiGrid(osKpisLabelled));
      children.push(spacer(140));
    }
    if (d.comp.noncompliant > 0) {
      children.push(callout(`${d.comp.noncompliant} device${d.comp.noncompliant>1?"s are":" is"} non-compliant with your organisation's policies.`, "bad"));
      children.push(spacer(140));
      if ((d.notCompliantList||[]).length > 0) {
        children.push(subLabel("Non-compliant devices"));
        const ncCols = [Math.floor(PW*0.35), Math.floor(PW*0.40), Math.floor(PW*0.25)];
        children.push(new Table({
          width:{size:PW,type:WidthType.DXA}, columnWidths:ncCols,
          rows:[
            new TableRow({children:[
              cell(para([run("Device",{size:18,bold:true,color:C.WHITE})]),{width:ncCols[0],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Primary user",{size:18,bold:true,color:C.WHITE})]),{width:ncCols[1],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Last seen",{size:18,bold:true,color:C.WHITE})]),{width:ncCols[2],bg:C.DARK,borders:allB(C.DARK)}),
            ]}),
            ...(d.notCompliantList||[]).map((dev,i)=>{
              const bg=i%2===0?C.BGRAY2:C.WHITE;
              const ls=dev.lastSync?new Date(dev.lastSync).toLocaleDateString("en-NZ"):"Never";
              return new TableRow({children:[
                cell(para([run(dev.name||"",{size:18})]),{width:ncCols[0],bg,borders:allB(C.BORDER)}),
                cell(para([run(dev.user||"Unknown",{size:18,color:C.GRAY})]),{width:ncCols[1],bg,borders:allB(C.BORDER)}),
                cell(para([run(ls,{size:18,color:C.GRAY})]),{width:ncCols[2],bg,borders:allB(C.BORDER)}),
              ]});
            }),
          ]
        }));
        children.push(spacer(140));
      }
    }
    children.push(subLabel("Encryption"));
    children.push(kpiGrid([
      k(d.encryption.encrypted,    "Encrypted",    "good"),
      k(d.encryption.notEncrypted, "Not encrypted", d.encryption.notEncrypted>0?"bad":"good"),
    ])); children.push(spacer(140));
    if (d.encryption.notEncrypted > 0) {
      children.push(callout(`${d.encryption.notEncrypted} device${d.encryption.notEncrypted>1?"s are":" is"} not encrypted. This is a significant data protection risk.`, "bad"));
      children.push(spacer(140));
      if ((d.notEncryptedList||[]).length > 0) {
        children.push(subLabel("Unencrypted devices"));
        const encCols = [Math.floor(PW*0.35), Math.floor(PW*0.40), Math.floor(PW*0.25)];
        children.push(new Table({
          width:{size:PW,type:WidthType.DXA}, columnWidths:encCols,
          rows:[
            new TableRow({children:[
              cell(para([run("Device",{size:18,bold:true,color:C.WHITE})]),{width:encCols[0],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Primary user",{size:18,bold:true,color:C.WHITE})]),{width:encCols[1],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("OS",{size:18,bold:true,color:C.WHITE})]),{width:encCols[2],bg:C.DARK,borders:allB(C.DARK)}),
            ]}),
            ...(d.notEncryptedList||[]).map((dev,i)=>{
              const bg=i%2===0?C.BGRAY2:C.WHITE;
              return new TableRow({children:[
                cell(para([run(dev.name||"",{size:18})]),{width:encCols[0],bg,borders:allB(C.BORDER)}),
                cell(para([run(dev.user||"Unknown",{size:18,color:C.GRAY})]),{width:encCols[1],bg,borders:allB(C.BORDER)}),
                cell(para([run(dev.os||"",{size:17,color:C.GRAY})]),{width:encCols[2],bg,borders:allB(C.BORDER)}),
              ]});
            }),
          ]
        }));
        children.push(spacer(140));
      }
    }
      if ((d.lowDisk||[]).length > 0) {
        children.push(callout(`${(d.lowDisk||[]).length} device${(d.lowDisk||[]).length>1?"s have":" has"} low disk space (less than 15% free).`, "warn"));
        children.push(spacer(140));
        children.push(subLabel("Devices with low disk space"));
        const dkCols=[Math.floor(PW*0.45),Math.floor(PW*0.28),Math.floor(PW*0.27)];
        children.push(new Table({width:{size:PW,type:WidthType.DXA},columnWidths:dkCols,rows:[
          new TableRow({children:[
            cell(para([run("Device",{size:18,bold:true,color:C.WHITE})]),{width:dkCols[0],bg:C.DARK,borders:allB(C.DARK)}),
            cell(para([run("Free space",{size:18,bold:true,color:C.WHITE})]),{width:dkCols[1],bg:C.DARK,borders:allB(C.DARK)}),
            cell(para([run("Free (GB)",{size:18,bold:true,color:C.WHITE})]),{width:dkCols[2],bg:C.DARK,borders:allB(C.DARK)}),
          ]}),
          ...(d.lowDisk||[]).map((x,i)=>{
            const bg=i%2===0?C.BGRAY2:C.WHITE;
            return new TableRow({children:[
              cell(para([run(x.name||"",{size:18})]),{width:dkCols[0],bg,borders:allB(C.BORDER)}),
              cell(para([run(x.pct+"%",{size:18,color:C.GRAY})]),{width:dkCols[1],bg,borders:allB(C.BORDER)}),
              cell(para([run(x.gb+" GB",{size:18,color:C.GRAY})]),{width:dkCols[2],bg,borders:allB(C.BORDER)}),
            ]});
          }),
        ]}));
        children.push(spacer(140));
      }

    // ── Security ──────────────────────────────────────────────────────────────
    children.push(...sectionHeaderBlock("Security Posture", C.PURPLE));
    children.push(para([run("Overview of the tenant's security configuration including Microsoft Secure Score, Conditional Access (CA) policy status, authentication methods, and identity protection alerts.", {size:19,color:"374151"})], {after:120}));
    {
      const secKpis = [];
      if (d.score) secKpis.push(k(`${d.score.pct}%`, "Secure Score", d.score.pct >= 70 ? "good" : d.score.pct >= 50 ? "warn" : "bad"));
      if (d.conditionalAccess) {
        secKpis.push(k(d.conditionalAccess.enabled,    "CA Enforced",    "good"));
        secKpis.push(k(d.conditionalAccess.reportOnly, "CA Report-only", d.conditionalAccess.reportOnly > 0 ? "warn" : "good"));
      }
      secKpis.push(k(d.risky||0, "Risky Users", (d.risky||0) > 0 ? "bad" : "good"));
      secKpis.push(k((d.compliancePolicies?.total)||0, "Compliance policies", "neu"));
      secKpis.push(k((d.appProtection?.total)||0, "App protection", (d.appProtection?.total||0) === 0 ? "warn" : "good"));
      if (d.securityDefaults !== null && d.securityDefaults !== undefined) secKpis.push(k(d.securityDefaults ? "On" : "Off", "Security Defaults", d.securityDefaults ? "good" : "neu"));
      children.push(kpiGrid(secKpis)); children.push(spacer(140));
      if (d.score) {
        const scoreMsg = d.score.pct >= 70
          ? `Microsoft Secure Score is ${d.score.pct}% (${d.score.cur}/${d.score.max}) - a healthy baseline.`
          : d.score.pct >= 50
            ? `Microsoft Secure Score is ${d.score.pct}% (${d.score.cur}/${d.score.max}). There is room to improve.`
            : `Microsoft Secure Score is ${d.score.pct}% (${d.score.cur}/${d.score.max}). This requires attention.`;
        children.push(para([run(scoreMsg, { size: 20 })], { after: 120 }));
      }
    }
    children.push(subLabel("Key policy status"));
    const kp = d.keyPolicies || {};
    children.push(policyTable([
      { label: "Legacy Authentication Block",  status: (kp.legacyAuthBlock || {}).status || "not found" },
      { label: "MFA - All Users",              status: (kp.mfaAllUsers    || {}).status || "not found" },
      { label: "Admin Phishing-Resistant MFA", status: (kp.adminMfa       || {}).status || "not found" },
      { label: "Geographic Restriction",       status: (kp.geoBlock        || {}).status || "not found" },
    ])); children.push(spacer(140));
    children.push(subLabel("Authentication methods"));
    const am = d.authMethods || {};
    children.push(policyTable([
      { label: "Microsoft Authenticator", status: am.authAppEnabled ? "enabled" : "disabled" },
      { label: "FIDO2 / Passkeys",        status: am.fido2Enabled   ? "enabled" : "disabled" },
      { label: "SMS sign-in",             status: am.smsEnabled     ? "enabled-warn" : "disabled" },
      { label: "Temporary Access Pass",   status: am.tapEnabled ? (am.tapReusable ? "enabled-warn" : "enabled") : "disabled" },
    ])); children.push(spacer(140));
    if (d.risky > 0) { children.push(callout(`${d.risky} account${d.risky>1?"s are":" is"} flagged as at-risk by Entra ID Protection. Reset passwords immediately.`, "bad")); children.push(spacer(140)); }
    if (d.appProtection?.total === 0) { children.push(callout("No app protection (MAM) policies found. If BYOD access is permitted, this is a gap.", "warn")); children.push(spacer(140)); }

    // ── Patch Status ─────────────────────────────────────────────────────────
    const ps = d.patchStatus || {};
    if (ps.current !== undefined) {
      children.push(...sectionHeaderBlock("Patch Status", C.BLUE));
      children.push(para([run("Patch currency is based on the last Intune check-in date. Devices that have not checked in within 30 days may be running unpatched software.", {size:19,color:"374151"})], {after:120}));
      children.push(kpiGrid([
        k(ps.current||0, "Current (≤30 days)", "good"),
        k(ps.over30||0,  "Over 30 days",       (ps.over30||0)>0?"warn":"good"),
        k(ps.over90||0,  "Over 90 days",       (ps.over90||0)>0?"bad":"good"),
      ]));
      children.push(spacer(140));
      if ((d.patchOver90||[]).length > 0) {
        children.push(callout(`${(d.patchOver90||[]).length} device${(d.patchOver90||[]).length>1?"s have":" has"} not checked in for 90+ days - 3 or more patch cycles behind.`, "bad"));
        children.push(spacer(140));
        children.push(subLabel("Devices with patch version older than 90 days"));
        const p90cols = [Math.floor(PW*0.35), Math.floor(PW*0.32), Math.floor(PW*0.33)];
        children.push(new Table({
          width:{size:PW,type:WidthType.DXA}, columnWidths:p90cols,
          rows:[
            new TableRow({children:[
              cell(para([run("Device",{size:18,bold:true,color:C.WHITE})]),{width:p90cols[0],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Primary user",{size:18,bold:true,color:C.WHITE})]),{width:p90cols[1],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Latest patch version",{size:18,bold:true,color:C.WHITE})]),{width:p90cols[2],bg:C.DARK,borders:allB(C.DARK)}),
            ]}),
            ...(d.patchOver90||[]).map((dev,i) => {
              const bg = i%2===0?C.WHITE:C.BGRAY2;
              return new TableRow({children:[
                cell(para([run(dev.name||"",{size:18})]),{width:p90cols[0],bg,borders:allB(C.BORDER)}),
                cell(para([run(dev.user||"Unknown",{size:18,color:C.GRAY})]),{width:p90cols[1],bg,borders:allB(C.BORDER)}),
                cell(para([run(dev.os||"Unknown",{size:16,color:C.GRAY,italics:true})]),{width:p90cols[2],bg,borders:allB(C.BORDER)}),
              ]});
            }),
          ]
        }));
        children.push(spacer(140));
      }
    }

    // ── User Data ─────────────────────────────────────────────────────────────
    const u = d.users || {};
    if (u.total > 0) {
      children.push(...sectionHeaderBlock("User Data", C.PURPLE));
      children.push(para([run("Summary of M365 licensed users, guest accounts, administrative role holders, and sign-in activity including accounts not used in the last 90 days and logins from outside Australia and New Zealand.", {size:19,color:"374151"})], {after:120}));
      children.push(kpiGrid([
        k(u.total||0,                          "Total Licensed users",       "info"),
        k(u.sharedMailboxes||0,               "Shared mailboxes",     "neu"),
        k(u.guests||0,                         "Guest users",          (u.guests||0)>0?"warn":"neu"),
        k(u.notSignedIn90Licensed||0, "M365 Users - Not Signed in 90 days+", (u.notSignedIn90Licensed||0)>0?"warn":"good"),
        k(u.notSignedIn90Guest||0,    "Guests - not signed in 90d+",    (u.notSignedIn90Guest||0)>0?"warn":"good"),
        k((u.adminRoles||[]).length,           "Admin role holders",   (u.adminRoles||[]).length>0?"warn":"neu"),
        k((u.externalSignIns||{}).total||0, "Overseas sign-ins (30d)", ((u.externalSignIns||{}).uniqueUsers||0)>0?"warn":"good"),
      ]));
      children.push(spacer(140));

      // Not signed in - moved to top of user section
      if ((u.notSignedIn90Licensed||0) > 0) {
        children.push(callout(`${u.notSignedIn90Licensed} M365 licensed user${u.notSignedIn90Licensed>1?"s have":" has"} not signed in for 90+ days. Service accounts and unlicensed accounts are excluded.`, "warn"));
        children.push(spacer(140));
        if ((u.notSignedIn90LicensedList||[]).length > 0) {
          children.push(subLabel("M365 licensed users - not signed in 90+ days"));
          const inCols = [Math.floor(PW*0.28), Math.floor(PW*0.37), Math.floor(PW*0.2), Math.floor(PW*0.15)];
          children.push(new Table({ width:{size:PW,type:WidthType.DXA}, columnWidths:inCols, rows:[
            new TableRow({children:[
              cell(para([run("Name",{size:18,bold:true,color:C.WHITE})]),{width:inCols[0],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Email",{size:18,bold:true,color:C.WHITE})]),{width:inCols[1],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Last sign-in",{size:18,bold:true,color:C.WHITE})]),{width:inCols[2],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Days",{size:18,bold:true,color:C.WHITE})]),{width:inCols[3],bg:C.DARK,borders:allB(C.DARK)}),
            ]}),
            ...u.notSignedIn90LicensedList.map((usr,i) => {
              const bg = i%2===0?C.BGRAY2:C.WHITE;
              const ls = usr.lastSignIn ? new Date(usr.lastSignIn).toLocaleDateString("en-NZ") : "Never";
              return new TableRow({children:[
                cell(para([run(usr.name||"",{size:18})]),{width:inCols[0],bg,borders:allB(C.BORDER)}),
                cell(para([run(usr.upn||"",{size:16,color:C.GRAY,italics:true})]),{width:inCols[1],bg,borders:allB(C.BORDER)}),
                cell(para([run(ls,{size:17,color:C.GRAY})]),{width:inCols[2],bg,borders:allB(C.BORDER)}),
                cell(para([run(usr.daysSince !== null && usr.daysSince !== undefined ? String(usr.daysSince) : "Never",{size:17,color:C.GRAY})]),{width:inCols[3],bg,borders:allB(C.BORDER)}),
              ]});
            }),
          ]}));
          children.push(spacer(140));
        }
      }
      if ((u.notSignedIn90Guest||0) > 0) {
        children.push(callout(`${u.notSignedIn90Guest} guest account${u.notSignedIn90Guest>1?"s have":" has"} not signed in for 90+ days. Review for stale guest access and remove if no longer needed.`, "warn"));
        children.push(spacer(140));
        if ((u.notSignedIn90GuestList||[]).length > 0) {
          children.push(subLabel("Guest accounts - not signed in 90+ days"));
          const gCols = [Math.floor(PW*0.28), Math.floor(PW*0.37), Math.floor(PW*0.2), Math.floor(PW*0.15)];
          children.push(new Table({ width:{size:PW,type:WidthType.DXA}, columnWidths:gCols, rows:[
            new TableRow({children:[
              cell(para([run("Name",{size:18,bold:true,color:C.WHITE})]),{width:gCols[0],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Email",{size:18,bold:true,color:C.WHITE})]),{width:gCols[1],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Last sign-in",{size:18,bold:true,color:C.WHITE})]),{width:gCols[2],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Days",{size:18,bold:true,color:C.WHITE})]),{width:gCols[3],bg:C.DARK,borders:allB(C.DARK)}),
            ]}),
            ...u.notSignedIn90GuestList.map((usr,i) => {
              const bg = i%2===0?C.BGRAY2:C.WHITE;
              const ls = usr.lastSignIn ? new Date(usr.lastSignIn).toLocaleDateString("en-NZ") : "Never";
              return new TableRow({children:[
                cell(para([run(usr.name||"",{size:18})]),{width:gCols[0],bg,borders:allB(C.BORDER)}),
                cell(para([run(usr.upn||"",{size:16,color:C.GRAY,italics:true})]),{width:gCols[1],bg,borders:allB(C.BORDER)}),
                cell(para([run(ls,{size:17,color:C.GRAY})]),{width:gCols[2],bg,borders:allB(C.BORDER)}),
                cell(para([run(usr.daysSince !== null && usr.daysSince !== undefined ? String(usr.daysSince) : "Never",{size:17,color:C.GRAY})]),{width:gCols[3],bg,borders:allB(C.BORDER)}),
              ]});
            }),
          ]}));
          children.push(spacer(140));
        }
      }

      // Licences
      if ((u.licenceSummary||[]).length > 0) {
        children.push(subLabel("Licence assignment"));
        // Card-style grid: 3 columns, each card shows licence name + used/available prominently
        const cardW = Math.floor(PW / 3);
        const cards = u.licenceSummary;
        for (let i = 0; i < cards.length; i += 3) {
          const row = cards.slice(i, i + 3);
          // Pad to 3 if needed
          while (row.length < 3) row.push(null);
          children.push(new Table({
            width:{size:PW,type:WidthType.DXA}, columnWidths:[cardW,cardW,cardW],
            rows:[new TableRow({children: row.map(l => {
              if (!l) return cell(para([],{after:0}),{w:cardW,bg:C.WHITE,borders:noB()});
              const label = l.available > 0 ? `${l.count} / ${l.available}` : String(l.count);
              const pct = l.available > 0 ? l.count / l.available : null;
              const valColor = (l.available > 0 && l.count > l.available) ? C.ERR : C.DARK;
              return cell([
                para([run(l.name,{size:16,color:C.GRAY})],{after:40}),
                para([run(label,{size:28,bold:true,color:valColor})],{after:20}),
                ...(l.available > 0 ? [para([run("licences used",{size:14,color:C.LGRAY})],{after:0})] : []),
              ],{w:cardW,bg:C.BGRAY2,borders:allB(C.BORDER)});
            })})]
          }));
          children.push(spacer(140));
        }
        children.push(spacer(140));
      }

      // Admin roles
      if ((u.adminRoles||[]).length > 0) {
        children.push(subLabel("Users with admin (privileged) roles"));
        const roleCols = [Math.floor(PW*0.42), Math.floor(PW*0.58)];
        children.push(new Table({
          width:{size:PW,type:WidthType.DXA}, columnWidths:roleCols,
          rows:[
            new TableRow({children:[
              cell(para([run("User",{size:18,bold:true,color:C.WHITE})]),{width:roleCols[0],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Roles",{size:18,bold:true,color:C.WHITE})]),{width:roleCols[1],bg:C.DARK,borders:allB(C.DARK)}),
            ]}),
            ...u.adminRoles.map((a,i) => {
              const bg = i%2===0?C.BGRAY2:C.WHITE;
              const roles = (a.roles||[a.role]);
              // Name + UPN stacked in left cell
              const nameCell = cell(
                [
                  para([run(a.name||"",{size:18,bold:true})],{after:20}),
                  para([run(a.upn||"",{size:16,color:C.GRAY,italics:true})],{after:0}),
                ],
                {width:roleCols[0],bg,borders:allB(C.BORDER)}
              );
              // Each role on its own line in right cell
              const roleCell = cell(
                roles.map((r,ri) => para([run(r,{size:18,color:C.GRAY})],{after:ri<roles.length-1?40:0})),
                {width:roleCols[1],bg,borders:allB(C.BORDER)}
              );
              return new TableRow({children:[nameCell, roleCell]});
            }),
          ]
        }));
        children.push(spacer(140));
      }

      // External sign-ins
      const extUsers = (u.externalSignIns||{}).byUser||[];
      const extTotal  = (u.externalSignIns||{}).total||0;
      const extWindow = (u.externalSignIns||{}).windowDays||30;
      if (extUsers.length > 0) {
        children.push(subLabel(`Unexpected overseas sign-ins - last ${extWindow} days`));
        children.push(callout(`${extTotal} successful login${extTotal>1?"s":""} from unexpected locations by ${u.externalSignIns.uniqueUsers} user${u.externalSignIns.uniqueUsers>1?"s":""}. Australia, New Zealand and Malaysia are excluded as expected locations.`, "warn"));
        children.push(spacer(140));
        const extCols = [Math.floor(PW*0.24), Math.floor(PW*0.30), Math.floor(PW*0.30), Math.floor(PW*0.16)];
        children.push(new Table({
          width:{size:PW,type:WidthType.DXA}, columnWidths:extCols,
          rows:[
            new TableRow({children:[
              cell(para([run("Name",{size:18,bold:true,color:C.WHITE})]),{width:extCols[0],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Email",{size:18,bold:true,color:C.WHITE})]),{width:extCols[1],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Country / Territory",{size:18,bold:true,color:C.WHITE})]),{width:extCols[2],bg:C.DARK,borders:allB(C.DARK)}),
              cell(para([run("Logins",{size:18,bold:true,color:C.WHITE})]),{width:extCols[3],bg:C.DARK,borders:allB(C.DARK)}),
            ]}),
            ...extUsers.map((e,i) => {
              const bg = i%2===0?C.BGRAY2:C.WHITE;
              const displayName = e.name || e.upn.split("@")[0];
              return new TableRow({children:[
                cell(para([run(displayName,{size:17})]),{width:extCols[0],bg,borders:allB(C.BORDER)}),
                cell(para([run(e.upn||"",{size:16,color:C.GRAY,italics:true})]),{width:extCols[1],bg,borders:allB(C.BORDER)}),
                cell(para([run((e.countries||[]).join(", "),{size:17,color:C.GRAY})]),{width:extCols[2],bg,borders:allB(C.BORDER)}),
                cell(para([run(String(e.eventCount||0),{size:17,color:C.GRAY})]),{width:extCols[3],bg,borders:allB(C.BORDER)}),
              ]});
            }),
          ]
        }));
        children.push(spacer(140));
      }

    }

    // ── SharePoint ────────────────────────────────────────────────────────────
    const sp = d.sharepoint || {};
    if ((sp.siteCount||0) > 0 || sp.error) {
      children.push(...sectionHeaderBlock("SharePoint / MS Teams", C.TEAL));
      children.push(para([run("Overview of SharePoint Online sites and storage usage for this tenant. Site counts are sourced from the Microsoft 365 usage reports and exclude personal OneDrive sites.", {size:19,color:"374151"})], {after:120}));
      if (sp.error) {
        children.push(callout("SharePoint data unavailable - check that Reports.Read.All permission is granted.", "warn"));
      } else {
        const spKpis = [
          k(sp.siteCount||0,    "Total sites",         "info"),
          k(sp.groupCount||0,   "M365 Group sites",    "neu"),
          k(sp.commCount||0,    "Communication sites", "neu"),
          k(sp.classicCount||0,       "Classic / other",  "neu"),
          ...((sp.m365GroupCount||0) > 0  ? [k(sp.m365GroupCount,  "M365 Groups",      "neu")] : []),
          ...((sp.securityGroupCount||0) > 0 ? [k(sp.securityGroupCount, "Security groups", "neu")] : []),
          k(`${sp.totalUsedGB||0} GB`, "Storage used",   "neu"),
          k(sp.inactiveSiteCount||0,"Inactive 180d+",    (sp.inactiveSiteCount||0)>0?"warn":"good"),
        ];
        if (sp.allocatedGB) spKpis.splice(2,0,k(`${sp.allocatedGB} GB`,"Storage allocated","neu"));
        children.push(kpiGrid(spKpis));
        children.push(spacer(140));
        if (sp.allocatedGB && sp.totalUsedGB) {
          const usedPct = Math.round((sp.totalUsedGB/sp.allocatedGB)*100);
          children.push(callout(`Storage: ${sp.totalUsedGB} GB used of ${sp.allocatedGB} GB allocated (${usedPct}% used).`, usedPct>80?"bad":usedPct>60?"warn":"good"));
          children.push(spacer(140));
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
      if (tKpis.length > 0) { children.push(kpiGrid(tKpis)); children.push(spacer(140)); }
      const slaKpis = [];
      if (manual.avgResponse)   slaKpis.push(k(manual.avgResponse,   "Avg response time",   "neu"));
      if (manual.avgResolution) slaKpis.push(k(manual.avgResolution, "Avg resolution time", "neu"));
      if (slaKpis.length > 0) { children.push(kpiGrid(slaKpis)); children.push(spacer(140)); }
      children.push(callout(manual.p1 ? "P1 incident: " + manual.p1 : "No P1 incidents recorded this period.", manual.p1 ? "bad" : "good"));
      children.push(spacer(140));
    }

    // ── Recommendations ───────────────────────────────────────────────────────
    if (manual.personal || manual.improvements || manual.cost || manual.roadmap) {
      children.push(...sectionHeaderBlock("Recommendations & Opportunities", C.ORANGE));
      if (manual.personal)     { children.push(personalCard(manual.personal)); children.push(spacer(140)); }
      if (manual.improvements) children.push(...recCard("Suggested improvements", C.BLUE,   manual.improvements));
      if (manual.cost)         children.push(...recCard("Cost-saving opportunities", C.TEAL, manual.cost));
      if (manual.roadmap)      children.push(...recCard("Tech roadmap", C.PURPLE,            manual.roadmap));
    }

    // ── Security Risk Register (bottom of report) ──────────────────────────────
    if ((d.risks || []).length > 0) {
      children.push(...sectionHeaderBlock("Security Risk Register", C.RED));
      children.push(para([run("The following risks have been identified based on the data collected from this tenant. Items are ranked by severity. High-severity items require prompt attention.", {size:20,color:"374151"})],{after:120}));
      children.push(riskTable(d.risks));
      children.push(spacer(140));
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
        // No running page header -- the cover page already carries the logo
        // and "Monthly IT Report · client" text, so repeating it as a
        // running header on every page (including the cover itself) was
        // redundant clutter.
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
