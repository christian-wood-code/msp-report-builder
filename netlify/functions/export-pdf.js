"use strict";

// PDF export (server-side) -- restyled to the approved report design: hero
// cover block, headline strip, accent section headers with a "THIS MONTH" tag,
// KPI cards with a coloured top edge, donuts, stacked bars, licence usage rows
// and pill tables. Same {client, from, to, preparer, today, iData, manual,
// fullLogoTransparent} input shape as export-docx.js.
//
// Everything is drawn with jsPDF primitives and Helvetica only: there is no
// Inter font file in this project, and nothing outside WinAnsi is emitted (no
// emoji / symbols). Gradients and shadows from the HTML design are
// approximated (strip gradient, flat cards) -- PDF has no CSS equivalents.

const { jsPDF } = require('jspdf');

const CORS = { 'Access-Control-Allow-Origin': '*' };

// ── Palette: the mockup's light-theme CSS variables, as RGB arrays ──────────
const C = {
  text: [15, 23, 42], muted: [100, 116, 139], line: [229, 233, 242], surf2: [248, 250, 252], white: [255, 255, 255],
  brand: [13, 124, 196], purple: [140, 12, 110], teal: [26, 128, 112], orange: [224, 100, 0], red: [214, 41, 62],
  good: [15, 157, 107], goodBg: [231, 246, 239],
  warn: [199, 119, 0], warnBg: [255, 244, 224],
  bad: [214, 41, 62], badBg: [253, 235, 238],
  info: [13, 124, 196], infoBg: [230, 242, 251],
  neu: [71, 85, 105], neuBg: [238, 241, 246],
};
const TONE = {
  good: { fg: C.good, bg: C.goodBg }, warn: { fg: C.warn, bg: C.warnBg }, bad: { fg: C.bad, bg: C.badBg },
  info: { fg: C.info, bg: C.infoBg }, neu: { fg: C.neu, bg: C.neuBg },
};
const KPI_TOP = { good: C.good, warn: C.warn, bad: C.bad, info: C.info, neu: C.line };
const KPI_VAL = { good: C.good, warn: C.warn, bad: C.bad, info: C.info, neu: C.text };

const MON_S = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const MON_L = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const ymd = s => { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? { y: +m[1], m: +m[2] - 1, d: +m[3] } : null; };
const fmtShort = s => { const p = ymd(s); return p ? `${p.d} ${MON_S[p.m]} ${p.y}` : ''; };
const fmtLong = s => { const p = ymd(s); return p ? `${p.d} ${MON_L[p.m]} ${p.y}` : ''; };
// Keep all drawn text inside WinAnsi so standard Helvetica never emits junk.
const clean = s => String(s == null ? '' : s)
  .replace(/[–—−]/g, '-').replace(/[‘’]/g, "'").replace(/[“”]/g, '"')
  .replace(/…/g, '...').replace(/≤/g, '<=').replace(/≥/g, '>=').replace(/ /g, ' ')
  .replace(/[^\x20-\x7E\xA0-\xFF]/g, '');
const mix = (a, b, t) => a.map((v, i) => Math.round(v + (b[i] - v) * t));
const num = n => Number(n).toLocaleString('en-US');
const k = (value, label, cls) => ({ value, label, cls });

function buildPdfDoc({ client, from, to, preparer, today, iData: d, manual = {}, fullLogoTransparent }) {
  const doc = new jsPDF({ unit: 'pt', format: 'a4' });
  const PW = doc.internal.pageSize.getWidth();
  const PH = doc.internal.pageSize.getHeight();
  const M = 38;
  const CW = PW - 2 * M;
  const BOT = PH - 46; // content bottom (footer lives below)
  let y = 0;

  // ── low-level helpers ─────────────────────────────────────────────────────
  const fill = c => doc.setFillColor(c[0], c[1], c[2]);
  const stroke = c => doc.setDrawColor(c[0], c[1], c[2]);
  const F = (style, size, color) => { doc.setFont('helvetica', style); doc.setFontSize(size); if (color) doc.setTextColor(color[0], color[1], color[2]); };
  const W = (s, style, size) => { doc.setFont('helvetica', style); doc.setFontSize(size); return doc.getTextWidth(s); };
  const alpha = a => doc.setGState(new doc.GState({ opacity: a, 'stroke-opacity': a }));
  const ensure = need => { if (y + need > BOT) { doc.addPage(); y = M; return true; } return false; };
  const box = (x, yy, w, h, r = 8, fc = C.white, bc = C.line) => {
    fill(fc); stroke(bc); doc.setLineWidth(0.7); doc.roundedRect(x, yy, w, h, r, r, 'FD');
  };
  const spaced = (text, x, yy, size, color, cs, opt) => {
    F('bold', size, color);
    if (opt && opt.align === 'right') x -= doc.getTextWidth(text) + cs * (text.length - 1);
    doc.setCharSpace(cs); doc.text(text, x, yy); doc.setCharSpace(0);
  };

  // ── pills ─────────────────────────────────────────────────────────────────
  const pillW = (t, size) => W(clean(t), 'bold', size) + size * 1.8;
  const pill = (t, tone, x, ymid, size = 7.5) => {
    t = clean(t);
    const w = pillW(t, size), h = size + 6, tn = TONE[tone] || TONE.neu;
    fill(tn.bg); doc.roundedRect(x, ymid - h / 2, w, h, h / 2, h / 2, 'F');
    F('bold', size, tn.fg); doc.text(t, x + w / 2, ymid + size * 0.35, { align: 'center' });
    return w;
  };
  // Lay a list of [text, tone] pills out in wrapping rows within `width`.
  const pillRows = (pills, width, size, gap = 4) => {
    const rows = [[]]; let x = 0;
    pills.forEach(p => {
      const w = pillW(p[0], size);
      if (x > 0 && x + w > width) { rows.push([]); x = 0; }
      rows[rows.length - 1].push({ p, x, w }); x += w + gap;
    });
    return { rows, h: rows.length * (size + 6) + (rows.length - 1) * 3 };
  };

  // ── rich (mixed bold/regular) wrapped text ────────────────────────────────
  const richLines = (segs, width, size) => {
    const words = [];
    segs.forEach(s => String(s.t).split(/\s+/).filter(Boolean).forEach(w => words.push({ w, b: !!s.b })));
    const sp = W(' ', 'normal', size);
    const lines = []; let cur = [], cw = 0;
    words.forEach(wd => {
      const w = W(wd.w, wd.b ? 'bold' : 'normal', size);
      const x = cur.length ? cw + sp : 0;
      if (cur.length && x + w > width) { lines.push(cur); cur = []; cw = 0; }
      const xx = cur.length ? cw + sp : 0;
      cur.push({ w: wd.w, b: wd.b, x: xx }); cw = xx + w;
    });
    if (cur.length) lines.push(cur);
    return lines;
  };
  const drawRich = (lines, x, y0, lh, size, color) => {
    lines.forEach((ln, i) => ln.forEach(wd => { F(wd.b ? 'bold' : 'normal', size, color); doc.text(wd.w, x + wd.x, y0 + i * lh); }));
  };

  // ── section header: rounded accent tile (no icon) + title + THIS MONTH tag ─
  const sectionHeader = (title, color, intro, minFollow = 110) => {
    let introLines = [];
    if (intro) { F('normal', 9, C.muted); introLines = doc.splitTextToSize(intro, CW); }
    const need = 24 + 12 + (introLines.length ? introLines.length * 12 + 8 : 0) + minFollow;
    ensure(need + (y > M + 1 ? 22 : 0));
    if (y > M + 1) y += 22;
    fill(mix(color, C.white, 0.86)); doc.roundedRect(M, y, 24, 24, 7, 7, 'F');
    fill(color); doc.roundedRect(M + 8, y + 8, 8, 8, 2.5, 2.5, 'F');
    F('bold', 16, C.text); doc.text(clean(title), M + 34, y + 17.5);
    spaced('THIS MONTH', PW - M, y + 16, 7, C.muted, 0.8, { align: 'right' });
    y += 24 + 12;
    if (introLines.length) {
      F('normal', 9, C.muted); doc.text(introLines, M, y + 2, { lineHeightFactor: 1.35 });
      y += introLines.length * 12 + 8;
    }
  };

  // small uppercase sub-heading (the mockup's h3)
  const h3 = (text, follow = 84) => {
    if (follow) ensure(24 + follow);
    y += 6;
    spaced(clean(text).toUpperCase(), M, y + 7, 7.5, C.muted, 0.7);
    y += 18;
  };
  const note = (text) => { F('normal', 8, C.muted); const ls = doc.splitTextToSize(clean(text), CW); doc.text(ls, M, y + 1); y += ls.length * 10 + 6; };

  // ── KPI cards (coloured top edge) ─────────────────────────────────────────
  const kpiCard = (x, yy, w, h, it, vs = 19) => {
    const tone = it.cls || 'neu';
    fill(KPI_TOP[tone]); doc.roundedRect(x, yy, w, 16, 7, 7, 'F');
    fill(C.line); doc.roundedRect(x, yy + 3, w, h - 3, 7, 7, 'F');
    fill(C.white); doc.roundedRect(x + 0.7, yy + 3.7, w - 1.4, h - 4.4, 6.3, 6.3, 'F');
    const vb = yy + 3 + 8 + vs * 0.82;
    F('bold', vs, KPI_VAL[tone]); doc.text(clean(String(it.value)), x + 12, vb);
    F('normal', 8, C.muted);
    doc.text(it.lines, x + 12, vb + 12.5, { lineHeightFactor: 1.15 });
  };
  const kpiHeight = (lines, vs) => 3 + 8 + vs * 0.82 + 12.5 + (lines - 1) * 9.6 + 11;
  const kpiGrid = (items, cols = 4, vs = 19) => {
    const gap = 10, w = (CW - gap * (cols - 1)) / cols;
    F('normal', 8, C.muted);
    items.forEach(it => { it.lines = doc.splitTextToSize(clean(it.label), w - 22); });
    const rows = [];
    for (let i = 0; i < items.length; i += cols) {
      const chunk = items.slice(i, i + cols);
      rows.push({ chunk, h: kpiHeight(Math.max(...chunk.map(c => c.lines.length)), vs) });
    }
    const total = rows.reduce((s, r) => s + r.h + gap, 0);
    if (total <= 260) ensure(total);
    rows.forEach(r => {
      ensure(r.h + gap);
      r.chunk.forEach((it, j) => kpiCard(M + j * (w + gap), y, w, r.h, it, vs));
      y += r.h + gap;
    });
  };
  const gridHeight = (items, cols = 4, vs = 19) => {
    const gap = 10, w = (CW - gap * (cols - 1)) / cols;
    F('normal', 8, C.muted);
    let tot = 0;
    for (let i = 0; i < items.length; i += cols) {
      const mx = Math.max(...items.slice(i, i + cols).map(it => doc.splitTextToSize(clean(it.label), w - 22).length));
      tot += kpiHeight(mx, vs) + gap;
    }
    return tot;
  };

  // ── callouts ──────────────────────────────────────────────────────────────
  const calloutBox = (h, tone) => {
    const tn = TONE[tone] || TONE.info;
    fill(tn.bg); stroke(mix(tn.fg, C.white, 0.72)); doc.setLineWidth(0.7);
    doc.roundedRect(M, y, CW, h, 8, 8, 'FD');
    fill(tn.fg); doc.circle(M + 15, y + 15.5, 2.8, 'F');
  };
  const callout = (lead, rest, tone = 'info') => {
    const segs = [];
    if (lead) segs.push({ t: clean(lead), b: true });
    if (rest) segs.push({ t: clean(rest), b: false });
    const lines = richLines(segs, CW - 44, 9);
    const h = lines.length * 12.5 + 15;
    ensure(h + 12);
    calloutBox(h, tone);
    drawRich(lines, M + 27, y + 18.5, 12.5, 9, C.text);
    y += h + 12;
  };
  // bold headline left, note right (licence renewals callout)
  const calloutSplit = (lead, rest, tone = 'info') => {
    const iw = CW - 44, lw = iw * 0.36, gapX = 26;
    const L = richLines([{ t: clean(lead), b: true }], lw, 9), R = richLines([{ t: clean(rest), b: false }], iw - lw - gapX, 9);
    const h = Math.max(L.length, R.length) * 12.5 + 15;
    ensure(h + 12);
    calloutBox(h, tone);
    drawRich(L, M + 27, y + 18.5, 12.5, 9, C.text);
    drawRich(R, M + 27 + lw + gapX, y + 18.5, 12.5, 9, C.text);
    y += h + 12;
  };

  // ── donut + stacked bar + legend ─────────────────────────────────────────
  const sector = (cx, cy, R, r, a0, a1, color) => {
    const n = Math.max(2, Math.ceil((a1 - a0) / (Math.PI / 90)));
    const pts = [];
    for (let i = 0; i <= n; i++) { const a = a0 + (a1 - a0) * i / n; pts.push([cx + R * Math.sin(a), cy - R * Math.cos(a)]); }
    for (let i = n; i >= 0; i--) { const a = a0 + (a1 - a0) * i / n; pts.push([cx + r * Math.sin(a), cy - r * Math.cos(a)]); }
    const segs = [];
    for (let i = 1; i < pts.length; i++) segs.push([pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]]);
    fill(color); doc.lines(segs, pts[0][0], pts[0][1], [1, 1], 'F', true);
  };
  const donut = (cx, cy, R, thick, pct, color, big, small) => {
    fill(C.neuBg); doc.circle(cx, cy, R, 'F');
    fill(C.white); doc.circle(cx, cy, R - thick, 'F');
    const p = Math.max(0, Math.min(100, pct));
    if (p >= 99.9) { fill(color); doc.circle(cx, cy, R, 'F'); fill(C.white); doc.circle(cx, cy, R - thick, 'F'); }
    else if (p > 0) sector(cx, cy, R, R - thick, 0, p / 100 * 2 * Math.PI, color);
    F('bold', R > 36 ? 18 : 16, C.text); doc.text(big, cx, cy + 2, { align: 'center' });
    if (small) { F('normal', 6.8, C.muted); doc.text(small, cx, cy + 12, { align: 'center' }); }
  };
  const stackBar = (x, yy, w, h, parts, total) => {
    fill(C.neuBg); doc.roundedRect(x, yy, w, h, h / 2, h / 2, 'F');
    doc.saveGraphicsState();
    doc.roundedRect(x, yy, w, h, h / 2, h / 2, null); doc.clip(); doc.discardPath();
    let xx = x;
    parts.forEach(p => {
      const pw = total > 0 ? p.v / total * w : 0;
      if (pw > 0) { fill(p.c); doc.rect(xx, yy, pw + 0.4, h, 'F'); xx += pw; }
    });
    doc.restoreGraphicsState();
  };
  const legendRows = (items, maxW, size = 8.5) => {
    const rows = [[]]; let x = 0;
    items.forEach(it => {
      const w = 13 + W(clean(it.t), 'normal', size) + 16;
      if (x > 0 && x + w - 16 > maxW) { rows.push([]); x = 0; }
      rows[rows.length - 1].push({ it, x }); x += w;
    });
    return rows;
  };
  const drawLegend = (items, x, yy, maxW, size = 8.5) => {
    const rows = legendRows(items, maxW, size);
    rows.forEach((r, i) => r.forEach(({ it, x: ox }) => {
      fill(it.c); doc.roundedRect(x + ox, yy + i * 14 - 6, 7, 7, 2, 2, 'F');
      F('normal', size, C.muted); doc.text(clean(it.t), x + ox + 13, yy + i * 14);
    }));
    return rows.length * 14;
  };
  const legendHeight = (items, maxW, size) => legendRows(items, maxW, size).length * 14;

  // ── rows card: label left, pill(s) right (never wraps) ───────────────────
  const rowsCardH = (n, titled = true) => 14 + (titled ? 22 : 0) + n * 25 + 3;
  const rowsCard = (x, yy, w, h, title, rows) => {
    box(x, yy, w, h);
    let cy = yy + 14;
    if (title) { spaced(title.toUpperCase(), x + 14, cy + 6, 7.5, C.muted, 0.7); cy += 22; }
    rows.forEach((r, i) => {
      if (i) { stroke(C.line); doc.setLineWidth(0.5); doc.line(x + 14, cy, x + w - 14, cy); }
      F('normal', 9, C.text); doc.text(clean(r.label), x + 14, cy + 15.5);
      let px = x + w - 14;
      for (let j = r.pills.length - 1; j >= 0; j--) { const pw = pillW(r.pills[j][0], 7.5); px -= pw; pill(r.pills[j][0], r.pills[j][1], px, cy + 12.5, 7.5); px -= 4; }
      cy += 25;
    });
  };

  // ── flowing card (rows continue onto the next page with the card re-opened)
  const cardFlow = (drawHead, headH) => {
    let y0 = 0, first = true;
    const open = () => { y0 = y; y += 4; first = true; if (drawHead) { drawHead(); y += headH; } };
    const close = () => { y += 4; stroke(C.line); doc.setLineWidth(0.7); doc.roundedRect(M, y0, CW, y - y0, 8, 8, 'S'); };
    open();
    return {
      row(h, draw, sep = true) {
        if (y + h + 10 > BOT) { close(); doc.addPage(); y = M; open(); }
        if (!first && sep) { stroke(C.line); doc.setLineWidth(0.5); doc.line(M + 12, y, PW - M - 12, y); }
        first = false;
        draw(y, h); y += h;
      },
      close() { close(); y += 12; },
    };
  };

  // ── data table: cols = [{h, w(weight), a:'l'|'r', k:'t'|'b'|'m'|'pill'|'pills'}] ──
  const dataTable = (cols, rows, title) => {
    if (!rows.length) return;
    const PX = 12, GAP = 8, inner = CW - 2 * PX, wsum = cols.reduce((s, c) => s + c.w, 0);
    const cw = cols.map(c => c.w / wsum * inner);
    const cx = []; let acc = M + PX; cw.forEach(w => { cx.push(acc); acc += w; });
    const meas = row => {
      let h = 19;
      const cells = row.map((v, i) => {
        const c = cols[i];
        if (c.k === 'pill' || c.k === 'pills') {
          const pr = pillRows(c.k === 'pill' ? [v] : v, cw[i] - GAP, 7);
          h = Math.max(h, pr.h + 8); return { pr };
        }
        F(c.k === 'b' ? 'bold' : 'normal', 8.5);
        const lines = doc.splitTextToSize(clean(v), cw[i] - GAP);
        h = Math.max(h, lines.length * 10.5 + 8); return { lines };
      });
      return { h, cells };
    };
    const ms = rows.map(meas);
    const drawHead = () => {
      cols.forEach((c, i) => {
        const t = c.h.toUpperCase();
        if (c.a === 'r') spaced(t, cx[i] + cw[i] - GAP, y + 15, 6.8, C.muted, 0.5, { align: 'right' });
        else spaced(t, cx[i], y + 15, 6.8, C.muted, 0.5);
      });
      stroke(C.line); doc.setLineWidth(0.7); doc.line(M + PX - 4, y + 22, PW - M - PX + 4, y + 22);
    };
    ensure((title ? 24 : 0) + 22 + 4 + ms[0].h + 14);
    if (title) h3(title, 0);
    const fl = cardFlow(drawHead, 22);
    ms.forEach(m => {
      fl.row(m.h, (ry, h) => {
        m.cells.forEach((cell, i) => {
          const c = cols[i];
          if (cell.pr) {
            const top = ry + (h - cell.pr.h) / 2;
            cell.pr.rows.forEach((r, ri) => {
              const rw = r.length ? r[r.length - 1].x + r[r.length - 1].w : 0;
              r.forEach(it => {
                const x0 = c.a === 'r' ? cx[i] + cw[i] - GAP - rw : cx[i];
                pill(it.p[0], it.p[1], x0 + it.x, top + ri * 13 + 6.5, 7);
              });
            });
          } else {
            F(c.k === 'b' ? 'bold' : 'normal', 8.5, c.k === 'm' ? C.muted : C.text);
            const base = ry + (h - cell.lines.length * 10.5) / 2 + 8;
            if (c.a === 'r') doc.text(cell.lines, cx[i] + cw[i] - GAP, base, { align: 'right', lineHeightFactor: 1.23 });
            else doc.text(cell.lines, cx[i], base, { lineHeightFactor: 1.23 });
          }
        });
      });
    });
    fl.close();
  };

  // ══════════════════════════════ Hero / cover ═══════════════════════════════
  const retentionText = d.metricsRetained
    ? 'This report is generated directly from your live Microsoft Graph data at the time of generation. '
      + 'Integricity retains only a small set of summary figures from it (such as device counts and Secure Score - '
      + 'no raw tenant data) for up to 13 months, to show month-on-month changes; this document itself will be '
      + 'securely destroyed within 30 days.'
    : 'This report is generated directly from your live Microsoft Graph data at the time of generation and '
      + 'nothing is retained afterward; this document itself will be securely destroyed within 30 days.';
  const disclaimerText = 'This report is a point-in-time, automated summary generated from data available via '
    + 'Microsoft Graph at the time of generation. It is provided as an informational snapshot to support your own '
    + 'evaluation and is not a comprehensive security audit, a compliance certification, or a guarantee of your '
    + "organisation's security posture. Integricity Technology accepts no liability for decisions made on the basis "
    + 'of this report. ' + retentionText
    + ' For a full assessment and remediation plan, contact us to discuss an engagement.';

  F('italic', 7.5);
  const discLines = doc.splitTextToSize(disclaimerText, CW);
  const HERO_H = 176 + discLines.length * 9.4 + 14;
  // navy gradient (left to right) + soft blue glow top-right
  const stops = [[0, [11, 20, 48]], [0.55, [19, 34, 77]], [1, [42, 20, 80]]];
  const gradAt = t => {
    for (let i = 1; i < stops.length; i++) {
      if (t <= stops[i][0]) { const [t0, c0] = stops[i - 1], [t1, c1] = stops[i]; return mix(c0, c1, (t - t0) / (t1 - t0)); }
    }
    return stops[stops.length - 1][1];
  };
  const STEP = 3;
  for (let x = 0; x < PW; x += STEP) { fill(gradAt((x + STEP / 2) / PW)); doc.rect(x, 0, STEP + 0.5, HERO_H, 'F'); }
  alpha(0.05); fill(C.brand);
  for (let r = 150; r >= 20; r -= 9) doc.circle(PW - 40, 8, r, 'F');
  alpha(1);

  const coverLogo = Buffer.from(fullLogoTransparent || '', 'base64');
  if (coverLogo.length > 100) {
    const logoH = 30, logoW = logoH * (960 / 247);
    try { doc.addImage(coverLogo, 'PNG', M, 26, logoW, logoH, undefined, 'SLOW'); } catch (_) { /* ignore malformed logo */ }
  }
  // glass chip (translucent white pill on the hero)
  const glass = (x, yy, parts, dot) => {
    let tw = 0; parts.forEach(p => { tw += W(p.t, p.b ? 'bold' : 'normal', 8); });
    const w = tw + 22 + (dot ? 10 : 0), h = 19;
    alpha(0.13); fill(C.white); doc.roundedRect(x, yy, w, h, h / 2, h / 2, 'F');
    alpha(0.25); stroke(C.white); doc.setLineWidth(0.6); doc.roundedRect(x, yy, w, h, h / 2, h / 2, 'S');
    alpha(1);
    let tx = x + 11;
    if (dot) { fill([232, 240, 255]); doc.circle(tx + 2, yy + h / 2, 2, 'F'); tx += 10; }
    parts.forEach(p => { F(p.b ? 'bold' : 'normal', 8, [232, 240, 255]); doc.text(p.t, tx, yy + 12.4); tx += W(p.t, p.b ? 'bold' : 'normal', 8); });
    return w;
  };
  const pf = ymd(from), pt = ymd(to);
  const monthYear = pf ? `${MON_L[pf.m]} ${pf.y}` : '';
  {
    const parts = [{ t: `Monthly IT report · ${monthYear}`, b: false }];
    let tw = W(parts[0].t, 'normal', 8);
    const w = tw + 32;
    glass(PW - M - w, 30, parts, true);
  }
  F('bold', 30, C.white); doc.text(clean(client), M, 106);
  const period = (pf && pt) ? `Report period: ${fmtLong(from)} - ${fmtLong(to)}` : '';
  if (period) { F('normal', 10.5, [185, 198, 230]); doc.text(period, M, 125); }
  {
    let cx0 = M;
    const chips = [
      [{ t: 'Prepared  ', b: false }, { t: clean(today), b: true }],
      [{ t: 'By  ', b: false }, { t: clean(preparer || 'Integricity Technology'), b: true }],
      [{ t: 'Source  ', b: false }, { t: 'Microsoft Intune & Graph API', b: true }],
    ];
    const hi = (d.risks || []).filter(r => r.severity === 'high').length;
    if (hi > 0) chips.push([{ t: `${hi} high-priority ${hi === 1 ? 'risk' : 'risks'}`, b: false }]);
    chips.forEach(c => { cx0 += glass(cx0, 139, c, false) + 8; });
  }
  F('italic', 7.5, [150, 165, 200]);
  doc.text(discLines, M, 177, { lineHeightFactor: 1.25 });
  y = HERO_H + 16;

  // ═════════════════════════ Headline strip (4 tiles) ═════════════════════════
  {
    const comp = d.comp || {};
    const total = d.total ?? 0;
    // Devices not evaluated (unknown) are excluded from the compliance score
    const evaluated = (comp.compliant ?? 0) + (comp.noncompliant ?? 0);
    const pctComp = evaluated > 0 ? Math.round((comp.compliant ?? 0) / evaluated * 100) : null;
    const sc = d.score;
    const scTone = sc ? (sc.pct >= 70 ? 'good' : sc.pct >= 50 ? 'warn' : 'bad') : 'neu';
    const u0 = d.users || {};
    const ren = u0.licenceRenewals || [];
    const next = ren.length ? ren.slice().sort((a, b) => String(a.renewalDate).localeCompare(String(b.renewalDate)))[0] : null;
    const tiles = [
      { l: 'Devices compliant', v: pctComp === null ? 'N/A' : `${pctComp}%`, t: pctComp === null ? 'neu' : pctComp >= 95 ? 'good' : 'warn', s: `${comp.compliant ?? 0} of ${evaluated} evaluated devices` },
      { l: 'Secure Score', v: sc ? `${sc.pct}%` : 'N/A', t: scTone, s: sc ? `${sc.cur} of ${sc.max} points` : 'Not available' },
      { l: 'Risky users', v: String(d.risky ?? 0), t: (d.risky ?? 0) > 0 ? 'bad' : 'good', s: (d.risky ?? 0) > 0 ? 'Accounts flagged at risk' : 'No accounts at risk' },
      u0.licenceRenewalsError
        ? { l: 'Renewals due (90 days)', v: 'N/A', t: 'neu', s: 'Data unavailable' }
        : { l: 'Renewals due (90 days)', v: String(ren.length), t: 'info', s: next ? `Next: ${fmtShort(next.renewalDate)}` : 'None due' },
    ];
    const gap = 10, w = (CW - 3 * gap) / 4, h = 68;
    tiles.forEach((t, i) => {
      const x = M + i * (w + gap);
      box(x, y, w, h, 10);
      let size = 6.8;
      while (size > 5.4 && W(t.l.toUpperCase(), 'bold', size) + 0.5 * t.l.length > w - 24) size -= 0.2;
      spaced(t.l.toUpperCase(), x + 12, y + 17, size, C.muted, 0.5);
      F('bold', 25, KPI_VAL[t.t] === C.text ? C.text : KPI_VAL[t.t]); doc.text(t.v, x + 12, y + 45);
      F('normal', 8, C.muted); doc.text(clean(t.s), x + 12, y + 58, { maxWidth: w - 20 });
    });
    y += h + 6;
  }

  // ══════════════════════════ Executive Summary ═══════════════════════════════
  if (manual.overview || manual.highlights || manual.concerns || manual.projects) {
    sectionHeader('Executive Summary', C.brand, null, 60);
    if (manual.overview) {
      F('normal', 10, C.text);
      const lines = doc.splitTextToSize(clean(manual.overview), CW);
      ensure(lines.length * 13 + 10);
      doc.text(lines, M, y + 4, { lineHeightFactor: 1.3 }); y += lines.length * 13 + 14;
    }
    if (manual.highlights) callout('Highlights:', manual.highlights, 'good');
    if (manual.concerns) callout('Concerns:', manual.concerns, 'warn');
    if (manual.projects) callout('Projects:', manual.projects, 'info');
  }

  // ══════════════════════ Month-on-Month Changes (iData.commentary) ═══════════
  if (d.commentary && d.commentary.length) {
    sectionHeader('Month-on-Month Changes', C.orange, null, 50);
    d.commentary.forEach(line => {
      const lines = richLines([{ t: clean(line), b: false }], CW - 46, 9);
      const h = lines.length * 12.5 + 14;
      ensure(h + 8);
      box(M, y, CW, h, 8, C.surf2);
      fill(C.info); doc.circle(M + 15, y + h / 2, 2.8, 'F');
      drawRich(lines, M + 27, y + 17.5, 12.5, 9, C.text);
      y += h + 8;
    });
  }

  // ══════════════════════════ Device & asset management ═══════════════════════
  {
    const comp = d.comp || {}, enc = d.encryption || {}, total = d.total ?? 0;
    sectionHeader('Device & asset management', C.brand, null, 70);
    kpiGrid([
      k(total, 'Total devices', 'info'),
      k(comp.compliant ?? 0, 'Compliant', 'good'),
      k(comp.noncompliant ?? 0, 'Non-compliant', (comp.noncompliant ?? 0) > 0 ? 'bad' : 'good'),
      k((d.lowDisk || []).length, 'Low disk (<15%)', (d.lowDisk || []).length ? 'warn' : 'good'),
    ]);

    // Compliance donut | Encryption & antivirus
    const colW = (CW - 14) / 2, hCard = 158;
    ensure(hCard + 14);
    box(M, y, colW, hCard); box(M + colW + 14, y, colW, hCard);
    spaced('COMPLIANCE', M + 14, y + 20, 7.5, C.muted, 0.7);
    const evaluated = (comp.compliant ?? 0) + (comp.noncompliant ?? 0);
    const pctComp = evaluated > 0 ? Math.round((comp.compliant ?? 0) / evaluated * 100) : 0;
    const R = 38;
    donut(M + 14 + R + 2, y + 22 + (hCard - 22) / 2, R, 11, pctComp, pctComp >= 95 ? C.good : C.warn, `${pctComp}%`, 'compliant');
    {
      const items = [
        { c: C.good, t: `${comp.compliant ?? 0} compliant` },
        { c: C.bad, t: `${comp.noncompliant ?? 0} non-compliant` },
        { c: C.neu, t: `${comp.unknown ?? 0} not evaluated (not counted)` },
        { c: C.warn, t: `${d.staleCount ?? 0} not checked in for 90+ days` },
      ];
      const lx = M + 14 + 2 * R + 20;
      const hh = items.length * 15;
      const ly = y + 22 + (hCard - 22) / 2 - hh / 2 + 10;
      items.forEach((it, i) => {
        fill(it.c); doc.roundedRect(lx, ly + i * 15 - 6, 7, 7, 2, 2, 'F');
        F('normal', 8, C.muted); doc.text(clean(it.t), lx + 12, ly + i * 15, { maxWidth: colW - (lx - M) - 20 });
      });
    }
    {
      const rx = M + colW + 14;
      spaced('ENCRYPTION & ANTIVIRUS', rx + 14, y + 20, 7.5, C.muted, 0.7);
      const mw = (colW - 28 - 8) / 2, mh = 56;
      const av = d.av;
      const mini = [
        k(enc.encrypted ?? 0, 'Encrypted', 'good'),
        k(enc.notEncrypted ?? 0, 'Not encrypted', (enc.notEncrypted ?? 0) > 0 ? 'bad' : 'good'),
      ];
      if (av) {
        const avBad = (av.outOfDate ?? 0) + (av.notActive ?? 0);
        mini.push(k(av.active ?? 0, 'AV active', 'good'));
        mini.push(k(avBad, 'AV inactive / out of date', avBad ? 'bad' : 'good'));
      }
      F('normal', 8, C.muted);
      mini.forEach((it, i) => {
        it.lines = doc.splitTextToSize(clean(it.label), mw - 20);
        kpiCard(rx + 14 + (i % 2) * (mw + 8), y + 34 + Math.floor(i / 2) * (mh + 8), mw, mh, it, 17);
      });
    }
    y += hCard + 14;

    // OS stacked bar
    {
      const parts = [
        ...(d.win25h2 > 0 ? [{ v: d.win25h2, c: [13, 124, 196], t: 'Windows 11 25H2' }] : []),
        ...(d.win24h2 > 0 ? [{ v: d.win24h2, c: [74, 160, 220], t: 'Windows 11 24H2' }] : []),
        ...(d.win11 > 0 ? [{ v: d.win11, c: [140, 196, 236], t: 'Windows 11' }] : []),
        ...(d.win10 > 0 ? [{ v: d.win10, c: [224, 100, 0], t: 'Windows 10' }] : []),
        ...(d.macOS > 0 ? [{ v: d.macOS, c: [140, 12, 110], t: 'macOS' }] : []),
        ...(d.linux > 0 ? [{ v: d.linux, c: [71, 85, 105], t: 'Linux' }] : []),
        ...(d.iosCount > 0 ? [{ v: d.iosCount, c: [245, 160, 26], t: 'iOS / iPad' }] : []),
        ...(d.androidCount > 0 ? [{ v: d.androidCount, c: [15, 157, 107], t: 'Android' }] : []),
      ];
      if (parts.length) {
        const legend = parts.map(p => ({ c: p.c, t: `${p.t} · ${p.v}` }));
        const lh = legendHeight(legend, CW - 28, 8.5);
        const h = 22 + 14 + 8 + lh + 14;
        ensure(h + 12);
        box(M, y, CW, h);
        spaced('OPERATING SYSTEMS', M + 14, y + 20, 7.5, C.muted, 0.7);
        const denom = Math.max(total, parts.reduce((s, p) => s + p.v, 0));
        stackBar(M + 14, y + 30, CW - 28, 12, parts, denom);
        drawLegend(legend, M + 14, y + 30 + 12 + 18, CW - 28, 8.5);
        y += h + 12;
      }
    }

    if ((d.comp?.noncompliant ?? 0) > 0) {
      const n = d.comp.noncompliant;
      callout(`${n} device${n > 1 ? 's are' : ' is'} non-compliant with your organisation's policies.`, '', 'bad');
    }
    if ((d.notCompliantList || []).length) {
      dataTable([{ h: 'Device', w: 26, k: 'b' }, { h: 'Primary user', w: 32 }, { h: 'OS', w: 20 }, { h: 'Last seen', w: 14 }],
        d.notCompliantList.map(x => [x.name || '', x.user || 'Unknown', x.os || '', x.lastSync ? fmtShort(x.lastSync) : 'Never']), 'Non-compliant devices');
    }
    if ((d.encryption?.notEncrypted ?? 0) > 0) {
      const n = d.encryption.notEncrypted;
      callout(`${n} device${n > 1 ? 's are' : ' is'} not encrypted.`, 'This is a significant data protection risk.', 'bad');
    }
    if ((d.notEncryptedList || []).length) {
      dataTable([{ h: 'Device', w: 28, k: 'b' }, { h: 'Primary user', w: 34 }, { h: 'OS', w: 28 }],
        d.notEncryptedList.map(x => [x.name || '', x.user || 'Unknown', x.os || '']), 'Devices not encrypted');
    }
    if ((d.lowDisk || []).length) {
      const n = d.lowDisk.length;
      callout(`${n} device${n > 1 ? 's have' : ' has'} low disk space (less than 15% free).`, '', 'warn');
      dataTable([{ h: 'Device', w: 28, k: 'b' }, { h: 'Primary user', w: 34 }, { h: 'Free', w: 12, a: 'r' }, { h: 'Free %', w: 12, k: 'pill', a: 'r' }],
        d.lowDisk.map(x => [x.name || '', x.user || '', `${x.gb} GB`, [`${x.pct}%`, x.pct < 5 ? 'bad' : 'warn']]), 'Low disk space');
    }
  }

  // ══════════════════════════ Security posture ═════════════════════════════════
  {
    sectionHeader('Security posture', C.purple, null, 160);
    const sc = d.score, ca = d.conditionalAccess;
    const idRows = [];
    if (ca) idRows.push({ label: 'Conditional Access policies', pills: [[`${ca.enabled} enforced`, 'good'], [`${ca.reportOnly} report-only`, ca.reportOnly > 0 ? 'warn' : 'neu']] });
    idRows.push({ label: 'Risky users', pills: [[String(d.risky ?? 0), (d.risky ?? 0) > 0 ? 'bad' : 'good']] });
    idRows.push({ label: 'Compliance policies', pills: [[String(d.compliancePolicies?.total ?? 0), 'neu']] });
    idRows.push({ label: 'App protection policies', pills: [[String(d.appProtection?.total ?? 0), (d.appProtection?.total ?? 0) === 0 ? 'warn' : 'neu']] });

    const gap = 14, leftW = (CW - gap) * 5 / 12, rightW = (CW - gap) * 7 / 12;
    const hRow = Math.max(rowsCardH(idRows.length), 140);
    ensure(hRow + 14);
    box(M, y, leftW, hRow);
    spaced('MICROSOFT SECURE SCORE', M + 14, y + 20, 7.5, C.muted, 0.7);
    if (sc) {
      const tone = sc.pct >= 70 ? C.good : sc.pct >= 50 ? C.warn : C.bad;
      const R = 34;
      donut(M + 14 + R, y + 22 + (hRow - 22) / 2, R, 10, sc.pct, tone, `${sc.pct}%`, `${sc.cur} / ${sc.max}`);
      const tx = M + 14 + 2 * R + 14, tw = leftW - (tx - M) - 12;
      F('bold', 9.5, C.text);
      const head = sc.pct >= 70 ? 'Strong position.' : sc.pct >= 50 ? 'Room to improve.' : 'Needs attention.';
      const hl = doc.splitTextToSize(head, tw);
      const ty = y + 22 + (hRow - 22) / 2 - 12;
      doc.text(hl, tx, ty);
      F('normal', 8, C.muted);
      doc.text(doc.splitTextToSize('Above 70% is a good target.', tw), tx, ty + hl.length * 12 + 2);
    } else {
      F('normal', 9, C.muted); doc.text('Secure Score unavailable', M + 14, y + 60);
    }
    rowsCard(M + leftW + gap, y, rightW, hRow, 'Identity & access', idRows);
    y += hRow + 14;

    const kp = d.keyPolicies || {};
    const st = { enforced: ['Enforced', 'good'], 'report-only': ['Report-only', 'warn'], 'not found': ['Not found', 'warn'], disabled: ['Disabled', 'bad'] };
    const stp = s => st[s] || [String(s), 'neu'];
    const kpRows = [
      { label: 'Block legacy authentication', pills: [stp(kp.legacyAuthBlock?.status || 'not found')] },
      { label: 'MFA for all users', pills: [stp(kp.mfaAllUsers?.status || 'not found')] },
      { label: 'MFA for admin portals', pills: [stp(kp.adminMfa?.status || 'not found')] },
      { label: 'Geographic restriction', pills: [stp(kp.geoBlock?.status || 'not found')] },
    ];
    const am = d.authMethods || {};
    const amRows = [
      { label: 'Authenticator app', pills: [am.authAppEnabled ? ['Enabled', 'good'] : ['Off', 'neu']] },
      { label: 'FIDO2 security keys', pills: [am.fido2Enabled ? ['Enabled', 'good'] : ['Off', 'neu']] },
      { label: 'Temporary Access Pass', pills: [am.tapEnabled ? (am.tapReusable ? ['Enabled - reusable', 'warn'] : ['Enabled', 'good']) : ['Off', 'neu']] },
      { label: 'SMS', pills: [am.smsEnabled ? ['Enabled - review', 'warn'] : ['Off', 'neu']] },
    ];
    const hh = rowsCardH(4);
    ensure(hh + 12);
    const hw = (CW - gap) / 2;
    rowsCard(M, y, hw, hh, 'Key policies', kpRows);
    rowsCard(M + hw + gap, y, hw, hh, 'Authentication methods', amRows);
    y += hh + 12;
    if ((d.risky ?? 0) > 0) callout(`${d.risky} account${d.risky > 1 ? 's are' : ' is'} flagged as at-risk by Entra ID Protection.`, 'Reset passwords immediately.', 'bad');
    if ((d.appProtection?.total ?? 0) === 0) callout('No app protection (MAM) policies found.', 'If BYOD access is permitted, this is a gap.', 'warn');
  }

  // ══════════════════════════ Patch status ═════════════════════════════════════
  {
    const ps = d.patchStatus || {};
    if (ps.current !== undefined) {
      sectionHeader('Patch status', C.brand, null, 60);
      const pt = (ps.current ?? 0) + (ps.over30 ?? 0) + (ps.over90 ?? 0);
      const parts = [{ v: ps.current ?? 0, c: C.good }, { v: ps.over30 ?? 0, c: C.warn }, { v: ps.over90 ?? 0, c: C.bad }];
      const legend = [
        { c: C.good, t: `Current · ${ps.current ?? 0}` },
        { c: C.warn, t: `30+ days behind · ${ps.over30 ?? 0}` },
        { c: C.bad, t: `90+ days behind · ${ps.over90 ?? 0}` },
      ];
      const h = 14 + 12 + 14 + 14 + 8;
      ensure(h + 12);
      box(M, y, CW, h);
      stackBar(M + 14, y + 14, CW - 28, 12, parts, pt || 1);
      drawLegend(legend, M + 14, y + 14 + 12 + 17, CW - 28, 8.5);
      y += h + 12;
      if ((d.patchOver90 || []).length) {
        callout(`${d.patchOver90.length} device${d.patchOver90.length > 1 ? 's have' : ' has'} not checked in for 90+ days - 3 or more patch cycles behind.`, '', 'bad');
                dataTable([{ h: 'Device', w: 24, k: 'b' }, { h: 'Primary user', w: 30 }, { h: 'Last seen', w: 14 }, { h: 'OS', w: 32 }],
          d.patchOver90.map(x => [x.name || '', x.user || 'Unknown', x.lastSeen ? fmtShort(x.lastSeen) : 'Never', x.os || 'Unknown']), 'Devices over 90 days without a check-in');
      }
    }
  }

  // ══════════════════════════ User data & licences ═════════════════════════════
  const u = d.users || {};
  if ((u.total ?? 0) > 0) {
    sectionHeader('User data & licences', C.purple, null, 80);
    const inactive = (u.notSignedIn90Licensed ?? 0) + (u.notSignedIn90Guest ?? 0);
    const ext = u.externalSignIns;
    const uk = [
      k(u.total ?? 0, 'Licensed users', 'info'),
      k(u.sharedMailboxes ?? 0, 'Shared mailboxes', 'neu'),
      k(u.guests ?? 0, 'Guest accounts', 'neu'),
      k(inactive, 'Inactive 90+ days', inactive > 0 ? 'warn' : 'good'),
    ];
    kpiGrid(uk);

    if (ext && !ext.timedOut && ext.sharedGateway && ext.sharedGateway.events > 0) {
      const sg = ext.sharedGateway;
      callout(`${sg.events.toLocaleString('en-NZ')} sign-ins through ${sg.ips.length} shared IP address${sg.ips.length > 1 ? 'es' : ''} were left out.`, `An overseas IP address used by ${sg.minUsers} or more different people is treated as shared infrastructure (typically a company VPN or gateway) rather than individual overseas logins.`, 'info');
    }
    if (ext && !ext.timedOut && ext.retentionFrom) {
      callout('Sign-in history limit.', `Microsoft only keeps sign-in logs for 30 days (7 days on free tenants), so sign-ins before ${ext.retentionFrom} can no longer be retrieved and are not covered by this check.`, 'info');
    }
    if (ext && !ext.timedOut && ext.partial) {
      callout(`Overseas sign-ins: only ${ext.windowDays} of ${ext.periodDays} days could be checked.`, `The sign-in log could not be fully retrieved in time, so results may be incomplete.${(ext.missing || []).length ? ' Not checked: ' + ext.missing.join(', ') + '.' : ''}`, 'info');
    }
    if (ext && ext.timedOut) {
      callout('Overseas sign-in data unavailable this run.', 'The sign-in log query timed out, so unexpected overseas sign-ins are not reported. Re-pulling usually resolves it.', 'info');
    }

    if ((u.licenceSummary || []).length) {
      h3('Licence assignment', 90);
      const paid = u.licenceSummary.filter(l => !l.free), free = u.licenceSummary.filter(l => l.free);
      const NAMEW = 190, BARW = 150, CNTW = 127, G = 12, X0 = M + 14;
      const drawLic = l => (ry, h) => {
        const n = l.count ?? 0, a = l.available ?? 0;
        F('bold', 8.8, C.text);
        const lines = doc.splitTextToSize(clean(l.name), NAMEW);
        doc.text(lines, X0, ry + (h - lines.length * 11) / 2 + 8.4, { lineHeightFactor: 1.25 });
        const bx = X0 + NAMEW + G, by = ry + h / 2 - 3.5;
        fill(C.neuBg); doc.roundedRect(bx, by, BARW, 7, 3.5, 3.5, 'F');
        const rx = bx + BARW + G + CNTW;
        if (a <= 0) {
          fill(C.neu); doc.roundedRect(bx, by, BARW, 7, 3.5, 3.5, 'F');
          const pw = pillW('none purchased', 6.5);
          pill('none purchased', 'neu', rx - pw, ry + h / 2, 6.5);
          F('normal', 8.5, C.muted); const t2 = ' assigned';
          const w2 = W(t2, 'normal', 8.5); doc.text(t2, rx - pw - 6 - w2, ry + h / 2 + 3);
          F('bold', 8.5, C.text); const t1 = String(n); doc.text(t1, rx - pw - 6 - w2 - W(t1, 'bold', 8.5), ry + h / 2 + 3);
        } else {
          const p = Math.min(100, n / a * 100);
          const col = n > a ? C.bad : n === a ? C.info : p >= 90 ? C.warn : l.free ? C.info : C.good;
          const fw = Math.max(p / 100 * BARW, 3);
          fill(col); doc.roundedRect(bx, by, fw, 7, 3.5, 3.5, 'F');
          const t2 = ` / ${num(a)}`, t1 = String(n);
          const w2 = W(t2, 'normal', 8.5);
          F('normal', 8.5, C.muted); doc.text(t2, rx - w2, ry + h / 2 + 3);
          F('bold', 8.5, C.text); doc.text(t1, rx - w2 - W(t1, 'bold', 8.5), ry + h / 2 + 3);
        }
      };
      ensure(70);
      const fl = cardFlow(null, 0);
      [['Paid licences', paid], ['Free licences', free]].filter(g => g[1].length).forEach(([label, items], gi) => {
        fl.row(gi ? 28 : 22, (ry, h) => spaced(label.toUpperCase(), X0, ry + h - 7, 7.5, C.muted, 0.7), false);
        items.forEach((l, li) => {
          F('bold', 8.8);
          const nl = doc.splitTextToSize(clean(l.name), NAMEW).length;
          fl.row(Math.max(25, nl * 11 + 12), drawLic(l), li > 0);
        });
      });
      fl.close();
    }

    h3('Upcoming licence renewals · next 90 days', 80);
    if (u.licenceRenewalsError) {
      callout('Licence renewal data unavailable - check that Organization.Read.All permission is granted.', '', 'warn');
    } else if ((u.licenceRenewals || []).length) {
      const rn = u.licenceRenewals.slice().sort((a, b) => String(a.renewalDate).localeCompare(String(b.renewalDate)));
      calloutSplit(`${rn.length} licence ${rn.length === 1 ? 'subscription is' : 'subscriptions are'} due for renewal within 90 days.`,
        'Monthly-billed subscriptions renew automatically each cycle; annual/multi-year ones are worth a closer look.', 'info');
      dataTable([{ h: 'Licence', w: 34, k: 'b' }, { h: 'Seats', w: 9, a: 'r' }, { h: 'Billing', w: 22, k: 'pill' }, { h: 'Renewal date', w: 16 }, { h: 'Days away', w: 13, k: 'pill', a: 'r' }],
        rn.map(r => [r.name || '', num(r.seats ?? 0),
          [String(r.cycle || '').replace(/Annual\/Multi-year/i, 'Annual / multi-year') || 'Unknown', /monthly/i.test(r.cycle || '') ? 'neu' : 'info'],
          fmtShort(r.renewalDate), [`${r.daysAway} days`, (r.daysAway ?? 99) <= 14 ? 'warn' : 'neu']]));
    } else {
      callout('No licence renewals due in the next 90 days.', '', 'good');
    }

    const dayText = x => (x.daysSince ? `${x.daysSince} days` : 'Never signed in');
    if ((u.notSignedIn90Licensed ?? 0) > 0) {
      const n = u.notSignedIn90Licensed;
      callout(`${n} M365 licensed user${n > 1 ? 's have' : ' has'} not signed in for 90+ days.`, 'Service accounts and unlicensed accounts are excluded.', 'warn');
    }
    if ((u.notSignedIn90LicensedList || []).length) {
      dataTable([{ h: 'User', w: 28, k: 'b' }, { h: 'Account', w: 34 }, { h: 'Last sign-in', w: 16 }, { h: 'Inactive', w: 16 }],
        u.notSignedIn90LicensedList.map(x => [x.name || '', x.upn || '', x.lastSignIn ? fmtShort(x.lastSignIn) : 'Never', dayText(x)]), 'Licensed users inactive for 90+ days');
    }
    if ((u.notSignedIn90Guest ?? 0) > 0) {
      const n = u.notSignedIn90Guest;
      callout(`${n} guest account${n > 1 ? 's have' : ' has'} not signed in for 90+ days.`, 'Review for stale guest access and remove if no longer needed.', 'warn');
    }
    if ((u.notSignedIn90GuestList || []).length) {
      dataTable([{ h: 'Guest', w: 28, k: 'b' }, { h: 'Account', w: 34 }, { h: 'Last sign-in', w: 16 }, { h: 'Inactive', w: 16 }],
        u.notSignedIn90GuestList.map(x => [x.name || '', String(x.upn || '').split('#EXT#')[0], x.lastSignIn ? fmtShort(x.lastSignIn) : 'Never', dayText(x)]), 'Guest accounts inactive for 90+ days');
    }
    if ((u.adminRoles || []).length) {
      dataTable([{ h: 'Name', w: 22, k: 'b' }, { h: 'Account', w: 30 }, { h: 'Roles', w: 48, k: 'pills' }],
        u.adminRoles.map(a => [a.name || 'Unknown', a.upn || '-', (a.roles || [a.role]).filter(Boolean).map(r => [r, /Global Administrator/.test(r) ? 'warn' : 'neu'])]), 'Admin role holders');
    }
    const extUsers = (ext && ext.byUser) || [];
    if (extUsers.length) {
      const extTotal = ext.total ?? 0;
      h3(`Unexpected overseas sign-ins · ${ext.partial ? `${ext.windowDays} of ${ext.periodDays} days checked` : `report period, ${ext.periodDays ?? 30} days`}`, 90);
      callout(`${extTotal} successful login${extTotal > 1 ? 's' : ''} from unexpected locations by ${ext.uniqueUsers} user${ext.uniqueUsers > 1 ? 's' : ''}.`,
        'Australia, New Zealand and Malaysia are excluded as expected locations.', 'warn');
      dataTable([{ h: 'Name', w: 22, k: 'b' }, { h: 'Email', w: 32 }, { h: 'Country / Territory', w: 32 }, { h: 'Logins', w: 10, a: 'r' }],
        extUsers.map(e => [e.name || (e.upn || '').split('@')[0], e.upn || '', (e.countries || []).join(', '), String(e.eventCount ?? 0)]));
    }
  }

  // ══════════════════════════ SharePoint / MS Teams ═════════════════════════════
  {
    const sp = d.sharepoint || {};
    if ((sp.siteCount ?? 0) > 0 || sp.error) {
      const intro = 'Overview of SharePoint Online sites and storage usage for this tenant. Site counts are sourced from the Microsoft 365 usage reports and exclude personal OneDrive sites.';
      if (sp.error) {
        sectionHeader('SharePoint / MS Teams', C.teal, intro, 60);
        callout('SharePoint data unavailable - check that Reports.Read.All permission is granted.', '', 'warn');
      } else {
        const usedPct = (sp.allocatedGB && sp.totalUsedGB) ? Math.round((sp.totalUsedGB / sp.allocatedGB) * 100) : null;
        const spKpis = [
          k(sp.siteCount ?? 0, 'Total sites', 'info'),
          k(sp.groupCount ?? 0, 'M365 Group sites', 'info'),
          k(sp.commCount ?? 0, 'Communication sites', 'neu'),
          k(sp.channelCount ?? 0, 'Teams channels', 'neu'),
          ...((sp.classicCount ?? 0) > 0 ? [k(sp.classicCount, 'Classic sites', 'neu')] : []),
          ...((sp.otherCount ?? 0) > 0 ? [k(sp.otherCount, 'Other (system sites)', 'neu')] : []),
          k(`${num(sp.totalUsedGB ?? 0)} GB`, 'Storage used', usedPct !== null && usedPct > 80 ? 'bad' : 'neu'),
          k(sp.inactiveSiteCount ?? 0, 'Inactive Sites and Channels 180 days +', (sp.inactiveSiteCount ?? 0) > 0 ? 'warn' : 'good'),
          k(sp.securityGroupCount ?? 'N/A', 'Security groups', 'info'),
        ];
        // keep header + intro + KPI grid + warning callout together on one page
        F('normal', 9, C.muted);
        const il = doc.splitTextToSize(intro, CW * 0.86).length;
        const blockH = 22 + 36 + il * 12 + 8 + gridHeight(spKpis, 3) + ((sp.inactiveSiteCount ?? 0) > 0 ? 52 : 0);
        ensure(blockH);
        sectionHeader('SharePoint / MS Teams', C.teal, intro, 60);
        kpiGrid(spKpis, 3);
        if ((sp.inactiveSiteCount ?? 0) > 0) {
          const n = sp.inactiveSiteCount;
          callout(`${n} site${n > 1 ? 's have' : ' has'} had no recorded file activity in 180+ days.`, 'Review for archiving or deletion. Note: archived sites may appear in this count.', 'warn');
        }
        if (usedPct !== null) {
          callout(`Storage: ${sp.totalUsedGB} GB used of ${sp.allocatedGB} GB allocated (${usedPct}% used).`, '', usedPct > 80 ? 'bad' : usedPct > 60 ? 'warn' : 'good');
        }
      }
    }
  }

  // ══════════════════════════ Ticketing & Support ═══════════════════════════════
  const hasTickets = manual.ticketsOpened || manual.ticketsClosed || manual.ticketsPending || manual.avgResponse || manual.avgResolution || manual.p1;
  if (hasTickets) {
    sectionHeader('Ticketing & Support', C.teal, null, 130);
    const tKpis = [];
    if (manual.ticketsOpened) tKpis.push(k(manual.ticketsOpened, 'Opened', 'neu'));
    if (manual.ticketsClosed) tKpis.push(k(manual.ticketsClosed, 'Closed', 'good'));
    if (manual.ticketsPending) tKpis.push(k(manual.ticketsPending, 'Pending', parseInt(manual.ticketsPending, 10) > 5 ? 'warn' : 'neu'));
    if (manual.avgResponse) tKpis.push(k(manual.avgResponse, 'Avg response time', 'neu'));
    if (manual.avgResolution) tKpis.push(k(manual.avgResolution, 'Avg resolution time', 'neu'));
    if (tKpis.length) kpiGrid(tKpis, Math.min(tKpis.length, 5), 18);
    callout(manual.p1 ? 'P1 incident:' : 'No P1 incidents recorded this period.', manual.p1 || '', manual.p1 ? 'bad' : 'good');
  }

  // ══════════════════════════ Recommendations & Opportunities ═══════════════════
  if (manual.personal || manual.improvements || manual.cost || manual.roadmap) {
    sectionHeader('Recommendations & Opportunities', C.orange, null, 70);
    const card = (title, colour, text) => {
      F('normal', 9.5);
      const lines = doc.splitTextToSize(clean(text), CW - 32);
      const h = lines.length * 12.5 + 34;
      ensure(h + 10);
      box(M, y, CW, h, 8, C.surf2);
      spaced(title.toUpperCase(), M + 16, y + 18, 7.5, colour, 0.7);
      F('normal', 9.5, C.text); doc.text(lines, M + 16, y + 33, { lineHeightFactor: 1.3 });
      y += h + 10;
    };
    if (manual.personal) callout('', manual.personal, 'info');
    if (manual.improvements) card('Suggested improvements', C.brand, manual.improvements);
    if (manual.cost) card('Cost-saving opportunities', C.teal, manual.cost);
    if (manual.roadmap) card('Tech roadmap', C.purple, manual.roadmap);
  }

  // ══════════════════════════ Security risk register ═════════════════════════════
  if ((d.risks || []).length) {
    sectionHeader('Security risk register', C.red, null, 100);
    const sev = { high: ['High', 'bad'], medium: ['Medium', 'warn'], low: ['Low', 'info'] };
    dataTable([{ h: 'Risk', w: 34, k: 'b' }, { h: 'Area', w: 17 }, { h: 'Rating', w: 11, k: 'pill' }, { h: 'Recommended action', w: 38 }],
      d.risks.map(r => [r.finding || '', r.area || '', sev[r.severity] || [String(r.severity || ''), 'neu'], r.action || '']));
  }

  // ── Footers ────────────────────────────────────────────────────────────────
  const pages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    stroke(C.line); doc.setLineWidth(0.5); doc.line(M, PH - 30, PW - M, PH - 30);
    F('normal', 8, C.muted);
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
