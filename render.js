// Draws an Excel range onto a canvas the way Excel's "Copy as Picture" shows it: column widths, row heights,
// fills, borders, fonts, number formats, merged cells, conditional formats, icon sets, charts and pictures.
// Hidden rows and columns are left out, as in Excel.
import { KEY } from './xlsx.js';
import { cfOverrides } from './cf.js';
import { drawChart } from './chart.js';
import { cellColor, cssFont, PX, EMU } from './paint.js';
import SSF from './vendor/ssf.js';

// Excel shows built-in date formats, and the "/" date separator, in the PC's regional style.
// These match an India-locale Excel (checked against Excel's own output for this workbook).
const LOCAL_BUILTIN = { 14: 'dd-mm-yyyy', 15: 'dd-mmm-yy', 16: 'dd-mmm', 17: 'mmm-yy', 22: 'dd-mm-yyyy hh:mm' };
const DATE_SEP = '-';
const localDate = code => code.replace(/(\\.|"[^"]*")|\//g, (m, keep) => keep || DATE_SEP);
const PAD = 2;
const key = (r, c) => r * KEY + c;

let measurer;
const scratch = () => measurer ??= (typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(1, 1) : document.createElement('canvas')).getContext('2d');

// Excel measures column widths in "maximum digit widths" of the workbook's default font.
export function digitWidth(book) {
  const c = scratch();
  c.font = cssFont(book.styles.font(0) || {});
  return Math.round(Math.max(...[...'0123456789'].map(d => c.measureText(d).width)));
}
const colPx = (w, mdw) => Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw);

// Widest digit of a font as a fraction of its size (for column widths at printer resolution).
function fontRatios(font) {
  const c = scratch();
  c.font = cssFont(font, 1000);
  return { digit: Math.max(...[...'0123456789'].map(d => c.measureText(d).width)) / 1000 };
}

// When printing, Excel sizes columns at printer resolution, which makes them a little wider relative to the rows
// than on screen. PDFs follow that (checked against Excel's own printed PDFs of this workbook, within ~3%).
// ponytail: rows keep their saved heights; Excel also re-fits auto-height rows when printing, which no simple
// formula matched for every font. Revisit if a report's page breaks drift from Excel's.
const PRINT_DPI = 600;

function sizes(book, sheet, print = false) {
  const font0 = book.styles.font(0) || {}, unit = print ? 96 / PRINT_DPI : 1; // device px -> CSS px
  const mdw = print ? (book.mdwPrint ??= Math.round(fontRatios(font0).digit * Math.round((font0.sz || 11) * PRINT_DPI / 72))) : (book.mdw ??= digitWidth(book));
  const defW = (sheet.defaultColWidth != null ? colPx(sheet.defaultColWidth, mdw) : Math.ceil((sheet.baseColWidth * mdw + 5) / 8) * 8) * unit;
  const colDef = c => sheet.cols.find(k => c >= k.min && c <= k.max);
  const round = print ? v => v : Math.round; // the screen snaps rows to whole pixels; paper does not
  return {
    colDef,
    colW: c => { const k = colDef(c); return !k ? defW : k.hidden ? 0 : k.width != null ? colPx(k.width, mdw) * unit : defW; },
    rowH: r => { const x = sheet.rows.get(r); return x?.hidden ? 0 : x?.ht != null ? round(x.ht * PX) : sheet.zeroHeight ? 0 : round(sheet.defaultRowHeight * PX); },
  };
}

// Width and per-row heights of a range in CSS pixels (print: as laid out for printing).
export function rangeRows(book, sheetName, rng, print = false) {
  const { colW, rowH } = sizes(book, book.sheet(sheetName), print), [r1, c1, r2, c2] = rng, rows = [];
  let W = 0;
  for (let c = c1; c <= c2; c++) W += colW(c);
  for (let r = r1; r <= r2; r++) rows.push({ r, h: rowH(r) });
  return { W, rows };
}

function general(v) {
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  const a = Math.abs(v);
  if (a >= 1e11 || (a && a < 1e-9)) return v.toExponential(5).replace(/\.?0+e/, 'e').replace(/e([+-])(\d)$/, 'e$10$2').toUpperCase();
  return String(+v.toPrecision(10));
}

// The text Excel would display for a value in a given number format.
export function display(book, sheet, v, numFmt) {
  if (v == null) return '';
  const code = book.styles.numFmts[numFmt] ?? LOCAL_BUILTIN[numFmt] ?? numFmt;
  if (typeof v === 'number') {
    if (v === 0 && !sheet.showZeros) return '';
    if (code === 0 || /^general$/i.test(code)) return general(v);
    try { return SSF.format(typeof code === 'string' && SSF.is_date(code) ? localDate(code) : code, v, { date1904: book.date1904 }); } catch { return general(v); }
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (v.error) return v.error;
  const s = typeof v === 'string' ? v : v.text;
  if (typeof code === 'string' && code.includes('@') && code !== '@') { try { return SSF.format(code, s); } catch { /* keep the raw text */ } }
  return s;
}

// ---- text layout ----
const metricsCache = new Map();
function metrics(ctx, css) {
  let m = metricsCache.get(css);
  if (!m) {
    ctx.font = css;
    const t = ctx.measureText('Hg'), size = parseFloat(/([\d.]+)px/.exec(css)[1]);
    m = { asc: t.fontBoundingBoxAscent ?? size * 0.8, desc: t.fontBoundingBoxDescent ?? size * 0.25 };
    metricsCache.set(css, m);
  }
  return m;
}

function layout(ctx, runs, maxW, wrap) {
  const lines = [[]];
  let lineW = 0;
  for (const run of runs) {
    const base = (run.font.sz || 11) * PX, va = run.font.va;
    const css = va ? cssFont(run.font, base * 0.65) : cssFont(run.font), shift = va === 'superscript' ? -base * 0.38 : va === 'subscript' ? base * 0.12 : 0;
    ctx.font = css;
    for (const part of (wrap ? run.text : run.text.replace(/\r?\n/g, '')).split(/(\r?\n)/)) {
      if (/^\r?\n$/.test(part)) { lines.push([]); lineW = 0; continue; }
      // a word wider than the cell is broken between characters, as Excel does
      const words = wrap ? part.split(/(?<= )/).flatMap(wd => {
        if (ctx.measureText(wd.trimEnd()).width <= maxW) return [wd];
        const out = [''];
        for (const ch of wd) ctx.measureText(out.at(-1) + ch).width > maxW && out.at(-1) ? out.push(ch) : (out[out.length - 1] += ch);
        return out;
      }) : [part];
      for (const word of words) {
        if (!word) continue;
        const w = ctx.measureText(word).width;
        if (wrap && lineW > 0 && lineW + ctx.measureText(word.trimEnd()).width > maxW) { lines.push([]); lineW = 0; }
        lines.at(-1).push({ text: word, font: run.font, css, w, shift });
        lineW += w;
      }
    }
  }
  return lines.map(segs => {
    const ms = segs.length ? segs.map(s => metrics(ctx, s.css)) : [metrics(ctx, cssFont(runs[0].font))];
    const tail = segs.at(-1), trail = tail && wrap ? tail.w - (ctx.font = tail.css, ctx.measureText(tail.text.trimEnd()).width) : 0;
    return { segs, w: segs.reduce((s, x) => s + x.w, 0) - trail, asc: Math.max(...ms.map(m => m.asc)), desc: Math.max(...ms.map(m => m.desc)) };
  });
}

function drawSegs(ctx, book, line, x, base) {
  for (const s of line.segs) {
    ctx.font = s.css;
    ctx.fillStyle = cellColor(s.font.color, book, '#000000');
    ctx.fillText(s.text, x, base + (s.shift || 0));
    if (s.font.u || s.font.strike) {
      const sz = (s.font.sz || 11) * PX, y = s.font.u ? base + sz * 0.12 : base - sz * 0.28;
      ctx.fillRect(x, y, s.w, Math.max(1, sz / 15));
    }
    x += s.w;
  }
}

function drawText(ctx, book, t) {
  const { rect, clip } = t, indent = t.indent * 9;
  let runs = t.runs, avail = rect.w - 2 * PAD - indent - (t.iconW || 0);
  ctx.save();
  ctx.beginPath(); ctx.rect(clip.x, clip.y, clip.w, clip.h); ctx.clip();
  ctx.textBaseline = 'alphabetic'; ctx.textAlign = 'left';

  if (t.rot) {
    const line = layout(ctx, runs, Infinity, false)[0], ang = t.rot === 255 ? 0 : t.rot <= 90 ? -t.rot : t.rot - 90;
    ctx.translate(rect.x + rect.w / 2, rect.y + rect.h / 2);
    if (t.rot === 255) { // stacked letters, top to bottom
      const chars = [...runs.map(r => r.text).join('')], m = line ? line.asc + line.desc : 12, total = chars.length * m;
      const top = t.v === 'top' ? -rect.h / 2 + PAD : t.v === 'center' ? -total / 2 : rect.h / 2 - PAD - total;
      ctx.textAlign = 'center';
      chars.forEach((ch, i) => { ctx.font = cssFont(runs[0].font); ctx.fillStyle = cellColor(runs[0].font.color, book, '#000'); ctx.fillText(ch, 0, top + m * (i + 0.8)); });
    } else if (line) {
      ctx.rotate(ang * Math.PI / 180);
      const along = Math.abs(ang) === 90 ? rect.h : rect.w, start = t.v === 'top' ? along / 2 - PAD - line.w : t.v === 'center' ? -line.w / 2 : -along / 2 + PAD;
      drawSegs(ctx, book, line, ang === 90 ? -start - line.w : start, (line.asc - line.desc) / 2);
    }
    ctx.restore();
    return;
  }

  let lines = layout(ctx, runs, avail, t.wrap);
  const widest = Math.max(...lines.map(l => l.w));
  if (t.shrink && !t.wrap && widest > avail && widest > 0) {
    const k = avail / widest;
    runs = runs.map(r => ({ ...r, font: { ...r.font, sz: (r.font.sz || 11) * k } }));
    lines = layout(ctx, runs, avail, false);
  } else if (t.general != null && !t.wrap) {
    // Excel's General format fits a number into as many characters as the column holds in the cell's font:
    // first by dropping decimals (never integer digits), then scientific notation, else ####.
    ctx.font = cssFont(runs[0].font);
    const v = t.general, maxChars = Math.max(1, Math.floor((rect.w - 5) / ctx.measureText('0').width));
    if (runs[0].text.length > maxChars || widest > avail) {
      const ok = s => s.length <= maxChars && ctx.measureText(s).width <= avail;
      const intDigits = Math.abs(v) >= 1 ? Math.floor(Math.log10(Math.abs(v))) + 1 : 1;
      let s = null;
      for (let p = 10; p >= intDigits && !s; p--) { const x = String(+v.toPrecision(p)); if (!/e/i.test(x) && ok(x)) s = x; }
      for (let d = 5; d >= 0 && !s; d--) { const x = v.toExponential(d).replace(/e([+-])(\d)$/, 'e$10$2').toUpperCase(); if (ok(x)) s = x; }
      lines = layout(ctx, [{ ...runs[0], text: s ?? '#'.repeat(maxChars) }], avail, false);
    }
  } else if (t.num && !t.wrap && widest > rect.w) { // whole cell width: a pixel of font difference must not cause ####
    // A formatted number that does not fit shows as ####, as in Excel.
    ctx.font = cssFont(runs[0].font);
    lines = layout(ctx, [{ ...runs[0], text: '#'.repeat(Math.max(1, Math.floor(avail / ctx.measureText('#').width))) }], avail, false);
  }
  const total = lines.reduce((s, l) => s + l.asc + l.desc, 0);
  let y = t.v === 'top' || t.v === 'justify' || t.v === 'distributed' ? rect.y + 1 : t.v === 'center' ? rect.y + (rect.h - total) / 2 : rect.y + rect.h - total - 1;
  for (const l of lines) {
    const x = t.h === 'right' ? rect.x + rect.w - PAD - indent - l.w : t.h === 'center' ? rect.x + (t.iconW || 0) + (rect.w - (t.iconW || 0) - l.w) / 2 : rect.x + PAD + indent + (t.iconW || 0);
    drawSegs(ctx, book, l, x, y + l.asc);
    y += l.asc + l.desc;
  }
  ctx.restore();
}

// ---- icon sets (traffic lights, stars, arrows) ----
function drawIcon(ctx, ic, x, y, s) {
  const n = parseInt(ic.set, 10) || 3, t = n > 1 ? ic.id / (n - 1) : 1;
  const color = t < 0.25 ? '#db5434' : t < 0.75 ? '#e9bd2b' : '#64a38d'; // Excel's icon-set colours
  const cx = x + s / 2, cy = y + s / 2, r = s / 2;
  ctx.save();
  if (/Star/i.test(ic.set)) {
    const star = () => { ctx.beginPath(); for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + i * Math.PI / 5, rr = i % 2 ? r * 0.45 : r; ctx.lineTo(cx + rr * Math.cos(a), cy + rr * Math.sin(a)); } ctx.closePath(); };
    star(); ctx.fillStyle = '#ffffff'; ctx.fill();
    if (ic.id > 0) { ctx.save(); star(); ctx.clip(); ctx.fillStyle = '#f5b800'; ctx.fillRect(x, y, ic.id >= n - 1 ? s : s / 2, s); ctx.restore(); }
    star(); ctx.strokeStyle = '#c99700'; ctx.lineWidth = 1; ctx.stroke();
  } else if (/Arrow|Triangle/i.test(ic.set)) {
    ctx.translate(cx, cy); ctx.rotate(t < 0.25 ? Math.PI / 2 : t < 0.75 ? 0 : -Math.PI / 2);
    ctx.fillStyle = color; ctx.beginPath();
    ctx.moveTo(r, 0); ctx.lineTo(0, -r); ctx.lineTo(0, -r * 0.4); ctx.lineTo(-r, -r * 0.4); ctx.lineTo(-r, r * 0.4); ctx.lineTo(0, r * 0.4); ctx.lineTo(0, r);
    ctx.closePath(); ctx.fill();
  } else {
    ctx.fillStyle = color; ctx.beginPath(); ctx.arc(cx, cy, r - 0.5, 0, Math.PI * 2); ctx.fill();
    ctx.strokeStyle = 'rgba(0,0,0,.35)'; ctx.lineWidth = 1; ctx.stroke();
  }
  ctx.restore();
}

// ---- borders ----
const BW = { hair: 1, thin: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };
const DASH = { hair: [1, 1], dotted: [1, 1], dashed: [3, 1], mediumDashed: [6, 2], dashDot: [5, 2, 1, 2], mediumDashDot: [6, 2, 2, 2], dashDotDot: [5, 2, 1, 2, 1, 2], mediumDashDotDot: [6, 2, 2, 2, 2, 2], slantDashDot: [6, 2, 2, 2] };
function stroke(ctx, book, e) {
  const { b } = e, w = BW[b.style] || 1, horiz = e.y1 === e.y2;
  ctx.strokeStyle = cellColor(b.color, book, '#000000');
  ctx.setLineDash(DASH[b.style] || []);
  const ext = w / 2, seg = (o, lw) => {
    ctx.lineWidth = lw; ctx.beginPath();
    if (horiz) { ctx.moveTo(e.x1 - ext, e.y1 + o); ctx.lineTo(e.x2 + ext, e.y2 + o); } else { ctx.moveTo(e.x1 + o, e.y1 - ext); ctx.lineTo(e.x2 + o, e.y2 + ext); }
    ctx.stroke();
  };
  if (b.style === 'double') { seg(-1, 1); seg(1, 1); } else seg(0, w);
}

const DENSITY = { gray0625: 0.0625, gray125: 0.125, lightGray: 0.25, mediumGray: 0.5, darkGray: 0.75 };
function fillCss(fill, book) {
  if (!fill) return null;
  const c = cellColor(fill.color, book, '#000000');
  if (fill.pattern === 'solid') return c;
  const d = DENSITY[fill.pattern] ?? 0.5, n = parseInt(c.slice(1), 16);
  return `rgb(${[16, 8, 0].map(s => Math.round(255 - (255 - ((n >> s) & 255)) * d)).join(',')})`;
}

// ---- the range ----
export function renderRange(book, sheetName, rng, opts = {}) {
  const sheet = book.sheet(sheetName), [r1, c1, r2, c2] = rng, styles = book.styles;
  const { colW, rowH, colDef } = sizes(book, sheet, opts.print);
  const ws = [], hs = [];
  for (let c = c1; c <= c2; c++) ws.push(colW(c));
  for (let r = r1; r <= r2; r++) hs.push(rowH(r));
  const xs = [0], ys = [0];
  for (const w of ws) xs.push(xs.at(-1) + w);
  for (const h of hs) ys.push(ys.at(-1) + h);
  const W = xs.at(-1), H = ys.at(-1);
  const span = (f, a, b) => { let s = 0; for (let i = a; i <= b; i++) s += f(i); return s; };
  const xAt = c => c < c1 ? -span(colW, c, c1 - 1) : c <= c2 + 1 ? xs[c - c1] : W + span(colW, c2 + 1, c - 1);
  const yAt = r => r < r1 ? -span(rowH, r, r1 - 1) : r <= r2 + 1 ? ys[r - r1] : H + span(rowH, r2 + 1, r - 1);

  const want = opts.width ? opts.width / Math.max(1, W) : opts.scale ?? 2; // opts.width: target canvas width in pixels
  const scale = Math.min(want, Math.sqrt((opts.maxPixels ?? 16e6) / Math.max(1, W * H)));
  const cw = Math.max(1, Math.ceil(W * scale)), ch = Math.max(1, Math.ceil(H * scale));
  // OffscreenCanvas lets the pixels move to an encoder worker without a copy; a CPU-backed context
  // (willReadFrequently) avoids slow GPU read-backs when encoding.
  const canvas = typeof OffscreenCanvas !== 'undefined' ? new OffscreenCanvas(cw, ch) : Object.assign(document.createElement('canvas'), { width: cw, height: ch });
  const ctx = canvas.getContext('2d', { willReadFrequently: true });
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, cw, ch); // whole canvas: fractional sizes leave an edge otherwise
  ctx.scale(scale, scale);

  const colStyle = ws.map((_, i) => colDef(c1 + i)?.s ?? 0);
  const styleId = (r, c) => { const cell = sheet.cells.get(key(r, c)); if (cell) return cell.s; const row = sheet.rows.get(r); return row?.s ?? (c >= c1 && c <= c2 ? colStyle[c - c1] : colDef(c)?.s ?? 0); };
  const xf = id => styles.xf(id) || styles.xf(0);
  const visible = (r, c) => hs[r - r1] > 0 && ws[c - c1] > 0;
  const cellRect = (r, c) => ({ x: xs[c - c1], y: ys[r - r1], w: ws[c - c1], h: hs[r - r1] });

  // merged areas, clipped to the range; drawn once from their first visible cell
  const mergeAt = new Map();
  for (const m of sheet.merges) {
    if (m[2] < r1 || m[0] > r2 || m[3] < c1 || m[1] > c2) continue;
    const a = Math.max(m[0], r1), b = Math.max(m[1], c1), e = Math.min(m[2], r2), f = Math.min(m[3], c2);
    let ar = a, ac = b;
    while (ar < e && !hs[ar - r1]) ar++;
    while (ac < f && !ws[ac - c1]) ac++;
    const info = { r1: m[0], c1: m[1], r2: m[2], c2: m[3], ar, ac, rect: { x: xs[b - c1], y: ys[a - r1], w: xs[f - c1 + 1] - xs[b - c1], h: ys[e - r1 + 1] - ys[a - r1] } };
    for (let r = a; r <= e; r++) for (let c = b; c <= f; c++) mergeAt.set(key(r, c), info);
  }
  const cfo = cfOverrides(book, sheet, rng);

  // plan text first: spilled text and merges decide which gridlines Excel leaves out
  const jobs = [], noGrid = new Set();
  const hasValue = (r, c) => sheet.cells.get(key(r, c))?.v != null || mergeAt.has(key(r, c));
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
    if (!visible(r, c)) continue;
    const m = mergeAt.get(key(r, c));
    if (m && (r !== m.ar || c !== m.ac)) continue;
    const sr = m ? m.r1 : r, sc = m ? m.c1 : c, cell = sheet.cells.get(key(sr, sc)), ov = cfo.get(key(sr, sc)) ?? (m && cfo.get(key(r, c)));
    if ((cell?.v == null || cell.v === '') && !ov?.icon) continue;
    const st = xf(cell ? cell.s : styleId(sr, sc));
    const text = ov?.hideValue ? '' : display(book, sheet, cell?.v, st.numFmt);
    const isNum = typeof cell?.v === 'number';
    const font = { ...styles.font(st.font), ...(ov?.font || {}) };
    const runs = cell?.v?.runs && text === cell.v.text ? cell.v.runs.map(x => ({ text: x.text, font: { ...font, ...(x.font || {}), ...(ov?.font || {}) } })) : [{ text, font }];
    let h = st.h;
    if (h === 'general') h = isNum ? 'right' : typeof cell?.v === 'boolean' || cell?.v?.error ? 'center' : 'left';
    if (h === 'fill' || h === 'justify' || h === 'distributed') h = 'left';
    let rect = m ? m.rect : cellRect(r, c), clip = rect;
    const iconW = ov?.icon ? Math.min(rect.h - 2, Math.max(12, (font.sz || 11) * PX)) + 3 : 0; // icons scale with the font
    if (ov?.icon) jobs.push({ icon: ov.icon, x: rect.x + PAD, y: rect.y + (rect.h - (iconW - 3)) / 2, s: iconW - 3 });
    if (!text) continue;

    if (h === 'centerContinuous') {
      let e = c;
      while (e < c2 && !hasValue(r, e + 1) && xf(styleId(r, e + 1)).h === 'centerContinuous') e++;
      rect = clip = { x: rect.x, y: rect.y, w: xs[e - c1 + 1] - rect.x, h: rect.h };
      for (let k = c; k < e; k++) noGrid.add(key(r, k));
      h = 'center';
    } else if (!m && !st.wrap && !st.shrink && !st.rot && !isNum) {
      // text spills into empty neighbours, as in Excel
      const c0 = scratch();
      const need = runs.reduce((s, x) => (c0.font = cssFont(x.font), s + c0.measureText(x.text).width), 0) + 2 * PAD;
      if (need > rect.w) {
        let a = c, b = c;
        const grow = dir => { const n = dir > 0 ? b + 1 : a - 1; if (n < c1 || n > c2 || hasValue(r, n)) return false; dir > 0 ? (b = n) : (a = n); return true; };
        const width = () => xs[b - c1 + 1] - xs[a - c1];
        if (h === 'left') while (width() < need && grow(1));
        else if (h === 'right') while (width() < need && grow(-1));
        else { let ok = true; while (width() < need && ok) ok = grow(1) | grow(-1); }
        clip = { x: xs[a - c1], y: rect.y, w: width(), h: rect.h };
        for (let k = a; k < b; k++) noGrid.add(key(r, k));
        if (h === 'center') rect = { ...rect, x: rect.x + rect.w / 2 - need / 2, w: need };
      }
    }
    const isGeneral = isNum && (st.numFmt === 0 || /^general$/i.test(styles.numFmts[st.numFmt] || ''));
    jobs.push({ runs, rect, clip, h, v: st.v, wrap: st.wrap, shrink: st.shrink, rot: st.rot, indent: st.indent, iconW: h === 'right' ? 0 : iconW, general: isGeneral ? cell.v : null, num: isNum && !isGeneral, mdw: book.mdw });
  }


  if (sheet.showGrid) {
    ctx.strokeStyle = '#d4d4d4'; ctx.lineWidth = 1; ctx.beginPath();
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
      if (!visible(r, c)) continue;
      const k = key(r, c), m = mergeAt.get(k), x0 = xs[c - c1], x = xs[c - c1 + 1], y0 = ys[r - r1], y = ys[r - r1 + 1];
      if (!(m && c < m.c2) && !noGrid.has(k)) { ctx.moveTo(x - 0.5, y0); ctx.lineTo(x - 0.5, y); }
      if (!(m && r < m.r2)) { ctx.moveTo(x0, y - 0.5); ctx.lineTo(x, y - 0.5); }
    }
    ctx.stroke();
  }

  // fills
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
    if (!visible(r, c)) continue;
    const m = mergeAt.get(key(r, c));
    if (m && (r !== m.ar || c !== m.ac)) continue;
    const sr = m ? m.r1 : r, sc = m ? m.c1 : c, ov = cfo.get(key(sr, sc)) ?? (m && cfo.get(key(r, c)));
    const fill = ov && ov.fill !== undefined ? ov.fill : styles.fill(xf(styleId(sr, sc)).fill);
    const css = fillCss(fill, book);
    if (!css) continue;
    const rc = m ? m.rect : cellRect(r, c);
    ctx.fillStyle = css; ctx.fillRect(rc.x, rc.y, rc.w, rc.h);
  }

  // borders (thin first, so heavier lines end up on top)
  const edges = [];
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
    if (!visible(r, c)) continue;
    const k = key(r, c), m = mergeAt.get(k), b = styles.border(xf(styleId(r, c)).border) || {}, ob = cfo.get(k)?.border || {};
    const { x, y, w, h } = cellRect(r, c);
    const side = s => ob[s] || b[s];
    if (side('t') && !(m && r > m.r1)) edges.push({ b: side('t'), x1: x, y1: y, x2: x + w, y2: y });
    if (side('b') && !(m && r < m.r2)) edges.push({ b: side('b'), x1: x, y1: y + h, x2: x + w, y2: y + h });
    if (side('l') && !(m && c > m.c1)) edges.push({ b: side('l'), x1: x, y1: y, x2: x, y2: y + h });
    if (side('r') && !(m && c < m.c2)) edges.push({ b: side('r'), x1: x + w, y1: y, x2: x + w, y2: y + h });
  }
  edges.sort((p, q) => (BW[p.b.style] || 1) - (BW[q.b.style] || 1));
  for (const e of edges) stroke(ctx, book, e);
  ctx.setLineDash([]);

  for (const j of jobs) j.icon ? drawIcon(ctx, j.icon, j.x, j.y, j.s) : drawText(ctx, book, j);

  // charts and pictures anchored over the range
  for (const d of opts.drawings || []) {
    const x1 = xAt(d.from.c) + d.from.cOff / EMU, y1 = yAt(d.from.r) + d.from.rOff / EMU;
    const x2 = d.to ? xAt(d.to.c) + d.to.cOff / EMU : x1 + (d.ext?.cx || 0) / EMU, y2 = d.to ? yAt(d.to.r) + d.to.rOff / EMU : y1 + (d.ext?.cy || 0) / EMU;
    if (x2 <= 0 || y2 <= 0 || x1 >= W || y1 >= H || x2 - x1 < 2 || y2 - y1 < 2) continue;
    if (d.chartModel) drawChart(ctx, d.chartModel, x1, y1, x2 - x1, y2 - y1, book.theme);
    else if (d.bitmap) ctx.drawImage(d.bitmap, x1, y1, x2 - x1, y2 - y1);
  }
  canvas.logical = { w: W, h: H }; // size at 100% zoom, used for PDF page size
  return canvas;
}

// Plain text of a range (one line per visible row), for reports sent as a WhatsApp message.
export function rangeText(book, sheetName, rng) {
  const sheet = book.sheet(sheetName), [r1, c1, r2, c2] = rng, { colW, rowH } = sizes(book, sheet), lines = [];
  for (let r = r1; r <= r2; r++) {
    if (!rowH(r)) continue;
    const parts = [];
    for (let c = c1; c <= c2; c++) {
      const cell = colW(c) && sheet.cells.get(key(r, c));
      const t = cell ? display(book, sheet, cell.v, (book.styles.xf(cell.s) || book.styles.xf(0)).numFmt) : '';
      if (t) parts.push(t);
    }
    lines.push(parts.join(' '));
  }
  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trim();
}
