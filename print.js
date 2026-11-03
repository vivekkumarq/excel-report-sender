// Lays reports out on paper the way Excel's "Print to PDF" does for these files:
// the sheet's own margins, the range scaled to the page width (or the whole page), page breaks at row
// boundaries when a page is full, plus the sheet's manual page breaks. Pages are left/top aligned.
import { rangeRows } from './render.js';

const PAPER = { letter: [612, 792], a4: [595.28, 841.89] }; // points

export function pageSize({ paper = 'letter', orientation = 'portrait' }) {
  const [w, h] = PAPER[paper] || PAPER.letter;
  return orientation === 'landscape' ? [h, w] : [w, h];
}

const marginsOf = (cfg, sheet) => typeof cfg.margins === 'number'
  ? { l: cfg.margins, r: cfg.margins, t: cfg.margins, b: cfg.margins }
  : cfg.margins || sheet.margins;

// One range -> pages: [{ pw, ph, parts: [{ rng, x, y, w, h }] }], sizes in points.
export function rangePages(book, sheetName, rng, cfg) {
  const [pw, ph] = pageSize(cfg), m = marginsOf(cfg, book.sheet(sheetName));
  const { W, rows } = rangeRows(book, sheetName, rng, true), H = rows.reduce((s, x) => s + x.h, 0);
  const uw = pw - m.l - m.r, uh = ph - m.t - m.b;
  const k = cfg.fit === 'page' ? Math.min(uw / W, uh / H) : uw / W; // points per CSS pixel
  const breaks = new Set(book.sheet(sheetName).rowBreaks || []), spans = [];
  let start = rows[0]?.r, acc = 0;
  rows.forEach(({ r, h }, i) => {
    if (acc > 0 && acc + h * k > uh + 0.01) { spans.push([start, r - 1, acc]); start = r; acc = 0; }
    acc += h * k;
    if (breaks.has(r) || i === rows.length - 1) { spans.push([start, r, acc]); start = r + 1; acc = 0; }
  });
  return spans.filter(s => s[2] > 0).map(([a, b, h]) => ({ pw, ph, parts: [{ rng: [a, rng[1], b, rng[3]], x: m.l, y: m.t, w: W * k, h, print: true }] }));
}

// Several ranges stacked down the page at full text width, `perPage` to a page (like pictures pasted in Word).
export function stackedPages(book, refs, cfg) {
  const [pw, ph] = pageSize(cfg), m = marginsOf(cfg, {}), uw = pw - m.l - m.r, gap = cfg.gap ?? 1.4, pages = [];
  refs.forEach((d, i) => {
    const { W, rows } = rangeRows(book, d.sheet, d.range), H = rows.reduce((s, x) => s + x.h, 0);
    if (i % (cfg.perPage || 1) === 0) pages.push({ pw, ph, parts: [], y: m.t });
    const page = pages.at(-1), h = H * uw / W;
    page.parts.push({ sheet: d.sheet, rng: d.range, x: m.l, y: page.y, w: uw, h, print: false });
    page.y += h + gap;
  });
  return pages;
}
