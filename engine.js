// Turns one report into its finished output: the text of a message, a PNG, or a whole PDF. Runs inside an engine
// worker, or on the page itself in browsers whose workers cannot draw.
import { renderRange, rangeText } from './render.js';
import { rangePages, stackedPages } from './print.js';
import { jpegPdf } from './pdf.js';

const PDF_DPI = 300, QUALITY = 0.92;

const encode = (cv, type) => cv.convertToBlob ? cv.convertToBlob({ type, quality: QUALITY })
  : new Promise((ok, no) => cv.toBlob(b => b ? ok(b) : no(new Error('Could not encode the picture (the report may be too large).')), type, QUALITY));
const free = cv => { cv.width = cv.height = 0; }; // hand the pixel memory back now, not at the next garbage collection

// task: { as: 'text' | 'image' | 'pdf', refs: [{ sheet, range }], page } (page: the PDF layout from reports.js)
export async function produce(book, task) {
  const [first] = task.refs;
  if (task.as === 'text') return { text: task.refs.map(d => rangeText(book, d.sheet, d.range)).join('\n\n') };

  if (task.as === 'image') {
    const cv = renderRange(book, first.sheet, first.range, { drawings: book.drawings(first.sheet) });
    try { return { blob: await encode(cv, 'image/png'), w: cv.width, h: cv.height }; } finally { free(cv); }
  }

  const pages = task.refs.length > 1 ? stackedPages(book, task.refs, task.page) : rangePages(book, first.sheet, first.range, task.page);
  const out = [];
  for (const p of pages) {
    const images = [];
    for (const part of p.parts) {
      const sheet = part.sheet ?? first.sheet;
      const cv = renderRange(book, sheet, part.rng, { drawings: book.drawings(sheet), print: part.print, width: part.w * PDF_DPI / 72 });
      try {
        images.push({ jpeg: new Uint8Array(await (await encode(cv, 'image/jpeg')).arrayBuffer()), w: cv.width, h: cv.height, x: part.x, y: part.y, dw: part.w, dh: part.h });
      } finally { free(cv); }
    }
    out.push({ pw: p.pw, ph: p.ph, images });
  }
  return { blob: jpegPdf(out), pages: pages.length };
}
