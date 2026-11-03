// Loads only what the configured reports need from a workbook, as fast as the browser allows:
//  1. Read the zip directory and inflate just the small XML parts (workbook, strings, styles, theme, drawings).
//     Every worksheet is skipped, so a 37 MB data sheet nobody reports on costs nothing.
//  2. Send the compressed bytes of only the needed worksheets to a pool of Web Workers, which inflate and parse
//     them in parallel (largest sheet first, onto the least-busy worker). Styles and charts are parsed meanwhile.
//  3. Load any extra sheet a conditional-format rule reads from.
// Formula results are taken from the values Excel saved, so nothing is recalculated.
import { inflateSync } from './vendor/fflate.js';
import { parseWorkbook, parseStyles, parseSST, parseTheme, parseRels, relsPath, parseDrawing, parseSheet } from './xlsx.js';
import { parseChart } from './chart.js';
import { pool } from './pool.js';

const dec = new TextDecoder();

function zipIndex(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let e = bytes.length - 22;
  const stop = Math.max(0, bytes.length - 65557);
  while (e >= stop && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < stop) throw new Error('This is not an .xlsx file.');
  const out = new Map();
  for (let i = 0, n = dv.getUint16(e + 10, true), p = dv.getUint32(e + 16, true); i < n; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('The .xlsx file is damaged.');
    const nl = dv.getUint16(p + 28, true), lh = dv.getUint32(p + 42, true);
    out.set(dec.decode(bytes.subarray(p + 46, p + 46 + nl)).replace(/^\//, ''), {
      method: dv.getUint16(p + 10, true), comp: dv.getUint32(p + 20, true), size: dv.getUint32(p + 24, true),
      start: lh + 30 + dv.getUint16(lh + 26, true) + dv.getUint16(lh + 28, true),
    });
    p += 46 + nl + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
  }
  return out;
}

export async function loadBook(file, wanted) {
  const t0 = performance.now(), timing = {};
  const bytes = new Uint8Array(await file.arrayBuffer());
  timing.read = performance.now() - t0;

  const zip = zipIndex(bytes);
  const raw = name => { const e = zip.get(name); return e && bytes.subarray(e.start, e.start + e.comp); };
  const get = name => { const e = zip.get(name); return e ? (e.method ? inflateSync(raw(name), { out: new Uint8Array(e.size) }) : raw(name)) : null; };
  const text = name => { const b = get(name); return b ? dec.decode(b) : ''; };

  const wb = parseWorkbook(text('xl/workbook.xml'), text('xl/_rels/workbook.xml.rels'));
  const sst = parseSST(text('xl/sharedStrings.xml'));
  const sheetByName = new Map(wb.sheets.map(s => [s.name, s]));

  const names = new Map(), missing = [];
  for (const n of wanted) {
    const all = wb.names.filter(d => d.name.toLowerCase() === n.toLowerCase() && d.range && sheetByName.has(d.sheet));
    const d = all.find(x => x.local == null) || all[0];
    d ? names.set(n, d) : missing.push(n);
  }

  // Parse in workers when the browser has them; any worker failure falls back to parsing on this thread.
  const inline = job => parseSheet(dec.decode(job.method ? inflateSync(job.bytes, { out: new Uint8Array(job.size) }) : job.bytes), sst);
  const workers = typeof Worker === 'undefined' ? null : pool(new URL('./parse-worker.js', import.meta.url), { max: Math.min(navigator.hardwareConcurrency || 4, 8), init: { sst } });
  const parse = job => workers ? workers.run(job, { cost: job.size }).catch(() => inline(job)) : Promise.resolve().then(() => inline(job));
  const sheets = new Map();
  const parseSheets = list => Promise.all([...list]
    .filter(n => sheetByName.has(n) && !sheets.has(n) && zip.has(sheetByName.get(n).path))
    .map(n => ({ n, e: zip.get(sheetByName.get(n).path) }))
    .sort((a, b) => b.e.size - a.e.size) // longest job first
    .map(({ n, e }) => { sheets.set(n, null); return parse({ bytes: raw(sheetByName.get(n).path).slice(), method: e.method, size: e.size }).then(s => sheets.set(n, s)); }));

  const t1 = performance.now();
  const needed = new Set([...names.values()].map(d => d.sheet));
  const parsing = parseSheets(needed);

  // While the workers parse, the main thread reads styles, theme, charts and pictures.
  const styles = parseStyles(text('xl/styles.xml'));
  const themePath = [...zip.keys()].find(k => /^xl\/theme\/theme\d*\.xml$/.test(k));
  const theme = parseTheme(themePath ? text(themePath) : '');
  const drawings = new Map(), bitmaps = [];
  for (const n of needed) {
    const sp = sheetByName.get(n).path, dp = Object.values(parseRels(text(relsPath(sp)), sp)).find(p => /drawings\/drawing\d*\.xml$/.test(p));
    if (!dp || !zip.has(dp)) continue;
    const anchors = parseDrawing(text(dp), parseRels(text(relsPath(dp)), dp));
    for (const a of anchors) {
      if (a.chart && zip.has(a.chart)) a.chartModel = parseChart(text(a.chart));
      if (a.image && zip.has(a.image) && typeof createImageBitmap === 'function') bitmaps.push(createImageBitmap(new Blob([get(a.image)])).then(b => { a.bitmap = b; }, () => {}));
    }
    drawings.set(n, anchors);
  }
  await parsing;
  const extra = new Set();
  for (const n of needed) for (const s of sheets.get(n).refSheets) if (!sheets.has(s)) extra.add(s);
  await parseSheets(extra);
  await Promise.all(bitmaps);
  timing.parse = performance.now() - t1;
  timing.workers = workers?.size ?? 0;
  timing.sheets = [...sheets.keys()].length;
  timing.skipped = wb.sheets.length - timing.sheets;
  workers?.close();

  return { book: { styles, theme, date1904: wb.date1904, sheet: n => sheets.get(n) }, names, missing, drawings, timing };
}
