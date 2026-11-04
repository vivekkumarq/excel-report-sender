// Opens a workbook lazily: nothing is read until it is asked for, so a 37 MB data sheet nobody reports on costs
// nothing. The zip directory comes from the end of the file; each part is then read straight from the File
// (Blob.slice), inflated and parsed. Formula results are the values Excel saved, so nothing is recalculated.
// The same code runs on the page, in the engine workers (each opens the File on its own) and in Node.
import { inflateSync } from './vendor/fflate.js';
import { parseWorkbook, styleParts, styleTables, parseSST, parseTheme, parseRels, relsPath, parseDrawing, parseSheet } from './xlsx.js';
import { parseChart } from './chart.js';

const dec = new TextDecoder();
const read = async (blob, from, to) => new Uint8Array(await blob.slice(from, to).arrayBuffer());

export async function openZip(blob) {
  const tailAt = Math.max(0, blob.size - 65557), tail = await read(blob, tailAt, blob.size); // end record + longest comment
  const dv = new DataView(tail.buffer);
  let e = tail.length - 22;
  while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
  if (e < 0) throw new Error('This is not an Excel workbook (.xlsx).');
  const count = dv.getUint16(e + 10, true), size = dv.getUint32(e + 12, true), at = dv.getUint32(e + 16, true);
  const cd = at >= tailAt ? tail.subarray(at - tailAt, at - tailAt + size) : await read(blob, at, at + size);
  const v = new DataView(cd.buffer, cd.byteOffset, cd.byteLength), entries = new Map();
  for (let i = 0, p = 0; i < count; i++) {
    if (p + 46 > cd.length || v.getUint32(p, true) !== 0x02014b50) throw new Error('The workbook file is damaged.');
    const nl = v.getUint16(p + 28, true);
    entries.set(dec.decode(cd.subarray(p + 46, p + 46 + nl)).replace(/^\//, ''), {
      method: v.getUint16(p + 10, true), comp: v.getUint32(p + 20, true), size: v.getUint32(p + 24, true), at: v.getUint32(p + 42, true),
    });
    p += 46 + nl + v.getUint16(p + 30, true) + v.getUint16(p + 32, true);
  }
  const bytes = async name => {
    const z = entries.get(name);
    if (!z) return null;
    const h = new DataView((await read(blob, z.at, z.at + 30)).buffer); // the local header's lengths can differ from the directory's
    const start = z.at + 30 + h.getUint16(26, true) + h.getUint16(28, true), raw = await read(blob, start, start + z.comp);
    return z.method ? inflateSync(raw, { out: new Uint8Array(z.size) }) : raw;
  };
  return { has: n => entries.has(n), list: () => [...entries.keys()], bytes, text: async n => { const b = await bytes(n); return b ? dec.decode(b) : ''; } };
}

// Sheets, named ranges and the 1904 flag: all the page itself needs to plan the work.
export async function readWorkbook(zip) {
  const [xml, rels] = await Promise.all([zip.text('xl/workbook.xml'), zip.text('xl/_rels/workbook.xml.rels')]);
  if (!xml) throw new Error('This is not an Excel workbook (.xlsx).');
  return parseWorkbook(xml, rels);
}

// Named ranges as listed in Excel's Name Box (case-insensitive; a workbook-level name wins over a sheet-level one).
export function findNames(wb, wanted) {
  const sheets = new Set(wb.sheets.map(s => s.name)), refs = new Map(), missing = [];
  for (const n of wanted) {
    const all = wb.names.filter(d => d.name.toLowerCase() === n.toLowerCase() && d.range && sheets.has(d.sheet));
    const d = all.find(x => x.local == null) || all[0];
    d ? refs.set(n, { sheet: d.sheet, range: d.range }) : missing.push(n);
  }
  return { refs, missing };
}

// What every engine needs before it can draw anything: sheets and names, shared strings, styles and theme.
// It is plain data, so the page reads it once and gives each worker a copy instead of each reading it again.
export async function readIndex(zip) {
  const themePath = zip.list().find(k => /^xl\/theme\/theme\d*\.xml$/.test(k));
  const [wb, sst, styles, theme] = await Promise.all([readWorkbook(zip), zip.text('xl/sharedStrings.xml'), zip.text('xl/styles.xml'), themePath ? zip.text(themePath) : '']);
  return { wb, sst: parseSST(sst), styles: styleParts(styles), theme: parseTheme(theme) };
}

export async function openBook(blob, index) {
  const zip = await openZip(blob);
  const { wb, sst, styles, theme } = index ?? await readIndex(zip);
  const paths = new Map(wb.sheets.filter(s => zip.has(s.path)).map(s => [s.name, s.path]));

  const once = (cache, name, make) => cache.get(name) ?? cache.set(name, make()).get(name);
  const sheets = new Map(), parsing = new Map(), drawings = new Map(), drawingJobs = new Map();
  const loadSheet = name => once(parsing, name, async () => sheets.set(name, parseSheet(await zip.text(paths.get(name)), sst)));
  const loadDrawings = name => once(drawingJobs, name, async () => {
    const sp = paths.get(name), dp = Object.values(parseRels(await zip.text(relsPath(sp)), sp)).find(p => /drawings\/drawing\d*\.xml$/.test(p));
    const anchors = dp && zip.has(dp) ? parseDrawing(await zip.text(dp), parseRels(await zip.text(relsPath(dp)), dp)) : [];
    await Promise.all(anchors.map(async a => {
      if (a.chart && zip.has(a.chart)) a.chartModel = parseChart(await zip.text(a.chart));
      if (a.image && zip.has(a.image) && typeof createImageBitmap === 'function') a.bitmap = await createImageBitmap(new Blob([await zip.bytes(a.image)])).catch(() => null);
    }));
    drawings.set(name, anchors);
  });

  return {
    styles: styleTables(styles), theme, date1904: wb.date1904,
    sheet: n => sheets.get(n),
    drawings: n => drawings.get(n) || [],
    // Gets sheets ready to draw: parsed, with their charts and pictures, plus every sheet their conditional
    // formats read from.
    async load(names) {
      const own = [...new Set(names)].filter(n => paths.has(n));
      await Promise.all(own.flatMap(n => [loadSheet(n), loadDrawings(n)]));
      await Promise.all(own.flatMap(n => [...(sheets.get(n).refSheets || [])]).filter(n => paths.has(n)).map(loadSheet));
    },
  };
}
