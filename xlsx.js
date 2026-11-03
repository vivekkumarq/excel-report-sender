// Minimal .xlsx (OOXML) reader with just what the report renderer needs.
// Plain regex scanning, no DOM, so the same code runs on the page, in Web Workers and in Node.

const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
export const unesc = s => (s.includes('&') || s.includes('_x'))
  ? s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => e[0] !== '#' ? ENT[e] ?? m
      : String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : +e.slice(1)))
    .replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16)))
  : s;

export const attrs = s => { const o = {}; for (const m of s.matchAll(/([\w:]+)="([^"]*)"/g)) o[m[1]] = m[2]; return o; };
const on = v => v === '1' || v === 'true';

// ---- cell references (1-based rows/cols) ----
export const KEY = 16384;
export const colNum = s => { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n; };
export const colName = n => { let s = ''; for (; n > 0; n = (n - 1) / 26 | 0) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
const ref1 = s => { const m = /^\$?([A-Z]*)\$?(\d*)$/.exec(s); return m && [m[2] ? +m[2] : 0, m[1] ? colNum(m[1]) : 0]; };
export function parseRange(ref) {
  const [a, b = a] = ref.split(':'), p = ref1(a), q = ref1(b);
  if (!p || !q) return null;
  return [Math.min(p[0], q[0]) || 1, Math.min(p[1], q[1]) || 1, Math.max(p[0], q[0]) || 1048576, Math.max(p[1], q[1]) || 16384];
}

// ---- package relationships ----
const norm = path => { const o = []; for (const p of path.split('/')) p === '..' ? o.pop() : p && p !== '.' && o.push(p); return o.join('/'); };
export const dirOf = p => p.replace(/[^/]+$/, '');
export const relsPath = p => p.replace(/([^/]+)$/, '_rels/$1.rels');
export function parseRels(xml, fromPath) {
  const out = {};
  for (const m of (xml || '').matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
    const a = attrs(m[1]);
    if (a.TargetMode !== 'External') out[a.Id] = norm(a.Target[0] === '/' ? a.Target : dirOf(fromPath) + a.Target);
  }
  return out;
}

// ---- workbook: sheets + defined names ----
export function parseWorkbook(xml, relsXml) {
  const rels = parseRels(relsXml, 'xl/workbook.xml');
  const sheets = [...xml.matchAll(/<sheet\b([^>]*?)\/?>/g)].map(m => {
    const a = attrs(m[1]);
    return { name: unesc(a.name), path: rels[a['r:id']], hidden: !!a.state && a.state !== 'visible' };
  });
  const names = [];
  for (const m of xml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) {
    const a = attrs(m[1]), t = /^(?:'((?:[^']|'')+)'|([^!']+))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/.exec(unesc(m[2]).trim());
    names.push({
      name: unesc(a.name), hidden: on(a.hidden), local: a.localSheetId != null ? +a.localSheetId : null,
      sheet: t ? (t[1] ? t[1].replace(/''/g, "'") : t[2]) : null, range: t ? parseRange(t[3].replace(/\$/g, '')) : null,
    });
  }
  return { sheets, names, date1904: /date1904="(1|true)"/.test(xml) };
}

// ---- theme colours, in SpreadsheetML theme-index order ----
const THEME_KEYS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
const OFFICE = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
export function parseTheme(xml) {
  const cs = /<a:clrScheme\b[\s\S]*?<\/a:clrScheme>/.exec(xml || '')?.[0] || '';
  return THEME_KEYS.map((k, i) => {
    const m = new RegExp(`<a:${k}>([\\s\\S]*?)</a:${k}>`).exec(cs);
    return (m && (/lastClr="([0-9A-Fa-f]{6})"/.exec(m[1]) || /srgbClr val="([0-9A-Fa-f]{6})"/.exec(m[1]))?.[1]) || OFFICE[i];
  });
}

// ---- styles ----
export const color = s => {
  if (s == null) return null;
  const a = attrs(s), tint = +(a.tint || 0);
  return a.rgb ? { rgb: a.rgb.slice(-6), tint } : a.theme != null ? { theme: +a.theme, tint }
    : a.indexed != null ? { indexed: +a.indexed, tint } : a.auto ? { auto: true } : null;
};
const colorIn = (x, tag = 'color') => color(new RegExp(`<${tag}\\b([^>]*?)\\/?>`).exec(x || '')?.[1]);

// Font properties present in the XML only (tri-state), so it works for cell fonts, rich-text runs and CF overrides.
export function parseFont(x) {
  const f = {};
  const flag = t => { const m = new RegExp(`<${t}(?:\\s+val="([^"]*)")?\\s*/>`).exec(x); if (m) f[t] = m[1] == null || on(m[1]); };
  const val = t => new RegExp(`<${t}\\s+val="([^"]*)"`).exec(x)?.[1];
  flag('b'); flag('i'); flag('strike');
  const name = val('name') ?? val('rFont'), sz = val('sz'), u = val('u'), c = colorIn(x), va = val('vertAlign');
  if (va && va !== 'baseline') f.va = va; // superscript / subscript
  if (name != null) f.name = unesc(name);
  if (sz != null) f.sz = +sz;
  if (u != null) f.u = u !== 'none'; else if (/<u\s*\/>/.test(x)) f.u = true;
  if (c) f.color = c;
  return f;
}

function parseFill(x, dxf) {
  const p = /<patternFill\b([^>]*?)(?:\/>|>([\s\S]*?)<\/patternFill>)/.exec(x);
  if (p) {
    const type = attrs(p[1]).patternType || (dxf ? 'solid' : 'none'), fg = colorIn(p[2], 'fgColor'), bg = colorIn(p[2], 'bgColor');
    if (type === 'none') return null;
    // A CF (dxf) solid fill keeps its colour in bgColor; a normal solid fill uses fgColor.
    return { color: (dxf ? bg || fg : type === 'solid' ? fg || bg : fg || bg) || { auto: true }, pattern: type };
  }
  const g = /<gradientFill\b[\s\S]*?<\/gradientFill>/.exec(x);
  return g ? { color: colorIn(g[0]), pattern: 'solid' } : null;
}

function parseBorder(x) {
  const side = n => {
    const m = new RegExp(`<${n}\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${n}>)`).exec(x), st = m && attrs(m[1]).style;
    return st && st !== 'none' ? { style: st, color: colorIn(m[2]) } : null;
  };
  return { l: side('left') || side('start'), r: side('right') || side('end'), t: side('top'), b: side('bottom') };
}

const block = (xml, tag) => new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml)?.[1] || '';
const items = (xml, tag) => [...xml.matchAll(new RegExp(`<${tag}\\b[^>]*?(?:/>|>([\\s\\S]*?)</${tag}>)`, 'g'))].map(m => m[1] || '');

export function parseStyles(xml) {
  const numFmts = {};
  for (const m of xml.matchAll(/<numFmt\b([^>]*?)\/?>/g)) { const a = attrs(m[1]); numFmts[+a.numFmtId] = unesc(a.formatCode); }
  const fonts = items(block(xml, 'fonts'), 'font').map(parseFont);
  const fills = items(block(xml, 'fills'), 'fill').map(f => parseFill(f, false));
  const borders = items(block(xml, 'borders'), 'border').map(parseBorder);
  const xfs = [...block(xml, 'cellXfs').matchAll(/<xf\b([^>]*?)(?:\/>|>([\s\S]*?)<\/xf>)/g)].map(m => {
    const a = attrs(m[1]), al = /<alignment\b([^>]*?)\/?>/.exec(m[2] || '');
    const g = al ? attrs(al[1]) : {};
    return {
      numFmt: +(a.numFmtId || 0), font: +(a.fontId || 0), fill: +(a.fillId || 0), border: +(a.borderId || 0),
      h: g.horizontal || 'general', v: g.vertical || 'bottom', wrap: on(g.wrapText) || g.horizontal === 'justify' || g.vertical === 'justify',
      indent: +(g.indent || 0), rot: +(g.textRotation || 0), shrink: on(g.shrinkToFit),
    };
  });
  const dxfs = items(block(xml, 'dxfs'), 'dxf').map(parseDxf);
  const indexed = [...block(xml, 'indexedColors').matchAll(/rgb="([0-9A-Fa-f]{8})"/g)].map(m => m[1].slice(2));
  return { numFmts, fonts, fills, borders, xfs, dxfs, indexed: indexed.length ? indexed : null };
}

export function parseDxf(x) {
  const font = /<font\b[^>]*>([\s\S]*?)<\/font>/.exec(x), fill = /<fill\b[^>]*>([\s\S]*?)<\/fill>/.exec(x), border = /<border\b[^>]*>([\s\S]*?)<\/border>/.exec(x);
  return { font: font ? parseFont(font[1]) : null, fill: fill ? parseFill(fill[1], true) : null, border: border ? parseBorder(border[1]) : null };
}

// ---- shared strings (plain or rich text) ----
const texts = x => unesc([...x.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g)].map(m => m[1] || '').join(''));
export function richText(x) {
  if (!x.includes('<r>')) return texts(x);
  const runs = [...x.matchAll(/<r>([\s\S]*?)<\/r>/g)].map(m => {
    const pr = /<rPr>([\s\S]*?)<\/rPr>/.exec(m[1]);
    return { text: texts(m[1].replace(/<rPr>[\s\S]*?<\/rPr>/, '')), font: pr ? parseFont(pr[1]) : null };
  });
  return { text: runs.map(r => r.text).join(''), runs };
}
export const parseSST = xml => [...(xml || '').matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m => richText(m[1]));

// ---- worksheet ----
// Cell values come from the cached results Excel saved, so no formula is ever recalculated.
export function parseSheet(xml, sst) {
  const sd = xml.indexOf('<sheetData'), se = xml.lastIndexOf('</sheetData>');
  const head = sd < 0 ? xml : xml.slice(0, sd), data = sd < 0 ? '' : xml.slice(sd, se), tail = se < 0 ? '' : xml.slice(se);

  const rows = new Map(), cells = new Map();
  for (const m of data.matchAll(/<row\b([^>]*?)\/?>/g)) {
    const a = attrs(m[1]);
    const st = on(a.customFormat) ? +(a.s || 0) : null;
    if (a.ht != null || on(a.hidden) || st != null) rows.set(+a.r, { ht: a.ht != null ? +a.ht : null, hidden: on(a.hidden), s: st });
  }
  for (const m of data.matchAll(/<c ([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
    const a = m[1], rc = / ?r="([A-Z]+)(\d+)"/.exec(a);
    if (!rc) continue;
    const s = +(/ s="(\d+)"/.exec(a)?.[1] || 0), t = / t="(\w+)"/.exec(a)?.[1], body = m[2] || '';
    let v = null;
    if (t === 'inlineStr') v = richText(/<is>([\s\S]*?)<\/is>/.exec(body)?.[1] || '');
    else {
      const raw = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body)?.[1]; // may carry xml:space="preserve"
      if (raw != null) v = t === 's' ? sst[+raw] ?? '' : t === 'str' ? unesc(raw) : t === 'b' ? raw === '1' : t === 'e' ? { error: raw } : +raw;
    }
    if (v != null || s) cells.set(+rc[2] * KEY + colNum(rc[1]), { v, s });
  }

  const cols = [...head.matchAll(/<col\b([^>]*?)\/?>/g)].map(m => {
    const a = attrs(m[1]);
    return { min: +a.min, max: +a.max, width: a.width != null ? +a.width : null, hidden: on(a.hidden), s: a.style != null ? +a.style : null };
  });
  const fp = attrs(/<sheetFormatPr\b([^>]*)>/.exec(head)?.[1] || ''), view = attrs(/<sheetView\b([^>]*)>/.exec(head)?.[1] || '');
  const merges = [...tail.matchAll(/<mergeCell ref="([^"]+)"/g)].map(m => parseRange(m[1])).filter(Boolean);
  // print setup: margins (inches -> points) and manual page breaks (break after row `id`)
  const pm = attrs(/<pageMargins\b([^>]*?)\/?>/.exec(tail)?.[1] || ''), inch = (k, d) => (pm[k] != null ? +pm[k] : d) * 72;
  const margins = { l: inch('left', 0.7), r: inch('right', 0.7), t: inch('top', 0.75), b: inch('bottom', 0.75) };
  const rowBreaks = [...(/<rowBreaks\b[^>]*>([\s\S]*?)<\/rowBreaks>/.exec(tail)?.[1] || '').matchAll(/<brk id="(\d+)"/g)].map(m => +m[1]);

  const cf = [];
  for (const m of tail.matchAll(/<conditionalFormatting\b([^>]*)>([\s\S]*?)<\/conditionalFormatting>/g)) {
    const ranges = attrs(m[1]).sqref.split(/\s+/).map(parseRange).filter(Boolean);
    cf.push({ ranges, rules: [...m[2].matchAll(/<cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/cfRule>)/g)].map(r => cfRule(attrs(r[1]), r[2] || '', /<formula>([\s\S]*?)<\/formula>/g)) });
  }
  // Excel 2010+ rules (cross-sheet references, icon sets) live in the extension list.
  for (const m of tail.matchAll(/<x14:conditionalFormatting\b[^>]*>([\s\S]*?)<\/x14:conditionalFormatting>/g)) {
    const ranges = (/<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(m[1])?.[1] || '').trim().split(/\s+/).map(parseRange).filter(Boolean);
    cf.push({ ranges, rules: [...m[1].matchAll(/<x14:cfRule\b([^>]*?)(?:\/>|>([\s\S]*?)<\/x14:cfRule>)/g)].map(r => {
      const body = r[2] || '', dxf = /<x14:dxf>([\s\S]*?)<\/x14:dxf>/.exec(body);
      const rule = cfRule(attrs(r[1]), body, /<xm:f>([\s\S]*?)<\/xm:f>/g);
      if (dxf) rule.dxf = parseDxf(dxf[1]);
      return rule;
    }) });
  }
  // Sheets named inside CF formulas: the loader fetches these too so the rules can be evaluated.
  const refSheets = new Set();
  for (const b of cf) for (const r of b.rules) for (const f of r.formulas) for (const s of f.matchAll(/(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!/g)) refSheets.add((s[1] || s[2]).replace(/''/g, "'"));

  return {
    cells, rows, cols, merges, cf, refSheets: [...refSheets], margins, rowBreaks,
    defaultColWidth: fp.defaultColWidth != null ? +fp.defaultColWidth : null, baseColWidth: +(fp.baseColWidth || 8),
    defaultRowHeight: +(fp.defaultRowHeight || 15), zeroHeight: on(fp.zeroHeight),
    showGrid: view.showGridLines !== '0' && view.showGridLines !== 'false', showZeros: view.showZeros !== '0' && view.showZeros !== 'false',
    drawingRid: /<drawing r:id="([^"]+)"/.exec(tail)?.[1] || null,
  };
}

function cfRule(a, body, fRe) {
  const rule = {
    type: a.type, op: a.operator, prio: +(a.priority || 0), stop: on(a.stopIfTrue), dxf: a.dxfId != null ? +a.dxfId : null,
    formulas: [...body.replace(/<(?:x14:)?iconSet\b[\s\S]*?<\/(?:x14:)?iconSet>/, '').matchAll(fRe)].map(f => unesc(f[1])),
    text: a.text != null ? unesc(a.text) : null,
    rank: +(a.rank || 10), percent: on(a.percent), bottom: on(a.bottom),
  };
  const icon = /<(?:x14:)?iconSet\b([^>]*)>([\s\S]*?)<\/(?:x14:)?iconSet>/.exec(body);
  if (a.type === 'iconSet' || icon) {
    const ia = attrs(icon?.[1] || '');
    rule.icons = {
      set: ia.iconSet || '3TrafficLights1', reverse: on(ia.reverse), showValue: ia.showValue !== '0',
      cfvo: [...(icon?.[2] || '').matchAll(/<(?:x14:)?cfvo\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:x14:)?cfvo>)/g)].map(c => {
        const ca = attrs(c[1]);
        return { type: ca.type, val: ca.val ?? unesc(/<xm:f>([\s\S]*?)<\/xm:f>/.exec(c[2] || '')?.[1] ?? '0'), gte: ca.gte !== '0' };
      }),
      custom: [...(icon?.[2] || '').matchAll(/<x14:cfIcon\b([^>]*?)\/?>/g)].map(c => { const ca = attrs(c[1]); return { set: ca.iconSet, id: +ca.iconId }; }),
    };
  }
  return rule;
}

// ---- drawings: charts and pictures anchored to cells ----
export function parseDrawing(xml, rels) {
  const out = [];
  for (const m of xml.matchAll(/<xdr:(twoCellAnchor|oneCellAnchor)\b[^>]*>([\s\S]*?)<\/xdr:\1>/g)) {
    const x = m[2];
    const pt = tag => {
      const t = new RegExp(`<xdr:${tag}>([\\s\\S]*?)</xdr:${tag}>`).exec(x);
      const g = n => +(new RegExp(`<xdr:${n}>(-?\\d+)</xdr:${n}>`).exec(t?.[1] || '')?.[1] || 0);
      return t && { c: g('col') + 1, cOff: g('colOff'), r: g('row') + 1, rOff: g('rowOff') };
    };
    const chart = /<c:chart\b[^>]*r:id="([^"]+)"/.exec(x), pic = /<a:blip\b[^>]*r:embed="([^"]+)"/.exec(x);
    if (!chart && !pic) continue; // plain shapes and connector arrows are not drawn
    const ext = attrs(/<xdr:ext\b([^>]*?)\/?>/.exec(x)?.[1] || '');
    out.push({ from: pt('from'), to: pt('to'), ext: ext.cx ? { cx: +ext.cx, cy: +ext.cy } : null, chart: chart ? rels[chart[1]] : null, image: pic ? rels[pic[1]] : null });
  }
  return out;
}
