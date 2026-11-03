// Bar/column charts, drawn from the values Excel caches inside the chart XML (no formulas needed).
// ponytail: only bar/column charts (clustered or stacked, vertical) are drawn, since those are all the reports use.
// Add line/pie/horizontal-bar drawing here when a report needs one.
import { unesc } from './xlsx.js';
import { drawingColorSpec, drawingColor, family, PX, EMU } from './paint.js';
import SSF from './vendor/ssf.js';

const tag = (x, t) => new RegExp(`<c:${t} val="([^"]*)"`).exec(x || '')?.[1];
const inner = (x, t) => new RegExp(`<${t}\\b[^>]*>([\\s\\S]*?)</${t}>`).exec(x || '')?.[1];

// Text properties from <c:txPr> / <c:rich>: first defRPr or rPr wins.
function textProps(x) {
  if (!x) return {};
  const body = /<a:bodyPr\b([^>]*?)\/?>/.exec(x)?.[1] || '', pr = /<a:(?:defRPr|rPr)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:(?:defRPr|rPr)>)/.exec(x);
  const rot = /rot="(-?\d+)"/.exec(body)?.[1], sz = pr && /\bsz="(\d+)"/.exec(pr[1])?.[1];
  return {
    rot: rot != null && Math.abs(+rot) <= 5400000 ? +rot / 60000 : null, // -60000000 means "automatic"
    sz: sz ? +sz / 100 : null, b: pr ? /\bb="1"/.test(pr[1]) : false, i: pr ? /\bi="1"/.test(pr[1]) : false,
    caps: pr ? /\bcap="all"/.test(pr[1]) : false, color: drawingColorSpec(inner(pr?.[2], 'a:solidFill')),
  };
}
function richText(x) {
  if (!x) return null;
  const rich = inner(x, 'c:rich');
  if (!rich) return null;
  const text = [...rich.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)].map(p => [...p[1].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map(t => unesc(t[1])).join('')).join('\n');
  return { text, ...textProps(rich) };
}
const fillOf = sp => sp == null ? null : sp.replace(/<a:ln\b[\s\S]*?<\/a:ln>/g, '').includes('<a:noFill/>') ? 'none' : drawingColorSpec(inner(sp.replace(/<a:ln\b[\s\S]*?<\/a:ln>/g, ''), 'a:solidFill'));
const lineOf = sp => {
  const ln = /<a:ln\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:ln>)/.exec(sp || '');
  if (!ln) return null;
  if ((ln[2] || '').includes('<a:noFill/>')) return 'none';
  return { color: drawingColorSpec(inner(ln[2], 'a:solidFill')), w: Math.max(1, +(/w="(\d+)"/.exec(ln[1])?.[1] || 9525) / EMU) };
};
const points = x => { const o = []; for (const p of (x || '').matchAll(/<c:pt idx="(\d+)"[^>]*>\s*<c:v>([\s\S]*?)<\/c:v>/g)) o[+p[1]] = unesc(p[2]); return o; };
const fmtOf = x => unesc(/<c:formatCode>([\s\S]*?)<\/c:formatCode>/.exec(x || '')?.[1] || 'General');
function labelsOf(x) {
  if (!x) return null;
  const nf = /<c:numFmt formatCode="([^"]*)" sourceLinked="0"/.exec(x);
  return { show: tag(x, 'showVal') === '1' && tag(x, 'delete') !== '1', pos: tag(x, 'dLblPos') || 'outEnd', fmt: nf ? unesc(nf[1]) : null, tx: textProps(inner(x, 'c:txPr')) };
}

export function parseChart(xml) {
  const x = xml.replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '');
  const p0 = x.indexOf('<c:plotArea>'), p1 = x.indexOf('</c:plotArea>');
  if (p0 < 0) return null;
  const head = x.slice(0, p0), plot = x.slice(p0, p1), after = x.slice(p1);
  const bar = inner(plot, 'c:barChart');
  if (!bar) return null;
  const series = [...bar.matchAll(/<c:ser>([\s\S]*?)<\/c:ser>/g)].map(m => {
    const s = m[1], cat = inner(s, 'c:cat'), val = inner(s, 'c:val');
    return {
      name: unesc(/<c:v>([\s\S]*?)<\/c:v>/.exec(inner(s, 'c:tx') || '')?.[1] || ''),
      fill: fillOf(inner(s.replace(/<c:dLbls>[\s\S]*?<\/c:dLbls>/, '').replace(/<c:dPt>[\s\S]*?<\/c:dPt>/g, ''), 'c:spPr')),
      cats: points(cat), catFmt: fmtOf(cat), vals: points(val).map(v => v == null || v === '' ? null : +v), fmt: fmtOf(val),
      labels: labelsOf(inner(s, 'c:dLbls')),
    };
  });
  const axis = t => {
    const a = inner(plot, t);
    if (!a) return null;
    const noTitle = a.replace(/<c:title>[\s\S]*?<\/c:title>/, ''), nf = /<c:numFmt formatCode="([^"]*)" sourceLinked="([01])"/.exec(noTitle);
    return {
      del: tag(a, 'delete') === '1', title: richText(inner(a, 'c:title')), tx: textProps(inner(noTitle, 'c:txPr')),
      fmt: nf && nf[2] === '0' ? unesc(nf[1]) : null, min: tag(a, 'min'), max: tag(a, 'max'), major: tag(a, 'majorUnit'),
      grid: a.includes('<c:majorGridlines'), gridLine: lineOf(inner(inner(a, 'c:majorGridlines'), 'c:spPr')),
      line: lineOf(inner(noTitle.replace(/<c:majorGridlines>[\s\S]*?<\/c:majorGridlines>/, ''), 'c:spPr')), noLabels: tag(a, 'tickLblPos') === 'none',
    };
  };
  const titleXml = inner(head, 'c:title'), legend = inner(after, 'c:legend'), space = /<\/c:chart>\s*<c:spPr>([\s\S]*?)<\/c:spPr>/.exec(x)?.[1];
  let title = tag(head, 'autoTitleDeleted') === '1' && !titleXml ? null : richText(titleXml);
  if (titleXml && !title?.text && series.length === 1) title = { ...textProps(inner(titleXml, 'c:txPr')), text: series[0].name };
  return {
    dir: tag(bar, 'barDir') || 'col', grouping: tag(bar, 'grouping') || 'clustered', gap: +(tag(bar, 'gapWidth') ?? 150), overlap: tag(bar, 'overlap'),
    series, labels: labelsOf(inner(bar.replace(/<c:ser>[\s\S]*?<\/c:ser>/g, ''), 'c:dLbls')),
    cat: axis('c:catAx') || axis('c:dateAx'), val: axis('c:valAx'), title,
    legend: legend ? { pos: tag(legend, 'legendPos') || 'r', tx: textProps(inner(legend, 'c:txPr')) } : null,
    bg: fillOf(space), border: lineOf(space), tx: textProps(inner(after.replace(/<c:legend>[\s\S]*?<\/c:legend>/, ''), 'c:txPr')),
  };
}

const fmt = (code, v) => { try { return SSF.format(code || 'General', v); } catch { return String(v); } };
// Excel-like automatic axis: ~5% headroom, then the largest 1/2/5 step that still gives at least 5 intervals.
function autoScale(lo, hi) {
  if (hi === lo) hi = lo + 1;
  const top = hi > 0 ? hi + 0.05 * (hi - lo) : hi, bot = lo < 0 ? lo - 0.05 * (hi - lo) : lo, span = top - bot;
  let step = 10 ** Math.ceil(Math.log10(span)), k = 0;
  while (span / step < 5) step *= [0.5, 0.4, 0.5][k++ % 3]; // 10 -> 5 -> 2 -> 1 -> 0.5 ...
  return { min: Math.floor(bot / step) * step, max: Math.ceil(top / step) * step, step };
}

export function drawChart(ctx, ch, X, Y, W, H, theme) {
  const col = (spec, d) => drawingColor(spec, theme, d);
  const font = (tx, dflt, base = ch.tx) => ({ size: (tx?.sz ?? base?.sz ?? dflt) * PX, b: tx?.b ?? false, i: tx?.i ?? false, caps: tx?.caps, color: col(tx?.color ?? base?.color, '#595959') });
  const setFont = f => { ctx.font = `${f.i ? 'italic ' : ''}${f.b ? 'bold ' : ''}${f.size.toFixed(2)}px ${family('Calibri')}`; ctx.fillStyle = f.color; };
  const label = (f, s) => f.caps ? String(s).toUpperCase() : String(s);

  ctx.save();
  ctx.beginPath(); ctx.rect(X, Y, W, H); ctx.clip();
  if (ch.bg !== 'none') { ctx.fillStyle = col(ch.bg, '#ffffff'); ctx.fillRect(X, Y, W, H); }
  if (ch.border !== 'none') { ctx.strokeStyle = col(ch.border?.color, '#d9d9d9'); ctx.lineWidth = ch.border?.w || 1; ctx.strokeRect(X + 0.5, Y + 0.5, W - 1, H - 1); }

  const pad = 8;
  let top = Y + pad, bottom = Y + H - pad, left = X + pad, right = X + W - pad;
  ctx.textBaseline = 'alphabetic';

  if (ch.title?.text) {
    const f = font(ch.title, 14);
    setFont(f); ctx.textAlign = 'center';
    for (const line of label(f, ch.title.text).split('\n')) { top += f.size; ctx.fillText(line, X + W / 2, top); top += f.size * 0.25; }
    top += 4;
  }
  const series = ch.series, legend = ch.legend;
  if (legend && series.length) {
    const f = font(legend.tx, 9), key = f.size * 0.6;
    setFont(f);
    const items = series.map((s, i) => ({ s, i, w: key + 4 + ctx.measureText(label(f, s.name)).width }));
    const drawItem = (it, x, y) => {
      ctx.fillStyle = seriesColor(it.s, it.i); ctx.fillRect(x, y - key, key, key);
      ctx.fillStyle = f.color; ctx.textAlign = 'left'; ctx.fillText(label(f, it.s.name), x + key + 4, y);
    };
    if (legend.pos === 't' || legend.pos === 'b') {
      const total = items.reduce((s, it) => s + it.w + 12, -12);
      let x = X + (W - total) / 2;
      const y = legend.pos === 't' ? top + f.size : bottom;
      for (const it of items) { drawItem(it, x, y); x += it.w + 12; }
      if (legend.pos === 't') top += f.size * 1.4 + 4; else bottom -= f.size * 1.4 + 4;
    } else {
      const wMax = Math.max(...items.map(it => it.w));
      let y = (top + bottom) / 2 - (items.length * f.size * 1.4) / 2 + f.size;
      for (const it of items) { drawItem(it, right - wMax, y); y += f.size * 1.4; }
      right -= wMax + 10;
    }
  }
  function seriesColor(s, i) {
    if (s.fill && s.fill !== 'none') return col(s.fill, '#4472c4');
    return '#' + theme[4 + (i % 6)];
  }

  const valAx = ch.val || {}, catAx = ch.cat || {};
  const valTitle = valAx.title?.text && !valAx.del ? valAx.title : null, catTitle = catAx.title?.text && !catAx.del ? catAx.title : null;
  const fVT = font(valTitle, 10), fCT = font(catTitle, 10);
  if (valTitle) left += fVT.size * 1.3;
  if (catTitle) bottom -= fCT.size * 1.3;

  // value scale
  const stacked = ch.grouping === 'stacked' || ch.grouping === 'percentStacked';
  let lo = 0, hi = 0;
  const nCat = Math.max(0, ...series.map(s => Math.max(s.vals.length, s.cats.length)));
  if (stacked) for (let k = 0; k < nCat; k++) { let p = 0, n = 0; for (const s of series) { const v = s.vals[k]; if (v > 0) p += v; else if (v < 0) n += v; } hi = Math.max(hi, p); lo = Math.min(lo, n); }
  else for (const s of series) for (const v of s.vals) if (v != null) { hi = Math.max(hi, v); lo = Math.min(lo, v); }
  const sc = autoScale(lo, hi);
  const vmin = valAx.min != null ? +valAx.min : sc.min, vmax = valAx.max != null ? +valAx.max : sc.max, step = valAx.major != null ? +valAx.major : sc.step;

  const fVL = font(valAx.tx, 9), valFmt = valAx.fmt || series[0]?.fmt;
  const ticks = [];
  for (let v = vmin; v <= vmax + step / 1e6; v += step) ticks.push(+v.toPrecision(12));
  setFont(fVL);
  const valW = valAx.del || valAx.noLabels ? 0 : Math.max(...ticks.map(v => ctx.measureText(fmt(valFmt, v)).width));
  left += valW + 6;

  // category labels: horizontal if they fit, otherwise slanted (-45) like Excel's auto rotation
  const fCL = font(catAx.tx, 9), cats = Array.from({ length: nCat }, (_, k) => {
    const raw = series.find(s => s.cats[k] != null)?.cats[k];
    return raw == null ? String(k + 1) : label(fCL, /^-?\d+(\.\d+)?$/.test(raw) && series[0].catFmt !== 'General' ? fmt(series[0].catFmt, +raw) : raw);
  });
  setFont(fCL);
  const slot = (right - left) / Math.max(1, nCat), catW = Math.max(0, ...cats.map(s => ctx.measureText(s).width));
  const rot = catAx.tx?.rot ?? (catW > slot - 4 ? -45 : 0), rad = rot * Math.PI / 180;
  const catH = catAx.del || catAx.noLabels ? 0 : rot ? Math.abs(Math.sin(rad)) * catW + Math.abs(Math.cos(rad)) * fCL.size : fCL.size * 1.2;
  bottom -= catH + 6;
  const pw = right - left, ph = bottom - top, yOf = v => bottom - (v - vmin) / (vmax - vmin) * ph;

  // gridlines, axis line
  ctx.lineWidth = 1;
  if (valAx.grid !== false && valAx.gridLine !== 'none') {
    ctx.strokeStyle = col(valAx.gridLine?.color, '#d9d9d9');
    for (const v of ticks) { const y = Math.round(yOf(v)) + 0.5; ctx.beginPath(); ctx.moveTo(left, y); ctx.lineTo(right, y); ctx.stroke(); }
  }
  if (catAx.line !== 'none') { ctx.strokeStyle = col(catAx.line?.color, '#d9d9d9'); const y = Math.round(yOf(Math.max(vmin, Math.min(0, vmax)))) + 0.5; ctx.beginPath(); ctx.moveTo(left, y); ctx.lineTo(right, y); ctx.stroke(); }
  if (!valAx.del && !valAx.noLabels) {
    setFont(fVL); ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
    for (const v of ticks) ctx.fillText(fmt(valFmt, v), left - 6, yOf(v));
  }

  // bars (Excel geometry: bar width from gap width and overlap)
  const n = series.length, ov = (ch.overlap != null ? +ch.overlap : stacked ? 100 : 0) / 100, g = ch.gap / 100;
  const barW = stacked ? slot / (1 + g) : slot / (1 + (n - 1) * (1 - ov) + g), stepX = barW * (1 - ov), groupW = stacked ? barW : barW + (n - 1) * stepX;
  const zero = yOf(Math.max(vmin, Math.min(0, vmax))), bars = [];
  for (let k = 0; k < nCat; k++) {
    const x0 = left + slot * k + (slot - groupW) / 2;
    let pos = 0, neg = 0;
    series.forEach((s, i) => {
      const v = s.vals[k];
      if (v == null) return;
      const x = stacked ? x0 : x0 + i * stepX;
      let y1 = zero, y2 = yOf(v);
      if (stacked) { const base = v >= 0 ? pos : neg; y1 = yOf(base); y2 = yOf(base + v); v >= 0 ? (pos += v) : (neg += v); }
      ctx.fillStyle = seriesColor(s, i);
      ctx.fillRect(x, Math.min(y1, y2), barW, Math.abs(y2 - y1));
      bars.push({ s, v, x: x + barW / 2, y: y2, base: y1 });
    });
  }
  // data labels
  for (const b of bars) {
    const L = b.s.labels ?? ch.labels;
    if (!L?.show) continue;
    const f = font(L.tx, 9);
    setFont(f);
    const t = fmt(L.fmt || b.s.fmt, b.v), up = b.v >= 0, vertical = L.tx?.rot != null && Math.abs(L.tx.rot) >= 45;
    const y = L.pos === 'ctr' ? (b.y + b.base) / 2 : L.pos === 'inEnd' ? b.y + (up ? 3 : -3) : L.pos === 'inBase' ? b.base + (up ? -3 : 3) : b.y + (up ? -3 : 3);
    const away = L.pos === 'outEnd' || !L.pos ? up : !up;
    ctx.save(); ctx.translate(b.x, y);
    if (vertical) { ctx.rotate(-Math.PI / 2); ctx.textAlign = L.pos === 'ctr' ? 'center' : away ? 'left' : 'right'; ctx.textBaseline = 'middle'; }
    else { ctx.textAlign = 'center'; ctx.textBaseline = L.pos === 'ctr' ? 'middle' : away ? 'bottom' : 'top'; }
    ctx.fillText(t, 0, 0);
    ctx.restore();
  }
  // category labels
  if (!catAx.del && !catAx.noLabels) {
    setFont(fCL); ctx.textBaseline = rot ? 'middle' : 'top';
    cats.forEach((s, k) => {
      const cx = left + slot * (k + 0.5), cy = bottom + 6;
      if (!rot) { ctx.textAlign = 'center'; ctx.fillText(s, cx, cy); return; }
      ctx.save(); ctx.translate(cx, cy); ctx.rotate(rad); ctx.textAlign = 'right'; ctx.fillText(s, 0, 0); ctx.restore();
    });
  }
  // axis titles
  if (valTitle) { setFont(fVT); ctx.save(); ctx.translate(X + pad + fVT.size * 0.5, (top + bottom) / 2); ctx.rotate(-Math.PI / 2); ctx.textAlign = 'center'; ctx.textBaseline = 'middle'; ctx.fillText(label(fVT, valTitle.text), 0, 0); ctx.restore(); }
  if (catTitle) { setFont(fCT); ctx.textAlign = 'center'; ctx.textBaseline = 'bottom'; ctx.fillText(label(fCT, catTitle.text), left + pw / 2, Y + H - pad - (legend?.pos === 'b' ? fCT.size * 1.4 : 0)); }
  ctx.restore();
}
