// Conditional formatting: a small Excel formula evaluator plus rule matching.
// It covers what CF rules use in practice: comparisons, arithmetic, &, AND/OR/IF/TRIM/SUM-style functions,
// relative/absolute references and references to other sheets. A rule it cannot read is skipped.
import { KEY, colNum } from './xlsx.js';

const TOK = /\s*(?:("(?:[^"]|"")*")|([A-Za-z_][\w.]*)\(|(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!|(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)|(\d+(?:\.\d*)?(?:[eE][+-]?\d+)?|\.\d+)|(TRUE|FALSE)\b|(<>|<=|>=|[-+*/^&=<>%(),]))/y;

function tokenize(src) {
  const s = src.trim().replace(/^=/, ''), out = [];
  let sheet = null;
  TOK.lastIndex = 0;
  while (TOK.lastIndex < s.length) {
    const at = TOK.lastIndex, m = TOK.exec(s);
    if (!m) { if (/^\s*$/.test(s.slice(at))) break; throw new Error(`unsupported: ${s.slice(at)}`); }
    if (m[1] != null) out.push({ k: 'str', v: m[1].slice(1, -1).replace(/""/g, '"') });
    else if (m[2] != null) out.push({ k: 'fn', v: m[2].toUpperCase() });
    else if (m[3] != null || m[4] != null) sheet = (m[3] ?? m[4]).replace(/''/g, "'");
    else if (m[5] != null) { out.push({ k: 'ref', v: m[5].toUpperCase(), sheet }); sheet = null; }
    else if (m[6] != null) out.push({ k: 'num', v: +m[6] });
    else if (m[7] != null) out.push({ k: 'bool', v: m[7].toUpperCase() === 'TRUE' });
    else out.push({ k: 'op', v: m[8] });
  }
  return out;
}

const ERR = { error: '#VALUE!' };
export const isErr = v => v != null && typeof v === 'object' && !Array.isArray(v);
const num = v => v == null ? 0 : typeof v === 'number' ? v : typeof v === 'boolean' ? +v : typeof v === 'string' && v.trim() !== '' ? Number(v) : NaN;
const str = v => v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
const fin = n => Number.isFinite(n) ? n : ERR;
const first = v => Array.isArray(v) ? v[0] : v;
export const truthy = v => { v = first(v); return typeof v === 'boolean' ? v : typeof v === 'number' ? v !== 0 : false; };

// Excel ordering: numbers < text < booleans; text compares case-insensitively; a blank cell acts as 0 / "" / FALSE.
const kind = v => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
export function cmp(a, b) {
  if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const d = kind(a) - kind(b);
  if (d) return d;
  if (typeof a === 'string') { a = a.toLowerCase(); b = b.toLowerCase(); }
  return a < b ? -1 : a > b ? 1 : 0;
}

function bin(op, a, b) {
  if (Array.isArray(a) || Array.isArray(b)) {
    const n = Math.max(Array.isArray(a) ? a.length : 1, Array.isArray(b) ? b.length : 1);
    return Array.from({ length: n }, (_, i) => bin(op, Array.isArray(a) ? a[i] : a, Array.isArray(b) ? b[i] : b));
  }
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  switch (op) {
    case '+': return fin(num(a) + num(b));
    case '-': return fin(num(a) - num(b));
    case '*': return fin(num(a) * num(b));
    case '/': return fin(num(a) / num(b));
    case '^': return fin(num(a) ** num(b));
    case '&': return str(a) + str(b);
  }
  const c = cmp(a, b);
  return op === '=' ? c === 0 : op === '<>' ? c !== 0 : op === '<' ? c < 0 : op === '>' ? c > 0 : op === '<=' ? c <= 0 : c >= 0;
}

// Numbers from arguments: a range contributes only its numbers; a direct argument is coerced.
const nums = args => { const o = []; for (const a of args) { if (Array.isArray(a)) { for (const x of a) if (typeof x === 'number') o.push(x); else if (isErr(x)) return ERR; } else if (isErr(a)) return a; else if (a != null) o.push(num(a)); } return o.some(Number.isNaN) ? ERR : o; };
const agg = f => a => { const v = nums(a); return isErr(v) ? v : f(v); };
const logic = f => a => { const v = a.flat().filter(x => x != null && typeof x !== 'string'); return v.some(isErr) ? ERR : v.length ? f(v.map(x => num(x) !== 0)) : ERR; };
const round = (x, d) => { const p = 10 ** d; return Math.sign(x) * Math.round(Math.abs(x) * p) / p; };
const FN = {
  AND: logic(v => v.every(Boolean)), OR: logic(v => v.some(Boolean)), NOT: a => !truthy(a[0]),
  IF: a => truthy(a[0]) ? a[1] ?? true : a[2] ?? false, IFERROR: a => isErr(first(a[0])) ? a[1] : a[0],
  ISBLANK: a => first(a[0]) == null, ISNUMBER: a => typeof first(a[0]) === 'number', ISTEXT: a => typeof first(a[0]) === 'string',
  ISERROR: a => isErr(first(a[0])), TRIM: a => str(first(a[0])).trim().replace(/ +/g, ' '), LEN: a => str(first(a[0])).length,
  UPPER: a => str(first(a[0])).toUpperCase(), LOWER: a => str(first(a[0])).toLowerCase(),
  ABS: a => fin(Math.abs(num(first(a[0])))), ROUND: a => fin(round(num(first(a[0])), num(first(a[1])))),
  SUM: agg(v => v.reduce((s, x) => s + x, 0)), MIN: agg(v => v.length ? Math.min(...v) : 0), MAX: agg(v => v.length ? Math.max(...v) : 0),
  AVERAGE: agg(v => v.length ? v.reduce((s, x) => s + x, 0) / v.length : ERR), COUNT: agg(v => v.length),
};

const PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
const refPart = s => { const m = /^(\$?)([A-Z]+)(\$?)(\d+)$/.exec(s); return { c: colNum(m[2]), absC: !!m[1], r: +m[4], absR: !!m[3] }; };

// Compiles a formula into fn(ctx); ctx = { dr, dc, get(sheet|null, row, col) }. dr/dc shift relative references.
export function compile(src) {
  const t = tokenize(src), sheets = new Set();
  let i = 0;
  const peek = () => t[i], take = () => t[i++];
  const expect = v => { if (take()?.v !== v) throw new Error('bad formula'); };
  function ref(x) {
    if (x.sheet) sheets.add(x.sheet);
    const [a, b] = x.v.split(':').map(refPart), sh = x.sheet;
    const at = (p, ctx) => [p.absR ? p.r : p.r + ctx.dr, p.absC ? p.c : p.c + ctx.dc];
    if (!b) return ctx => { const [r, c] = at(a, ctx); return ctx.get(sh, r, c); };
    return ctx => {
      const [r1, c1] = at(a, ctx), [r2, c2] = at(b, ctx), out = [];
      for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) out.push(ctx.get(sh, r, c));
      return out;
    };
  }
  function primary() {
    const x = take();
    if (!x) throw new Error('bad formula');
    if (x.k === 'num' || x.k === 'str' || x.k === 'bool') { const v = x.v; return () => v; }
    if (x.k === 'ref') return ref(x);
    if (x.k === 'fn') {
      const f = FN[x.v];
      if (!f) throw new Error(`unsupported function ${x.v}`);
      const args = [];
      if (peek()?.v !== ')') do args.push(expr(0)); while (peek()?.v === ',' && take());
      expect(')');
      return ctx => f(args.map(e => e(ctx)));
    }
    if (x.v === '(') { const e = expr(0); expect(')'); return e; }
    if (x.v === '-') { const e = postfix(primary()); return ctx => bin('-', 0, e(ctx)); }
    if (x.v === '+') return postfix(primary());
    throw new Error('bad formula');
  }
  function postfix(e) { while (peek()?.v === '%') { take(); const f = e; e = ctx => bin('/', f(ctx), 100); } return e; }
  function expr(min) {
    let left = postfix(primary());
    for (;;) {
      const op = peek(), p = op?.k === 'op' && PREC[op.v];
      if (!p || p < min) return left;
      take();
      const l = left, r = expr(op.v === '^' ? p : p + 1);
      left = ctx => bin(op.v, l(ctx), r(ctx));
    }
  }
  const fn = expr(0);
  if (i < t.length) throw new Error('bad formula');
  fn.sheets = [...sheets];
  return fn;
}

const OPS = {
  equal: (v, a) => cmp(v, a) === 0, notEqual: (v, a) => cmp(v, a) !== 0,
  greaterThan: (v, a) => cmp(v, a) > 0, lessThan: (v, a) => cmp(v, a) < 0,
  greaterThanOrEqual: (v, a) => cmp(v, a) >= 0, lessThanOrEqual: (v, a) => cmp(v, a) <= 0,
  between: (v, a, b) => { const [lo, hi] = cmp(a, b) <= 0 ? [a, b] : [b, a]; return cmp(v, lo) >= 0 && cmp(v, hi) <= 0; },
  notBetween: (v, a, b) => !OPS.between(v, a, b),
};

const plain = v => v == null ? null : typeof v === 'object' && !('error' in v) ? v.text : v;
const numbersIn = (sheet, ranges) => {
  const out = [];
  for (const [k, cell] of sheet.cells) {
    if (typeof cell.v !== 'number') continue;
    const r = Math.floor(k / KEY), c = k % KEY;
    if (ranges.some(q => r >= q[0] && r <= q[2] && c >= q[1] && c <= q[3])) out.push(cell.v);
  }
  return out;
};

function makeTest(book, sheet, rule, blk) {
  const base = blk.ranges[0]; // relative references are written for the first range's top-left cell
  const get = (sh, r, c) => plain((sh ? book.sheet(sh) : sheet)?.cells.get(r * KEY + c)?.v);
  const val = (r, c) => get(null, r, c);
  const fs = rule.formulas.map(f => compile(f));
  if (fs.some(f => f.sheets.some(s => !book.sheet(s)))) return null; // referenced sheet not in the file
  const ev = (f, r, c) => first(f({ dr: r - base[0], dc: c - base[1], get }));
  const has = (v, t) => str(v).toLowerCase().includes(t);
  const text = (rule.text || '').toLowerCase();
  switch (rule.type) {
    case 'expression': return (r, c) => truthy(ev(fs[0], r, c));
    case 'cellIs': {
      const op = OPS[rule.op];
      return op && ((r, c) => {
        const v = val(r, c), a = ev(fs[0], r, c), b = fs[1] && ev(fs[1], r, c);
        return !isErr(v) && !isErr(a) && !isErr(b) && op(v, a, b);
      });
    }
    case 'containsText': return (r, c) => has(val(r, c), text);
    case 'notContainsText': return (r, c) => !has(val(r, c), text);
    case 'beginsWith': return (r, c) => str(val(r, c)).toLowerCase().startsWith(text);
    case 'endsWith': return (r, c) => str(val(r, c)).toLowerCase().endsWith(text);
    case 'containsBlanks': return (r, c) => str(val(r, c)).trim() === '';
    case 'notContainsBlanks': return (r, c) => str(val(r, c)).trim() !== '';
    case 'containsErrors': return (r, c) => isErr(val(r, c));
    case 'notContainsErrors': return (r, c) => !isErr(val(r, c));
    case 'top10': {
      const v = numbersIn(sheet, blk.ranges).sort((a, b) => rule.bottom ? a - b : b - a);
      const n = Math.max(1, rule.percent ? Math.floor(v.length * rule.rank / 100) : rule.rank), th = v[Math.min(n, v.length) - 1];
      return v.length && ((r, c) => { const x = val(r, c); return typeof x === 'number' && (rule.bottom ? x <= th : x >= th); });
    }
    case 'iconSet': {
      const ic = rule.icons, v = numbersIn(sheet, blk.ranges), lo = Math.min(...v), hi = Math.max(...v);
      const sorted = [...v].sort((a, b) => a - b);
      const th = ic.cfvo.map(o => o.type === 'percent' ? lo + (hi - lo) * +o.val / 100 : o.type === 'percentile' ? sorted[Math.round((sorted.length - 1) * +o.val / 100)] : o.type === 'min' ? lo : o.type === 'max' ? hi : +o.val);
      return (r, c) => {
        const x = val(r, c);
        if (typeof x !== 'number') return false;
        let idx = 0;
        for (let k = 1; k < th.length; k++) if (ic.cfvo[k].gte ? x >= th[k] : x > th[k]) idx = k;
        if (ic.reverse) idx = th.length - 1 - idx;
        const cu = ic.custom[idx], icon = cu ? { set: cu.set, id: cu.id } : { set: ic.set, id: idx };
        return { icon: icon.set === 'NoIcons' ? null : { ...icon, showValue: ic.showValue }, hideValue: !ic.showValue };
      };
    }
  }
  return null;
}

// Map cellKey -> merged override { font, fill, border, icon, hideValue } for the cells of `rng` that match a rule.
export function cfOverrides(book, sheet, rng) {
  const [R1, C1, R2, C2] = rng, hits = new Map();
  for (const blk of sheet.cf) {
    if (!blk.ranges.length) continue;
    for (const rule of blk.rules) {
      let test;
      try { test = makeTest(book, sheet, rule, blk); } catch { continue; }
      if (!test) continue;
      for (const q of blk.ranges) {
        const r1 = Math.max(q[0], R1), r2 = Math.min(q[2], R2), c1 = Math.max(q[1], C1), c2 = Math.min(q[3], C2);
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
          const res = test(r, c);
          if (!res) continue;
          const k = r * KEY + c, list = hits.get(k) || hits.set(k, []).get(k);
          list.push({ prio: rule.prio, stop: rule.stop, dxf: typeof rule.dxf === 'number' ? book.styles.dxf(rule.dxf) : rule.dxf, res });
        }
      }
    }
  }
  const out = new Map();
  for (const [k, list] of hits) {
    list.sort((a, b) => a.prio - b.prio);
    const o = { font: {}, fill: undefined, border: {}, icon: null, hideValue: false };
    for (const h of list) {
      if (h.res.icon !== undefined && !o.icon) { o.icon = h.res.icon; o.hideValue = h.res.hideValue; }
      const d = h.dxf;
      if (d?.font) for (const p in d.font) if (!(p in o.font)) o.font[p] = d.font[p];
      if (d?.fill && o.fill === undefined) o.fill = d.fill;
      if (d?.border) for (const s of ['l', 'r', 't', 'b']) if (d.border[s] && !o.border[s]) o.border[s] = d.border[s];
      if (h.stop) break;
    }
    out.set(k, o);
  }
  return out;
}
