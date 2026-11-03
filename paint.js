// Colour and font helpers shared by the cell renderer and the chart renderer.
import { attrs } from './xlsx.js';

export const PX = 96 / 72; // points -> CSS pixels
export const EMU = 9525;   // drawing units per CSS pixel

// Excel's legacy 64-colour palette (indexed colours); 64/65 are "system" colours and fall back to the default.
const INDEXED = ('000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF ' +
  '800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF ' +
  '000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 ' +
  '3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333').split(' ');

const hex2rgb = h => [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16));
const rgb2hex = rgb => rgb.map(v => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('');
function toHsl(hex) {
  const [r, g, b] = hex2rgb(hex).map(v => v / 255), mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
  if (mx === mn) return [0, 0, l];
  const d = mx - mn, s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
  const h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
  return [h / 6, s, l];
}
function fromHsl(h, s, l) {
  if (!s) return rgb2hex([l * 255, l * 255, l * 255]);
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const f = t => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  return rgb2hex([f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255]);
}
const withL = (hex, fn) => { const [h, s, l] = toHsl(hex); return fromHsl(h, s, Math.max(0, Math.min(1, fn(l)))); };

// SpreadsheetML colour {rgb | theme | indexed, tint} -> '#rrggbb'
export function cellColor(c, book, dflt) {
  if (!c || c.auto) return dflt;
  const hex = c.rgb ?? (c.theme != null ? book.theme[c.theme] : c.indexed != null ? (book.styles.indexed?.[c.indexed] ?? INDEXED[c.indexed]) : null);
  if (!hex) return dflt;
  return '#' + (c.tint ? withL(hex, l => c.tint < 0 ? l * (1 + c.tint) : l * (1 - c.tint) + c.tint) : hex);
}

// DrawingML colour (charts) -> spec, then spec + theme -> '#rrggbb'
export function drawingColorSpec(x) {
  const m = /<a:(srgbClr|schemeClr|sysClr|prstClr)\b([^>]*?)(?:\/>|>([\s\S]*?)<\/a:\1>)/.exec(x || '');
  if (!m) return null;
  const a = attrs(m[2]), mods = {};
  for (const t of (m[3] || '').matchAll(/<a:(lumMod|lumOff|tint|shade)\s+val="(\d+)"/g)) mods[t[1]] = +t[2] / 100000;
  return { kind: m[1], val: m[1] === 'sysClr' ? a.lastClr : a.val, mods };
}
const SCHEME = { lt1: 0, bg1: 0, dk1: 1, tx1: 1, lt2: 2, bg2: 2, dk2: 3, tx2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9, hlink: 10, folHlink: 11 };
const PRESET = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080' };
export function drawingColor(spec, theme, dflt) {
  if (!spec) return dflt;
  let hex = spec.kind === 'schemeClr' ? theme[SCHEME[spec.val] ?? 1] : spec.kind === 'prstClr' ? PRESET[spec.val] ?? '000000' : spec.val;
  if (!hex) return dflt;
  const { lumMod, lumOff, tint, shade } = spec.mods;
  if (lumMod != null || lumOff != null) hex = withL(hex, l => l * (lumMod ?? 1) + (lumOff ?? 0));
  if (tint != null) hex = rgb2hex(hex2rgb(hex).map(v => 255 - (255 - v) * tint));
  if (shade != null) hex = rgb2hex(hex2rgb(hex).map(v => v * shade));
  return '#' + hex;
}

// Excel font name -> CSS font stack. Installed Office fonts win; otherwise metric-compatible web fonts keep
// text the same width (Carlito = Calibri, Tinos = Times New Roman, Arimo = Arial, Cousine = Courier New).
const FALLBACK = { calibri: 'Carlito', 'calibri light': 'Carlito', 'times new roman': 'Tinos', arial: 'Arimo', 'arial mt': 'Arimo', 'arial narrow': 'Arimo', 'courier new': 'Cousine' };
const SERIF = /times|cambria|georgia|bookman|garamond|antiqua|century|castellar|copperplate|roman|serif/i;
export const WEB_FONTS = ['Carlito', 'Tinos', 'Arimo', 'Cousine'];
export function family(name) {
  const n = String(name || 'Calibri').replace(/\0/g, '').replace(/\s*\(body\)\s*/i, '').trim() || 'Calibri';
  const fb = FALLBACK[n.toLowerCase()];
  return `"${n}", ${fb ? `"${fb}", ` : ''}${SERIF.test(n) ? '"Tinos", serif' : '"Carlito", sans-serif'}`;
}
export const cssFont = (f, size = (f.sz || 11) * PX) => `${f.i ? 'italic ' : ''}${f.b ? 'bold ' : ''}${size.toFixed(2)}px ${family(f.name)}`;
