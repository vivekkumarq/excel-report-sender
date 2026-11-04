// Metric-compatible stand-ins for the Office fonts, served with the site (SIL Open Font License, see
// vendor/fonts/*-OFL.txt): Carlito = Calibri, Tinos = Times New Roman, Arimo = Arial, Cousine = Courier New.
// Installed Office fonts still win (see family() in paint.js); these keep text the same width everywhere else.
// Each face comes in two files, as on Google Fonts: "latin" and "latin-ext" (accented letters, ₹ and so on).
const LATIN = 'U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, U+0304, U+0308, U+0329, U+2000-206F, U+20AC, U+2122, U+2191, U+2193, U+2212, U+2215, U+FEFF, U+FFFD';
const EXT = 'U+0100-02BA, U+02BD-02C5, U+02C7-02CC, U+02CE-02D7, U+02DD-02FF, U+0304, U+0308, U+0329, U+1D00-1DBF, U+1E00-1E9F, U+1EF2-1EFF, U+2020, U+20A0-20AB, U+20AD-20C4, U+2113, U+2C60-2C7F, U+A720-A7FF';

// [family, file stem, style, weight]; Arimo is a variable font, one file for both weights.
const FACES = [
  ...['Carlito', 'Tinos', 'Cousine'].flatMap(f => [['normal', '400', ''], ['normal', '700', ''], ['italic', '400', 'i'], ['italic', '700', 'i']]
    .map(([style, weight, i]) => [f, `${f.toLowerCase()}-${weight}${i}`, style, weight])),
  ['Arimo', 'arimo', 'normal', '400 700'], ['Arimo', 'arimo-i', 'italic', '400 700'],
];

// Adds every face to a FontFaceSet (document.fonts on the page, self.fonts in a worker) and loads it all, so
// canvas text never falls back while a file is still downloading. Engines do this at start-up, before a workbook
// is picked. A face that fails to load is skipped: text then uses the next font in the stack.
export function loadFonts(set) {
  const faces = FACES.flatMap(([family, stem, style, weight]) => [['latin', LATIN], ['latin-ext', EXT]].map(([subset, unicodeRange]) =>
    new FontFace(family, `url("${new URL(`vendor/fonts/${stem}-${subset}.woff2`, import.meta.url)}")`, { style, weight, unicodeRange })));
  faces.forEach(f => set.add(f));
  return Promise.all(faces.map(f => f.load().catch(() => null))).then(() => {});
}
