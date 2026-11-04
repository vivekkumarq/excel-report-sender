// Self-check: node test/check.mjs [path/to/workbook.xlsx]
// Without a path it checks the formula engine and the PDF writer; with one it also loads every configured report.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { compile, cmp } from '../cf.js';
import { jpegPdf } from '../pdf.js';
import { parseRange } from '../xlsx.js';

const cells = { 'A1': 5, 'B1': 'Achieved ', 'C1': null, 'A2': 10, 'B2': 'x' };
const get = (sh, r, c) => cells[String.fromCharCode(64 + c) + r] ?? null;
const ev = (f, dr = 0, dc = 0) => compile(f)({ dr, dc, get });
assert.equal(ev('A1+2*3'), 11);
assert.equal(ev('-2^2'), 4);
assert.equal(ev('A1<10%'), false);
assert.equal(ev('TRIM(B1)="achieved"'), true);
assert.equal(ev('AND(A1>=5, A1<6)'), true);
assert.equal(ev('SUM(A1:A2)'), 15);
assert.equal(ev('A1>$A$2', 1), false); // relative A1 shifts to A2 (10 > 10)
assert.equal(ev('C1=""'), true);
assert.ok(cmp('a', 1) > 0 && cmp(null, 0) === 0);
assert.deepEqual(parseRange('B3:A1'), [1, 1, 3, 2]);

const img = { jpeg: new Uint8Array([0xff, 0xd8, 0xff, 0xd9]), w: 1, h: 1, x: 5, y: 5, dw: 10, dh: 10 };
const pdf = new Uint8Array(await jpegPdf([{ pw: 612, ph: 792, images: [img, img] }, { pw: 792, ph: 612, images: [img] }]).arrayBuffer());
assert.equal(new TextDecoder('latin1').decode(pdf).match(/\/Type \/Page\b/g).length, 2);
const txt = new TextDecoder('latin1').decode(pdf), xref = +/startxref\n(\d+)/.exec(txt)[1];
assert.ok(txt.startsWith('%PDF-1.4') && txt.slice(xref).startsWith('xref'));
for (const [i, o] of [...txt.slice(xref).matchAll(/(\d{10}) 00000 n/g)].entries()) assert.ok(txt.slice(+o[1]).startsWith(`${i + 1} 0 obj`));
console.log('formula engine and PDF writer: ok');

const { REPORTS } = await import('../reports.js');
const before = REPORTS.flatMap((r, i) => r.as === 'sep' ? [REPORTS[i - 1].name] : []);
assert.deepEqual(before, ['Alok_Kr_Singh_Sir_01', '_ccl_rake', 'AREAWISE_STOCK', 'CMD_Report_06',
  '_area_wise_mine_wise', '_dep_obr_ccl', 'Rainfall_Report', 'Yearly_Cap_Utlzn_Report', 'Top_20']);
assert.equal(REPORTS.filter(r => r.as === 'image').length, 23);
console.log('separators: ok');

const path = process.argv[2];
if (path) {
  const { loadBook } = await import('../book.js');
  const { rangeText } = await import('../render.js');
  const names = [...new Set(REPORTS.flatMap(r => [r.name ?? []].flat()))];
  const t = performance.now();
  const { book, names: refs, missing, timing } = await loadBook(new Blob([await readFile(path)]), names);
  console.log(`loaded ${refs.size}/${names.length} reports in ${Math.round(performance.now() - t)} ms`, timing);
  assert.deepEqual(missing, [], `missing names: ${missing}`);
  book.mdw = 8; // Calibri 12 (no canvas in Node to measure it)
  const sms = REPORTS.find(r => r.as === 'text');
  const text = rangeText(book, refs.get(sms.name).sheet, refs.get(sms.name).range);
  assert.match(text, /^\*Date:\*/);
  console.log(`${sms.name}: ${text.split('\n').length} lines\n${text.split('\n').slice(0, 4).join('\n')}\n…`);
}
