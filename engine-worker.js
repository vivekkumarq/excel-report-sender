// One engine thread. It opens the workbook from the File it is handed, reading only the sheets its reports need,
// and turns reports into finished files: parse, draw and encode all happen here, so the page stays free.
// Several of these run side by side. Messages:
//   { id, hello }                          -> { ready }   (fonts loaded, able to draw)
//   { id, bookId, file, index?, task }     -> the output of engine.js produce()
// `index` (see book.js readIndex) comes with the first task of each workbook.
import { openBook } from './book.js';
import { produce } from './engine.js';
import { loadFonts } from './fonts.js';

const canDraw = (() => {
  try { return typeof FontFace === 'function' && !!self.fonts && !!new OffscreenCanvas(1, 1).getContext('2d'); } catch { return false; }
})();
const fonts = canDraw ? loadFonts(self.fonts) : null;
let open = { id: null, book: null };

onmessage = async ({ data: { id, hello, bookId, file, index, task } }) => {
  try {
    if (hello) { await fonts; postMessage({ id, result: { ready: canDraw } }); return; }
    if (open.id !== bookId) open = { id: bookId, book: openBook(file, index) }; // a new workbook replaces the old one
    const book = await open.book;
    await book.load(task.refs.map(d => d.sheet));
    postMessage({ id, result: await produce(book, task) });
  } catch (e) {
    postMessage({ id, error: String(e?.message || e) });
  }
};
