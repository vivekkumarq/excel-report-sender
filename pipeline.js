// Runs every report through a set of engine workers (engine-worker.js) in parallel.
//  - The page reads the workbook's index once (sheets, names, shared strings, styles; book.js readIndex) and
//    hands it to every worker with its first report, then hands out reports, biggest first.
//  - Each worker opens the File itself and parses only the sheets of the reports it gets. A worker prefers
//    reports on sheets it has already parsed, then sheets nobody has parsed yet, then the biggest report left,
//    so the cores stay busy and few sheets are parsed twice.
//  - Workers start, and load their fonts, before a file is picked (warmUp), so a run starts at once.
//  - If a worker stops (out of memory, say), its report goes to the others. If workers cannot draw in this
//    browser at all, the same engine runs on the page instead.
import { openZip, readIndex, findNames, openBook } from './book.js';

// Parallel drawing is limited by memory bandwidth and the chip's shared power budget more than by core count:
// on a 4-core / 8-thread laptop, 6 threads beat 8. Phones and small machines get fewer, by memory.
const cores = navigator.hardwareConcurrency || 4;
const THREADS = Math.max(1, Math.min(8, Math.round(cores * 0.75), Math.max(2, Math.floor(navigator.deviceMemory || 8))));
const TASK_TIMEOUT = 120e3; // a report takes well under a second; a thread silent this long has died

const stopped = () => Object.assign(new Error('An engine thread stopped.'), { stopped: true });
function spawn() {
  const w = new Worker(new URL('./engine-worker.js', import.meta.url), { type: 'module' }), waiting = new Map();
  let seq = 0, dead = false;
  const kill = () => { dead = true; w.terminate(); for (const p of waiting.values()) { clearTimeout(p.timer); p.fail(stopped()); } waiting.clear(); };
  w.onmessage = ({ data }) => {
    const p = waiting.get(data.id);
    if (!p) return;
    waiting.delete(data.id); clearTimeout(p.timer);
    data.error ? p.fail(new Error(data.error)) : p.ok(data.result);
  };
  w.onerror = e => { e.preventDefault(); kill(); };
  return {
    kill,
    call: msg => dead ? Promise.reject(stopped()) : new Promise((ok, fail) => {
      const id = seq++;
      waiting.set(id, { ok, fail, timer: setTimeout(kill, TASK_TIMEOUT) });
      w.postMessage({ ...msg, id });
    }),
  };
}

let engines = null;
// Starts the workers and loads their fonts; resolves to the ones able to draw ([] means: draw on the page).
export function warmUp() {
  return engines ??= (async () => {
    if (typeof Worker === 'undefined' || typeof OffscreenCanvas === 'undefined') return [];
    const all = Array.from({ length: THREADS }, spawn);
    const ready = await Promise.all(all.map(e => e.call({ hello: true }).then(r => r.ready, () => false)));
    all.forEach((e, i) => ready[i] || e.kill());
    return all.filter((_, i) => ready[i]);
  })();
}

const cells = d => (d.range[2] - d.range[0] + 1) * (d.range[3] - d.range[1] + 1);
const WEIGHT = { text: 0.1, image: 1, pdf: 4 }; // PDF pages are drawn at 300 dpi

// reports: the REPORTS list. onPlan({ plan, missing }) comes first, with one entry per output in send order;
// then onItem(i, output) once per report as it finishes (in any order), output being { text } | { blob, pages? }
// | { error }. Resolves with timing once everything is done.
export async function processWorkbook(file, reports, { onPlan, onItem }) {
  const t0 = performance.now(), index = await readIndex(await openZip(file));
  const { refs, missing } = findNames(index.wb, [...new Set(reports.flatMap(r => [r.name ?? []].flat()))]);
  const tasks = [], plan = reports.map((r, i) => {
    if (r.as === 'sep') return { i, as: 'sep' };
    const list = [r.name].flat().filter(n => refs.has(n)).map(n => refs.get(n));
    if (!list.length) return null;
    tasks.push({ i, cost: list.reduce((s, d) => s + cells(d), 0) * WEIGHT[r.as], task: { as: r.as, refs: list, page: r.page } });
    return { i, as: r.as, file: r.file };
  }).filter(Boolean);
  onPlan({ plan, missing });

  let first = 0;
  const report = (i, out) => { first ||= performance.now() - t0; onItem(i, out); };
  const workers = await warmUp();
  const left = workers.length ? await onWorkers(workers, file, index, tasks, report) : tasks;
  if (left.length) await onPage(file, index, left, report);
  return { total: performance.now() - t0, first, threads: workers.length || 1 };
}

let books = 0;
async function onWorkers(workers, file, index, tasks, report) {
  const bookId = ++books, queue = [...tasks].sort((a, b) => b.cost - a.cost), holds = workers.map(() => new Set());
  const sheetsOf = t => t.task.refs.map(d => d.sheet);
  const pick = k => {
    let j = queue.findIndex(t => sheetsOf(t).every(s => holds[k].has(s)));
    if (j < 0) j = queue.findIndex(t => sheetsOf(t).every(s => !holds.some(h => h.has(s))));
    return queue.splice(Math.max(0, j), 1)[0];
  };
  await Promise.all(workers.map(async (w, k) => {
    let sent = false;
    while (queue.length) {
      const t = pick(k);
      sheetsOf(t).forEach(s => holds[k].add(s));
      try {
        report(t.i, await w.call({ bookId, file, task: t.task, index: sent ? undefined : index }));
        sent = true;
      } catch (e) {
        if (e.stopped) { queue.push(t); return; } // this thread is gone; the others take over its report
        report(t.i, { error: e.message });
        sent = true;
      }
    }
  }));
  return queue; // anything left only if every thread stopped
}

// The same engine on the page, one report at a time, letting the page repaint in between.
let pageFonts = null;
const breathe = () => globalThis.scheduler?.yield?.() ?? new Promise(r => { const c = new MessageChannel(); c.port1.onmessage = r; c.port2.postMessage(0); });
async function onPage(file, index, tasks, report) {
  const [{ produce }, { loadFonts }] = await Promise.all([import('./engine.js'), import('./fonts.js')]);
  await (pageFonts ??= loadFonts(document.fonts));
  const book = await openBook(file, index);
  for (const t of [...tasks].sort((a, b) => a.i - b.i)) {
    try {
      await book.load(t.task.refs.map(d => d.sheet));
      report(t.i, await produce(book, t.task));
    } catch (e) { report(t.i, { error: e.message || String(e) }); }
    await breathe();
  }
}
