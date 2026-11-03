import { REPORTS } from './reports.js';
import { loadBook } from './book.js';
import { renderRange, rangeText } from './render.js';
import { rangePages, stackedPages } from './print.js';
import { jpegPdf } from './pdf.js';
import { WEB_FONTS } from './paint.js';
import { pool } from './pool.js';

const $ = s => document.querySelector(s);
const el = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids.filter(k => k != null)); return e; };
const ms = n => n < 1000 ? `${Math.round(n)} ms` : `${(n / 1000).toFixed(2)} s`;
const PDF_DPI = 300; // resolution of the pictures inside PDFs

// Metric-compatible fonts for devices without Office fonts; loaded while the user picks a file.
const fontsReady = Promise.all(WEB_FONTS.flatMap(f => ['', 'bold ', 'italic ', 'italic bold '].map(v => document.fonts.load(`${v}16px "${f}"`)))).catch(() => {});

const input = $('#file'), drop = $('#drop');
input.addEventListener('change', () => input.files[0] && run(input.files[0]));
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) run(f); });

let busy = false, urls = [];
async function run(file) {
  if (busy) return;
  busy = true;
  urls.forEach(URL.revokeObjectURL); urls = [];
  $('#out').replaceChildren();
  const status = $('#status');
  status.className = 'status'; status.textContent = `Reading ${file.name}…`;
  try {
    const t0 = performance.now();
    const names = [...new Set(REPORTS.flatMap(r => [r.name].flat()))];
    const { book, names: refs, missing, drawings, timing } = await loadBook(file, names);
    await fontsReady;

    // 1. Plan every picture: one per image report, one per PDF page part.
    const jobs = [], plans = REPORTS.map((r, i) => {
      const list = [r.name].flat().filter(n => refs.has(n));
      if (!list.length) return null;
      if (r.as === 'text') return { kind: 'text', text: list.map(n => rangeText(book, refs.get(n).sheet, refs.get(n).range)).join('\n\n') };
      if (r.as === 'image') {
        const d = refs.get(list[0]), key = `i${i}`;
        jobs.push({ key, sheet: d.sheet, rng: d.range, type: 'image/png' });
        return { kind: 'image', key };
      }
      const pages = list.length > 1 ? stackedPages(book, list.map(n => refs.get(n)), r.page)
        : rangePages(book, refs.get(list[0]).sheet, refs.get(list[0]).range, r.page);
      pages.forEach((p, pi) => p.parts.forEach((part, j) => {
        part.key = `p${i}.${pi}.${j}`;
        jobs.push({ key: part.key, sheet: part.sheet ?? refs.get(list[0]).sheet, rng: part.rng, print: part.print, width: part.w * PDF_DPI / 72, type: 'image/jpeg' });
      }));
      return { kind: 'pdf', file: r.file, pages };
    });

    // 2. Draw on this thread while encoder workers turn finished drawings into PNG/JPEG in parallel.
    // Biggest pictures first (longest job first); only a few canvases are alive at once, so phones keep up.
    const t1 = performance.now(), blobs = new Map(), dims = new Map();
    const cores = navigator.hardwareConcurrency || 4, encoders = Math.max(1, Math.min(cores - 1, 8));
    const enc = typeof OffscreenCanvas !== 'undefined' && typeof Worker !== 'undefined' ? pool(new URL('./encode-worker.js', import.meta.url), { max: encoders }) : null;
    const toBlob = (cv, type) => cv.convertToBlob ? cv.convertToBlob({ type, quality: 0.92 })
      : new Promise((ok, no) => cv.toBlob(b => b ? ok(b) : no(new Error('Image encoding failed (the report may be too large).')), type, 0.92));
    const encode = async (cv, type) => {
      if (enc) try {
        const bitmap = cv.transferToImageBitmap ? cv.transferToImageBitmap() : await createImageBitmap(cv); // moves the pixels, no copy
        return await enc.run({ bitmap, type, quality: 0.92 }, { cost: bitmap.width * bitmap.height, transfer: [bitmap] });
      } catch { /* fall back below */ }
      return toBlob(cv, type);
    };
    const area = j => (j.rng[2] - j.rng[0] + 1) * (j.rng[3] - j.rng[1] + 1) * (j.width ? j.width / 1000 : 1);
    const slots = enc ? encoders * 2 : 2, inFlight = new Set();
    let tDraw = 0;
    for (const j of jobs.sort((a, b) => area(b) - area(a))) {
      while (inFlight.size >= slots) await Promise.race(inFlight);
      const s = performance.now(), cv = renderRange(book, j.sheet, j.rng, { drawings: drawings.get(j.sheet), print: j.print, width: j.width });
      tDraw += performance.now() - s;
      dims.set(j.key, { w: cv.width, h: cv.height });
      const job = encode(cv, j.type).then(b => blobs.set(j.key, b)).finally(() => { cv.width = cv.height = 0; inFlight.delete(job); });
      inFlight.add(job);
      await (globalThis.scheduler?.yield?.() ?? new Promise(r => setTimeout(r))); // let worker replies and repaints through
    }
    await Promise.all(inFlight);
    enc?.close();
    const tAll = performance.now() - t1;

    // 3. Files. Image names carry no report name; PDFs keep the names used in the samples.
    const prefix = (/^\s*(\d{1,2}[.\-]\d{1,2}[.\-]\d{2,4})/.exec(file.name)?.[1] ?? '').trim();
    let n = 0;
    const items = await Promise.all(plans.map(async p => {
      if (!p || p.kind === 'text') return p;
      if (p.kind === 'image') return { kind: 'image', file: new File([blobs.get(p.key)], `${prefix ? prefix + ' ' : ''}${String(++n).padStart(2, '0')}.png`, { type: 'image/png' }) };
      const pages = await Promise.all(p.pages.map(async pg => ({
        pw: pg.pw, ph: pg.ph,
        images: await Promise.all(pg.parts.map(async part => ({ jpeg: new Uint8Array(await blobs.get(part.key).arrayBuffer()), ...dims.get(part.key), x: part.x, y: part.y, dw: part.w, dh: part.h }))),
      })));
      return { kind: 'pdf', pages: pages.length, file: new File([jpegPdf(pages)], `${p.file}.pdf`, { type: 'application/pdf' }) };
    }));

    status.innerHTML = '';
    status.append(
      el('strong', { textContent: `Ready in ${ms(performance.now() - t0)}` }),
      ` · read ${ms(timing.read)} · unzip + parse ${ms(timing.parse)} (${timing.sheets} sheets on ${timing.workers} workers, ${timing.skipped} skipped) · draw + encode ${ms(tAll)} (drawing ${ms(tDraw)}, encoding in parallel)`,
    );
    if (missing.length) $('#out').append(el('p', { className: 'warn', textContent: `Not found in this workbook, skipped: ${missing.join(', ')}` }));
    show(groupsOf(items.filter(Boolean)));
  } catch (e) {
    status.className = 'status error';
    status.textContent = `Could not read the file: ${e.message || e}`;
    console.error(e);
  } finally {
    busy = false;
    input.value = '';
  }
}

// Consecutive items of one kind go out together, keeping the configured order.
function groupsOf(items) {
  const groups = [];
  for (const it of items) {
    const g = groups.at(-1);
    if (g && g.kind === it.kind) g.items.push(it); else groups.push({ kind: it.kind, items: [it] });
  }
  return groups.map(g => ({ ...g, files: g.items.map(x => x.file).filter(Boolean), text: g.items.map(x => x.text).filter(Boolean).join('\n\n') }));
}

// ---- sending ----
// With the helper extension, everything goes straight into the logged-in account's own chat ("You") in
// WhatsApp Web. Without it, the system share sheet is used (phones), where the user picks the chat.
const helper = () => document.documentElement.dataset.drsHelper;
let sending = null;
function sendViaHelper(groups, onProgress) {
  return new Promise((ok, fail) => {
    const id = crypto.randomUUID();
    const listen = e => {
      if (e.source !== window || e.data?.drs !== 'status' || e.data.id !== id) return;
      onProgress(e.data);
      if (e.data.state === 'done') { removeEventListener('message', listen); ok(); }
      if (e.data.state === 'error') { removeEventListener('message', listen); fail(new Error(e.data.message)); }
    };
    addEventListener('message', listen);
    const items = groups.map(g => g.kind === 'text' ? { kind: 'text', text: g.text } : { kind: 'files', files: g.files });
    postMessage({ drs: 'send', id, items }, location.origin);
  });
}
const shareable = g => g.kind === 'text' ? !!navigator.share : !!navigator.canShare?.({ files: g.files.slice(0, 1) });
// Chrome refuses a share of more than 10 files or 50 MB ("Permission denied"). Use as few parts as fit, with
// the files spread evenly in order (23 images -> 8 + 8 + 7).
const SHARE_MAX_FILES = 10, SHARE_MAX_BYTES = 50 * 1024 * 1024;
function shareParts(files) {
  for (let k = Math.max(1, Math.ceil(files.length / SHARE_MAX_FILES)); ; k++) {
    const base = Math.floor(files.length / k), extra = files.length % k, parts = [];
    for (let i = 0, at = 0; i < k; i++) { const n = base + (i < extra ? 1 : 0); parts.push(files.slice(at, at += n)); }
    if (parts.every(p => p.reduce((s, f) => s + f.size, 0) <= SHARE_MAX_BYTES) || k >= files.length) return parts.filter(p => p.length);
  }
}

const KIND = { text: 'Text message', image: 'Images', pdf: 'PDFs' };
function show(groups) {
  const out = $('#out');
  const all = el('section', { className: 'sendall' });
  const allBtn = el('button', { className: 'primary big', textContent: 'Send all to my WhatsApp' });
  const allNote = el('p', { className: 'note' });
  const refreshAll = () => {
    const ok = !!helper();
    allBtn.disabled = !ok || !!sending;
    allNote.textContent = ok ? 'Sends the text, images and PDFs, in order, to your own chat ("You") in WhatsApp Web.'
      : 'To send everything to your own WhatsApp chat in one click, add the WhatsApp helper to this browser (see below). On a phone, use the buttons of each step.';
  };
  allBtn.onclick = () => send(groups, allBtn, allNote);
  all.append(allBtn, allNote);
  out.append(all);
  refreshAll();
  addEventListener('message', e => { if (e.source === window && e.data?.drs === 'ready') refreshAll(); });

  groups.forEach((g, i) => {
    const body = g.kind === 'text' ? el('pre', { className: 'text', textContent: g.text })
      : g.kind === 'image' ? el('div', { className: 'thumbs' }, ...g.items.map((x, k) => {
        const url = URL.createObjectURL(x.file); urls.push(url);
        return el('a', { href: url, target: '_blank', rel: 'noopener' }, el('img', { src: url, alt: `Report image ${k + 1}`, loading: 'lazy' }));
      }))
      : el('ul', { className: 'files' }, ...g.items.map(x => {
        const url = URL.createObjectURL(x.file); urls.push(url);
        return el('li', {}, el('a', { href: url, target: '_blank', rel: 'noopener', textContent: x.file.name }), el('span', { textContent: ` ${x.pages} page${x.pages > 1 ? 's' : ''} · ${Math.round(x.file.size / 1024)} KB` }));
      }));

    // The share screen takes at most 10 files per share, so bigger steps go out in parts, one tap each.
    const parts = g.kind === 'text' ? [null] : shareParts(g.files);
    let part = 0;
    const label = () => parts.length > 1 ? `Send to WhatsApp (${part + 1} of ${parts.length})` : 'Send to WhatsApp';
    const sendBtn = el('button', { className: 'primary', textContent: label() });
    sendBtn.onclick = async () => {
      if (helper()) return send([g], sendBtn, note);
      if (!shareable(g)) return alertIn(card, 'This browser cannot share files. Use Download, then attach them in WhatsApp.');
      try {
        await navigator.share(g.kind === 'text' ? { text: g.text } : { files: parts[part] });
        card.querySelector('.err')?.remove();
        if (++part < parts.length) { sendBtn.textContent = label(); note.textContent = `Part ${part} of ${parts.length} shared. Tap again for the next ${parts[part].length}.`; return; }
        card.classList.add('done'); sendBtn.textContent = 'Shared ✓'; note.textContent = '';
      } catch (e) {
        if (e.name !== 'AbortError') alertIn(card, `Sharing failed: ${e.message}`);
      }
    };
    const extra = g.kind === 'text'
      ? el('button', { textContent: 'Copy', onclick: async e => { await navigator.clipboard.writeText(g.text); e.target.textContent = 'Copied ✓'; } })
      : el('button', { textContent: 'Download', onclick: () => g.files.forEach((f, k) => setTimeout(() => el('a', { href: URL.createObjectURL(f), download: f.name }).click(), k * 150)) });
    const note = el('p', { className: 'note' });
    const count = g.kind === 'text' ? '' : ` · ${g.items.length}`;
    const card = el('section', { className: 'step' },
      el('header', {}, el('span', { className: 'n', textContent: i + 1 }), el('h2', { textContent: KIND[g.kind] + count })),
      body, el('div', { className: 'actions' }, sendBtn, extra), note);
    out.append(card);
  });
}

async function send(groups, btn, note) {
  if (sending) return;
  const label = btn.textContent;
  sending = true;
  btn.disabled = true;
  note.className = 'note';
  try {
    await sendViaHelper(groups, s => { note.textContent = s.message; });
    btn.textContent = 'Sent ✓';
    note.textContent = 'Sent to your own WhatsApp chat.';
  } catch (e) {
    btn.textContent = label;
    note.className = 'note err';
    note.textContent = e.message;
  } finally {
    sending = null;
    btn.disabled = false;
  }
}

function alertIn(card, msg) {
  card.querySelector('.err')?.remove();
  card.append(el('p', { className: 'err', textContent: msg }));
}
