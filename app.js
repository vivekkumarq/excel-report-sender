import { REPORTS } from './reports.js';
import { loadBook } from './book.js';
import { renderRange, rangeText } from './render.js';
import { jpegPdf } from './pdf.js';
import { WEB_FONTS } from './paint.js';
import { pool } from './pool.js';

const $ = s => document.querySelector(s);
const el = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids.filter(k => k != null)); return e; };
const ms = n => n < 1000 ? `${Math.round(n)} ms` : `${(n / 1000).toFixed(2)} s`;

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

    // Draw on this thread while encoder workers turn finished drawings into PNG/JPEG in parallel.
    // Biggest reports go first (longest job first); only a few canvases are alive at once, so phones keep up.
    const t1 = performance.now(), formats = new Map(), sizes = new Map(), blobs = new Map();
    for (const r of REPORTS) if (r.as !== 'text') for (const n of [r.name].flat()) if (refs.has(n)) (formats.get(n) || formats.set(n, new Set()).get(n)).add(r.as === 'pdf' ? 'image/jpeg' : 'image/png');
    const area = n => { const [a, b, c, d] = refs.get(n).range; return (c - a + 1) * (d - b + 1); };
    // One core stays free for drawing; measured fastest on an 8-core PC (7 encoders, ~2 jobs queued per encoder).
    const cores = navigator.hardwareConcurrency || 4, encoders = Math.max(1, Math.min(cores - 1, 8));
    const enc = typeof OffscreenCanvas !== 'undefined' && typeof Worker !== 'undefined' ? pool(new URL('./encode-worker.js', import.meta.url), { max: encoders }) : null;
    const toBlob = (cv, type) => cv.convertToBlob ? cv.convertToBlob({ type, quality: 0.92 })
      : new Promise((ok, no) => cv.toBlob(b => b ? ok(b) : no(new Error('Image encoding failed (the report may be too large).')), type, 0.92));
    // The last format of a report takes the canvas pixels outright (no copy); earlier formats get a copy.
    const encode = async (cv, type, last) => {
      const cost = cv.width * cv.height;
      if (enc) try {
        const bitmap = last && cv.transferToImageBitmap ? cv.transferToImageBitmap() : await createImageBitmap(cv);
        return await enc.run({ bitmap, type, quality: 0.92 }, { cost, transfer: [bitmap] });
      } catch { /* fall back below */ }
      return toBlob(cv, type);
    };
    const slots = enc ? encoders * 2 : 2, inFlight = new Set();
    let tDraw = 0;
    for (const n of [...formats.keys()].sort((a, b) => area(b) - area(a))) {
      while (inFlight.size >= slots) await Promise.race(inFlight);
      const d = refs.get(n), s = performance.now(), cv = renderRange(book, d.sheet, d.range, { drawings: drawings.get(d.sheet) });
      tDraw += performance.now() - s;
      sizes.set(n, { w: cv.width, h: cv.height, pw: cv.logical.w * 0.75, ph: cv.logical.h * 0.75 });
      const types = [...formats.get(n)], encodes = [];
      for (const [i, type] of types.entries()) encodes.push(encode(cv, type, i === types.length - 1).then(b => blobs.set(`${n}|${type}`, b)));
      const job = Promise.all(encodes)
        .finally(() => { cv.width = cv.height = 0; inFlight.delete(job); }); // free canvas memory as soon as it is encoded
      inFlight.add(job);
      await (globalThis.scheduler?.yield?.() ?? new Promise(r => setTimeout(r))); // let worker replies and repaints through
    }
    await Promise.all(inFlight);
    enc?.close();
    const tAll = performance.now() - t1;

    const prefix = (/^\s*(\d{1,2}[.\-]\d{1,2}[.\-]\d{2,4})/.exec(file.name)?.[1] ?? '').trim();
    const fileName = (base, ext) => `${prefix ? prefix + ' ' : ''}${base}.${ext}`;
    const items = await Promise.all(REPORTS.map(async r => {
      const list = [r.name].flat().filter(n => refs.has(n));
      if (!list.length) return null;
      if (r.as === 'text') return { kind: 'text', label: r.name, text: list.map(n => rangeText(book, refs.get(n).sheet, refs.get(n).range)).join('\n\n') };
      if (r.as === 'image') return { kind: 'image', label: r.name, file: new File([blobs.get(`${list[0]}|image/png`)], fileName(r.name, 'png'), { type: 'image/png' }) };
      const pages = await Promise.all(list.map(async n => ({ jpeg: new Uint8Array(await blobs.get(`${n}|image/jpeg`).arrayBuffer()), ...sizes.get(n) })));
      return { kind: 'pdf', label: r.file || list.join(' + '), pages: list.length, file: new File([jpegPdf(pages)], fileName(r.file || r.name, 'pdf'), { type: 'application/pdf' }) };
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

// Consecutive items of one kind go out in a single WhatsApp share, keeping the configured order.
function groupsOf(items) {
  const groups = [];
  for (const it of items) {
    const g = groups.at(-1);
    if (g && g.kind === it.kind) g.items.push(it); else groups.push({ kind: it.kind, items: [it] });
  }
  return groups;
}

const canShareFiles = files => !!navigator.canShare?.({ files });
const KIND = { text: 'Text message', image: 'Images', pdf: 'PDFs' };

function show(groups) {
  const out = $('#out');
  groups.forEach((g, i) => {
    const files = g.items.map(x => x.file).filter(Boolean), text = g.items.map(x => x.text).filter(Boolean).join('\n\n');
    const body = g.kind === 'text' ? el('pre', { className: 'text', textContent: text })
      : g.kind === 'image' ? el('div', { className: 'thumbs' }, ...g.items.map(x => {
        const url = URL.createObjectURL(x.file); urls.push(url);
        return el('a', { href: url, target: '_blank', title: x.label }, el('img', { src: url, alt: x.label, loading: 'lazy' }), el('span', { textContent: x.label }));
      }))
      : el('ul', { className: 'files' }, ...g.items.map(x => {
        const url = URL.createObjectURL(x.file); urls.push(url);
        return el('li', {}, el('a', { href: url, target: '_blank', textContent: x.file.name }), el('span', { textContent: ` ${x.pages} page${x.pages > 1 ? 's' : ''} · ${Math.round(x.file.size / 1024)} KB` }));
      }));

    const send = el('button', { className: 'primary', textContent: 'Send to WhatsApp' });
    const shareable = g.kind === 'text' ? !!navigator.share : canShareFiles(files);
    send.disabled = !shareable;
    if (!shareable) send.title = 'This browser cannot share files. Use Download, then attach them in WhatsApp.';
    send.onclick = async () => {
      try {
        await navigator.share(g.kind === 'text' ? { text } : { files });
        card.classList.add('done'); send.textContent = 'Shared ✓';
        card.nextElementSibling?.querySelector('button.primary')?.focus();
      } catch (e) {
        if (e.name !== 'AbortError') alertIn(card, `Sharing failed: ${e.message}`);
      }
    };
    const extra = g.kind === 'text'
      ? el('button', { textContent: 'Copy', onclick: async e => { await navigator.clipboard.writeText(text); e.target.textContent = 'Copied ✓'; } })
      : el('button', { textContent: 'Download', onclick: () => files.forEach((f, k) => setTimeout(() => el('a', { href: URL.createObjectURL(f), download: f.name }).click(), k * 150)) });

    const count = g.kind === 'text' ? '' : ` · ${g.items.length}`;
    const card = el('section', { className: 'step' },
      el('header', {}, el('span', { className: 'n', textContent: i + 1 }), el('h2', { textContent: KIND[g.kind] + count })),
      body, el('div', { className: 'actions' }, send, extra));
    out.append(card);
  });
  out.querySelector('button.primary:not(:disabled)')?.focus();
}

function alertIn(card, msg) {
  card.querySelector('.err')?.remove();
  card.append(el('p', { className: 'err', textContent: msg }));
}
