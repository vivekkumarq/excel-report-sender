// The page: takes a workbook (picked, dropped, pasted, opened with the installed app or shared to it), shows the
// outputs in send order as they are made, and sends them: all at once through the WhatsApp helper, or step by
// step through the share screen on phones. The work itself happens in pipeline.js.
import { REPORTS } from './reports.js';
import { processWorkbook, warmUp } from './pipeline.js';
import { zipSync } from './vendor/fflate.js';

const $ = s => document.querySelector(s);
const el = (tag, props = {}, ...kids) => { const e = Object.assign(document.createElement(tag), props); e.append(...kids.filter(k => k != null && k !== false)); return e; };
const icon = name => { const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg'), u = document.createElementNS(s.namespaceURI, 'use'); s.setAttribute('class', 'i'); s.setAttribute('aria-hidden', 'true'); u.setAttribute('href', `#i-${name}`); s.append(u); return s; };
const secs = ms => ms < 1000 ? `${Math.round(ms)} ms` : `${(ms / 1000).toFixed(2)} s`;
const bytes = n => n < 1048576 ? `${Math.max(1, Math.round(n / 1024))} KB` : `${(n / 1048576).toFixed(1)} MB`;
const plural = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;

let toastTimer;
function toast(msg) {
  const t = $('#toast');
  t.textContent = msg; t.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 3200);
}

// ---- getting a workbook ----
const input = $('#file');
input.addEventListener('change', () => input.files[0] && open(input.files[0]));
$('#again').onclick = () => input.click();
const hasFiles = e => [...(e.dataTransfer?.types || [])].includes('Files');
let dragDepth = 0;
addEventListener('dragenter', e => { if (!hasFiles(e)) return; dragDepth++; $('#dropmask').hidden = false; warmUp(); });
addEventListener('dragleave', e => { if (hasFiles(e) && --dragDepth <= 0) { dragDepth = 0; $('#dropmask').hidden = true; } });
addEventListener('dragover', e => { if (hasFiles(e)) e.preventDefault(); });
addEventListener('drop', e => {
  if (!hasFiles(e)) return;
  e.preventDefault(); dragDepth = 0; $('#dropmask').hidden = true;
  const f = e.dataTransfer.files[0];
  if (f) open(f);
});
addEventListener('paste', e => { const f = e.clipboardData?.files?.[0]; if (f) { e.preventDefault(); open(f); } });
// Opened with the installed app (desktop: "Open with", or a double-click on an .xlsx).
window.launchQueue?.setConsumer(async params => { const h = params.files?.[0]; if (h) open(await h.getFile()); });
// Shared to the installed app from another app on a phone (sw.js keeps the file for this page to pick up).
if (new URLSearchParams(location.search).has('shared')) {
  history.replaceState(null, '', location.pathname);
  caches.open('drs-share').then(async c => {
    const res = await c.match('shared');
    if (!res) return;
    await c.delete('shared');
    open(new File([await res.blob()], decodeURIComponent(res.headers.get('x-name') || 'workbook.xlsx'), { type: res.headers.get('content-type') || '' }));
  });
}

// Start the engines (and their fonts) while the user is still choosing a file.
const idle = fn => window.requestIdleCallback ? requestIdleCallback(fn, { timeout: 1500 }) : setTimeout(fn, 200);
if (document.readyState === 'complete') idle(() => warmUp()); else addEventListener('load', () => idle(() => warmUp()), { once: true });
$('#drop').addEventListener('pointerenter', () => warmUp(), { once: true });
navigator.serviceWorker?.register('sw.js').catch(() => {});

let installPrompt = null;
addEventListener('beforeinstallprompt', e => { e.preventDefault(); installPrompt = e; $('#install').hidden = false; });
$('#install').onclick = async () => { if (!installPrompt) return; installPrompt.prompt(); await installPrompt.userChoice; installPrompt = null; $('#install').hidden = true; };
addEventListener('appinstalled', () => { $('#install').hidden = true; toast('Installed. You can now open .xlsx files with it.'); });

// ---- a run ----
let busy = false, run = null, urls = [];
const reportDate = name => {
  const m = /^\s*(\d{1,2})[.\-_](\d{1,2})[.\-_](\d{2,4})/.exec(name);
  if (!m) return { prefix: '', label: '' };
  const y = +m[3] < 100 ? 2000 + +m[3] : +m[3], d = new Date(y, +m[2] - 1, +m[1]);
  const ok = d.getDate() === +m[1] && d.getMonth() === +m[2] - 1;
  return { prefix: m[0].trim(), label: ok ? d.toLocaleDateString('en-IN', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }) : '' };
};

async function open(file) {
  if (busy) return toast('Still working on the last workbook…');
  busy = true;
  urls.forEach(URL.revokeObjectURL); urls = [];
  const { prefix, label } = reportDate(file.name);
  run = { file, prefix, slots: new Map(), groups: [], done: 0, total: 0 };
  guide = null;
  $('#start').hidden = true; $('#work').hidden = false;
  $('#file-name').textContent = file.name;
  $('#file-info').textContent = [label && `Report of ${label}`, bytes(file.size)].filter(Boolean).join(' · ');
  $('#steps').replaceChildren(); $('#stats').hidden = true; $('#warn').hidden = true;
  $('#send-all').disabled = $('#zip').disabled = true;
  setNote('');
  progress(0, 'Reading the workbook…');
  scrollTo({ top: 0, behavior: 'smooth' });
  const r = run;
  try {
    const stats = await processWorkbook(file, REPORTS, { onPlan: p => plan(r, p), onItem: (i, out) => fill(r, i, out) });
    finish(r, stats);
  } catch (e) {
    console.error(e);
    progress(0, '');
    $('#progress').classList.add('done');
    $('#steps').replaceChildren(el('li', { className: 'card failed-run' }, el('p', { className: 'err-line', textContent: `Could not read this file: ${e.message || e}` }), el('p', { className: 'note', textContent: 'Pick the Daily Performance Report workbook (.xlsx).' })));
  } finally {
    busy = false;
    input.value = '';
  }
}

function progress(pct, html) {
  const p = $('#progress');
  p.classList.toggle('done', pct >= 100);
  p.setAttribute('aria-valuenow', String(Math.round(pct)));
  p.querySelector('.fill').style.width = `${pct}%`;
  p.querySelector('.plabel').innerHTML = html;
}

// The plan arrives first: lay out every step, with placeholders that fill in as outputs are made.
function plan(r, { plan, missing }) {
  let pictures = 0;
  for (const e of plan) {
    const last = r.groups.at(-1);
    if (e.as === 'image') e.no = ++pictures;
    if (last && last.kind === e.as && e.as !== 'sep') last.entries.push(e); else r.groups.push({ kind: e.as, entries: [e] });
  }
  r.total = plan.filter(e => e.as !== 'sep').length;
  r.pictures = pictures;
  if (missing.length) { $('#warn').hidden = false; $('#warn').textContent = `Not in this workbook, so left out: ${missing.join(', ')}`; }
  $('#steps').replaceChildren(...r.groups.map((g, k) => stepView(r, g, k + 1)));
  progress(2, `Drawing ${plural(r.total, 'report')} on all cores…`);
}

const TITLE = {
  text: () => ['message', 'Message'],
  image: g => ['image', g.entries.length > 1 ? `Pictures ${g.entries[0].no}–${g.entries.at(-1).no}` : `Picture ${g.entries[0].no}`],
  pdf: () => ['file', 'PDF files'],
  sep: () => ['', 'Separator'],
};

function stepView(r, g, n) {
  g.node = el('li', { className: `step ${g.kind === 'sep' ? 'sep' : ''}` }, el('div', { className: 'rail' }, el('span', { className: 'num', textContent: n })));
  if (g.kind === 'sep') {
    g.ready = true;
    g.node.classList.add('ready');
    g.node.append(el('div', { className: 'card' }, el('span', { className: 'sep-dot', textContent: '.' }), el('span', { textContent: 'A "." message, so WhatsApp shows the pictures before and after it as separate groups' })));
    return g.node;
  }
  const [ic, title] = TITLE[g.kind](g);
  g.sub = el('small');
  g.note = el('p', { className: 'note' });
  g.body = el('div', { className: 'body' });
  g.send = el('button', { className: 'btn primary small', type: 'button', disabled: true }, icon('send'), el('span', { textContent: 'Send' }));
  g.extra = el('button', { className: 'btn ghost small', type: 'button', disabled: true }, icon(g.kind === 'text' ? 'copy' : 'download'), el('span', { textContent: g.kind === 'text' ? 'Copy' : 'Download' }));
  if (g.kind === 'text') g.body.append(g.pre = el('pre', { className: 'text', textContent: 'Reading…' }));
  if (g.kind === 'image') g.body.append(el('div', { className: 'thumbs' }, ...g.entries.map(e => (r.slots.set(e.i, { e, g, view: el('button', { className: 'thumb wait', type: 'button', disabled: true, ariaLabel: `Picture ${e.no}, being drawn` }, el('span', { className: 'tag', textContent: e.no })) }), r.slots.get(e.i).view))));
  if (g.kind === 'pdf') g.body.append(el('div', { className: 'files' }, ...g.entries.map(e => (r.slots.set(e.i, { e, g, view: el('a', { className: 'file wait' }, el('span', { className: 'badge', textContent: 'PDF' }), el('span', {}, el('b', { textContent: `${e.file}.pdf` }), el('small', { textContent: 'Being made…' }))) }), r.slots.get(e.i).view))));
  if (g.kind === 'text') for (const e of g.entries) r.slots.set(e.i, { e, g, view: g.pre });
  g.node.append(el('div', { className: 'card' },
    el('div', { className: 'step-head' }, el('div', { className: 'step-title' }, icon(ic), el('div', {}, el('h2', { textContent: title }), g.sub)), el('div', { className: 'step-acts' }, g.send, g.extra)),
    g.body, g.note));
  updateSub(g);
  return g.node;
}

function fill(r, i, out) {
  if (r !== run) return;
  const s = r.slots.get(i);
  if (!s) return;
  s.out = out;
  r.done++;
  const { e, g, view } = s;
  if (out.error) {
    s.failed = true;
    if (g.kind === 'image') { view.className = 'thumb failed'; view.replaceChildren(icon('alert'), el('span', { textContent: `Picture ${e.no} could not be drawn` })); view.title = out.error; }
    else if (g.kind === 'pdf') { view.className = 'file failed'; view.querySelector('small').textContent = `Could not be made: ${out.error}`; }
    else view.textContent = `Could not read this report: ${out.error}`;
  } else if (g.kind === 'text') {
    s.text = out.text;
    view.textContent = out.text;
  } else {
    const isImage = g.kind === 'image';
    s.file = new File([out.blob], isImage ? `${r.prefix ? `${r.prefix} ` : ''}${String(e.no).padStart(2, '0')}.png` : `${e.file}.pdf`, { type: isImage ? 'image/png' : 'application/pdf' });
    s.url = URL.createObjectURL(s.file); urls.push(s.url);
    if (isImage) {
      view.className = 'thumb'; view.disabled = false; view.ariaLabel = `Open picture ${e.no}`;
      view.prepend(el('img', { src: s.url, alt: `Report picture ${e.no}`, loading: 'lazy', decoding: 'async' }));
      view.onclick = () => openViewer(r, e.no);
    } else {
      view.className = 'file'; view.href = s.url; view.target = '_blank'; view.rel = 'noopener';
      view.querySelector('small').textContent = `${plural(out.pages, 'page')} · ${bytes(s.file.size)}`;
    }
  }
  if (g.entries.every(x => r.slots.get(x.i).out)) groupReady(r, g);
  progress(2 + 96 * r.done / Math.max(1, r.total), `Drawing reports… <b>${r.done}</b> of ${r.total}`);
}

const groupFiles = (r, g) => g.entries.map(x => r.slots.get(x.i)).filter(s => s.file).map(s => s.file);
const groupText = (r, g) => g.kind === 'sep' ? '.' : g.entries.map(x => r.slots.get(x.i).text).filter(Boolean).join('\n\n');
function updateSub(g, r) {
  if (!g.sub) return;
  if (!r || !g.ready) { g.sub.textContent = g.kind === 'text' ? 'Being read…' : `${plural(g.entries.length, g.kind === 'pdf' ? 'file' : 'picture')} · being drawn…`; return; }
  if (g.kind === 'text') { g.sub.textContent = plural(groupText(r, g).split('\n').length, 'line'); return; }
  const files = groupFiles(r, g);
  g.sub.textContent = `${plural(files.length, g.kind === 'pdf' ? 'file' : 'picture')} · ${bytes(files.reduce((s, f) => s + f.size, 0))}`;
}

// ---- sending ----
// With the helper extension, everything goes straight into the logged-in account's own chat ("You") in WhatsApp
// Web. Without it, the system share screen is used (phones), where the user picks the chat.
const helper = () => document.documentElement.dataset.drsHelper;
const SHARE_MAX_FILES = 10, SHARE_MAX_BYTES = 50 * 1024 * 1024; // Chrome refuses bigger shares ("Permission denied")
// As few parts as fit, with the files spread evenly in order (23 pictures -> 8 + 8 + 7).
function shareParts(files) {
  for (let k = Math.max(1, Math.ceil(files.length / SHARE_MAX_FILES)); ; k++) {
    const base = Math.floor(files.length / k), extra = files.length % k, parts = [];
    for (let i = 0, at = 0; i < k; i++) { const n = base + (i < extra ? 1 : 0); parts.push(files.slice(at, at += n)); }
    if (parts.every(p => p.reduce((s, f) => s + f.size, 0) <= SHARE_MAX_BYTES) || k >= files.length) return parts.filter(p => p.length);
  }
}

function groupReady(r, g) {
  g.ready = true;
  g.node.classList.add('ready');
  updateSub(g, r);
  const files = groupFiles(r, g), text = groupText(r, g);
  if (g.kind === 'text' ? !text : !files.length) return; // nothing usable in this step
  const parts = g.kind === 'text' ? [null] : shareParts(files);
  let part = 0;
  const label = () => parts.length > 1 ? `Send (${part + 1} of ${parts.length})` : 'Send';
  const sendLabel = g.send.querySelector('span');
  sendLabel.textContent = label();
  g.send.disabled = g.extra.disabled = false;
  g.send.onclick = async () => {
    if (helper()) return sendAll(r, [g]);
    const data = g.kind === 'text' ? { text } : { files: parts[part] };
    if (!navigator.share || (data.files && !navigator.canShare?.(data))) return note(g, 'err', 'This browser cannot share files. Use Download, then attach them in WhatsApp.');
    try {
      await navigator.share(data);
      note(g, '', '');
      if (++part < parts.length) { sendLabel.textContent = label(); note(g, '', `Part ${part} of ${parts.length} shared. Tap again for the next ${parts[part].length}.`); return; }
      part = 0; sendLabel.textContent = label();
      g.node.classList.add('sent'); note(g, 'ok', 'Shared.');
    } catch (e) {
      if (e.name !== 'AbortError') note(g, 'err', `Sharing failed: ${e.message}`);
    }
  };
  g.extra.onclick = async () => {
    if (g.kind === 'text') return navigator.clipboard.writeText(text).then(() => toast('Message copied'), () => toast('Could not copy; select the text instead.'));
    g.entries.map(x => r.slots.get(x.i)).filter(s => s.file).forEach((s, k) => setTimeout(() => el('a', { href: s.url, download: s.file.name }).click(), k * 150));
  };
}
const note = (g, kind, msg) => { g.note.className = `note ${kind}`; g.note.textContent = msg; };
const setNote = (msg, kind = '') => { const n = $('#send-note'); n.className = `note ${kind}`; n.textContent = msg; };

function finish(r, stats) {
  if (r !== run) return;
  const files = r.groups.flatMap(g => g.kind === 'sep' ? [] : groupFiles(r, g)), failed = [...r.slots.values()].filter(s => s.failed).length;
  progress(100, `<strong>Ready in ${secs(stats.total)}</strong><span>first result after ${secs(stats.first)}</span><span>${plural(stats.threads, 'engine thread')}</span><span>${bytes(files.reduce((s, f) => s + f.size, 0))} in all</span>${failed ? `<span class="err">${plural(failed, 'report')} failed</span>` : ''}`);
  const count = k => r.groups.filter(g => g.kind === k).reduce((s, g) => s + (k === 'sep' ? 1 : groupFiles(r, g).length || (k === 'text' && groupText(r, g) ? 1 : 0)), 0);
  $('#stats').replaceChildren(...[[count('text'), 'message'], [count('image'), 'picture'], [count('pdf'), 'PDF'], [count('sep'), 'dot']]
    .map(([v, what]) => el('div', { className: 'stat' }, el('b', { textContent: v }), el('span', { textContent: v === 1 ? what : `${what}s` }))));
  $('#stats').hidden = false;
  $('#zip').disabled = !files.length;
  refreshSendAll();
}

// Without the helper (phones), the big button walks through the send order instead: every tap opens the share
// screen for the next step, so nothing has to be found or tapped twice. A share needs a tap of its own.
let guide = null;
const shareQueue = r => r.groups.flatMap(g => g.kind === 'text' || g.kind === 'sep' ? (groupText(r, g) ? [{ g, data: { text: groupText(r, g) } }] : [])
  : shareParts(groupFiles(r, g)).map(files => ({ g, data: { files } })));

function refreshSendAll() {
  const btn = $('#send-all'), label = btn.querySelector('span');
  const ready = !!run && run.groups.length && run.groups.every(g => g.ready);
  if (!ready || sending) { btn.disabled = true; return; }
  document.querySelectorAll('.step.next').forEach(n => n.classList.remove('next'));
  if (helper()) {
    label.textContent = 'Send all to my WhatsApp';
    btn.disabled = false;
    if (!$('#send-note').textContent || guide) setNote('Sends every step below, in order, to your own chat ("You") in WhatsApp Web.');
    guide = null;
  } else if (navigator.share) {
    guide ??= { list: shareQueue(run), at: 0 };
    const left = guide.at < guide.list.length;
    label.textContent = left ? `Share step ${guide.at + 1} of ${guide.list.length}` : 'Everything shared';
    btn.disabled = !left;
    guide.list[guide.at]?.g.node.classList.add('next');
    setNote(left ? 'Each tap opens the share screen for the next step, in order: pick your own chat each time. On a computer, the WhatsApp helper (below) sends everything in one click.' : 'All steps shared.', left ? '' : 'ok');
  } else {
    label.textContent = 'Send all to my WhatsApp';
    btn.disabled = true;
    setNote('To send everything to your own WhatsApp chat in one click, add the WhatsApp helper to this browser (see below).');
  }
}
addEventListener('drs-helper', refreshSendAll);

let sending = false;
$('#send-all').onclick = async () => {
  if (!run) return;
  if (helper()) return sendAll(run, run.groups);
  const s = guide?.list[guide.at];
  if (!s) return;
  if (s.data.files && !navigator.canShare?.(s.data)) return setNote('This browser cannot share files. Use Download all, then attach them in WhatsApp.', 'err');
  try {
    await navigator.share(s.data);
    guide.at++;
    if (!guide.list.slice(guide.at).some(x => x.g === s.g)) s.g.node.classList.add('sent');
  } catch (e) {
    if (e.name !== 'AbortError') return setNote(`Sharing failed: ${e.message}`, 'err');
  }
  refreshSendAll();
};
async function sendAll(r, groups) {
  if (sending) return;
  const list = groups.filter(g => g.kind === 'sep' || g.kind === 'text' ? groupText(r, g) : groupFiles(r, g).length);
  const items = list.map(g => g.kind === 'text' || g.kind === 'sep' ? { kind: 'text', text: groupText(r, g) } : { kind: 'files', files: groupFiles(r, g) });
  if (!items.length) return;
  sending = true;
  $('#send-all').disabled = true;
  const bar = $('.send-progress'), barFill = bar.querySelector('.fill');
  bar.hidden = false; barFill.style.width = '0%';
  list.forEach(g => g.node.classList.remove('sending', 'sent'));
  setNote('Preparing…');
  try {
    await viaHelper(items, s => {
      setNote(s.message);
      // the helper reports "Sending 3 of 11…"; mark the steps to match
      const m = /(\d+)\s+of\s+(\d+)/.exec(s.message || '');
      if (s.state === 'sending' && m) {
        const at = +m[1] - 1;
        list.forEach((g, k) => { g.node.classList.toggle('sent', k < at); g.node.classList.toggle('sending', k === at); });
        barFill.style.width = `${100 * at / list.length}%`;
      }
    });
    list.forEach(g => { g.node.classList.remove('sending'); g.node.classList.add('sent'); });
    barFill.style.width = '100%';
    setNote(`Sent ${plural(items.length, 'step')} to your own WhatsApp chat.`, 'ok');
    toast('Sent to your WhatsApp');
  } catch (e) {
    list.forEach(g => g.node.classList.remove('sending'));
    setNote(e.message, 'err');
  } finally {
    sending = false;
    setTimeout(() => { bar.hidden = true; }, 1500);
    refreshSendAll();
  }
}

// The helper's page script (extension/site.js) takes { drs: 'send', id, items } and reports { drs: 'status' }.
function viaHelper(items, onProgress) {
  return new Promise((ok, fail) => {
    const id = crypto.randomUUID();
    const listen = e => {
      if (e.source !== window || e.data?.drs !== 'status' || e.data.id !== id) return;
      onProgress(e.data);
      if (e.data.state === 'done') { removeEventListener('message', listen); ok(); }
      if (e.data.state === 'error') { removeEventListener('message', listen); fail(new Error(e.data.message)); }
    };
    addEventListener('message', listen);
    postMessage({ drs: 'send', id, items }, location.origin);
  });
}

// ---- download everything ----
$('#zip').onclick = async () => {
  const r = run;
  if (!r) return;
  const entries = {}, enc = new TextEncoder();
  for (const g of r.groups) {
    if (g.kind === 'text' && groupText(r, g)) entries['Message.txt'] = [enc.encode(groupText(r, g)), { level: 6 }];
    if (g.kind === 'image' || g.kind === 'pdf') for (const f of groupFiles(r, g)) entries[f.name] = [new Uint8Array(await f.arrayBuffer()), { level: 0 }]; // already compressed
  }
  const url = URL.createObjectURL(new Blob([zipSync(entries)], { type: 'application/zip' }));
  el('a', { href: url, download: `${r.prefix ? `${r.prefix} ` : ''}Daily Reports.zip` }).click();
  setTimeout(() => URL.revokeObjectURL(url), 30e3);
  toast(`Saved ${plural(Object.keys(entries).length, 'file')} in one .zip`);
};

// ---- picture viewer ----
const viewer = $('#viewer'), vImg = $('#v-img');
let shown = [], at = 0;
function openViewer(r, no) {
  shown = [...r.slots.values()].filter(s => s.g.kind === 'image' && s.url).sort((a, b) => a.e.no - b.e.no);
  showAt(shown.findIndex(s => s.e.no === no));
  viewer.showModal();
}
function showAt(k) {
  if (!shown.length) return;
  at = (k + shown.length) % shown.length;
  const s = shown[at];
  vImg.src = s.url; vImg.alt = `Report picture ${s.e.no}`;
  $('#v-count').textContent = `Picture ${s.e.no} of ${run.pictures}`;
  Object.assign($('#v-dl'), { href: s.url, download: s.file.name });
}
$('#v-prev').onclick = () => showAt(at - 1);
$('#v-next').onclick = () => showAt(at + 1);
$('#v-close').onclick = () => viewer.close();
viewer.addEventListener('click', e => { if (e.target === viewer || e.target.classList.contains('v-stage')) viewer.close(); });
viewer.addEventListener('keydown', e => { if (e.key === 'ArrowLeft') showAt(at - 1); if (e.key === 'ArrowRight') showAt(at + 1); });
let swipeX = null;
viewer.addEventListener('pointerdown', e => { swipeX = e.clientX; });
viewer.addEventListener('pointerup', e => { if (swipeX != null && Math.abs(e.clientX - swipeX) > 60) showAt(at + (e.clientX < swipeX ? 1 : -1)); swipeX = null; });
