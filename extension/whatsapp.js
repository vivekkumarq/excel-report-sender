// Runs on WhatsApp Web. Opens the logged-in account's own chat ("You", or however the language names it) and hands
// each item to whatsapp-page.js, one after another, in order.
// Fast route: whatsapp-page.js opens the own chat directly (no reload). Fallback, only if WhatsApp Web's internals
// change: reload at /send?phone=<own number>, which is slower because WhatsApp looks the number up first.
const sleep = ms => new Promise(r => setTimeout(r, ms));
const waitFor = async (fn, ms) => { for (const end = Date.now() + ms; ;) { const v = fn(); if (v) return v; if (Date.now() > end) return null; await sleep(200); } };
const status = (job, state, message) => chrome.runtime.sendMessage({ drs: 'status', id: job.id, state, message }).catch(() => {});

function ownNumber() {
  for (const k of ['last-wid-md', 'last-wid']) {
    try { const v = JSON.parse(localStorage.getItem(k) || 'null'); if (v) return String(v).split(/[:@]/)[0]; } catch { /* next key */ }
  }
  return null;
}

// Ask the page-world script to do something and wait for its answer.
function page(drs, item) {
  const id = crypto.randomUUID();
  return new Promise((ok, fail) => {
    const listen = e => {
      if (e.source !== window || e.data?.drs !== 'wa-done' || e.data.id !== id) return;
      removeEventListener('message', listen);
      e.data.error ? fail(new Error(e.data.error)) : ok(e.data);
    };
    addEventListener('message', listen);
    window.postMessage({ drs, id, item }, location.origin);
  });
}
const asFiles = files => Promise.all(files.map(async f => new File([await (await fetch(`data:${f.type};base64,${f.data}`)).blob()], f.name, { type: f.type })));

let running = false;
async function start() {
  if (running) return;
  const { job } = await chrome.storage.local.get('job');
  if (!job || Date.now() - job.created > 10 * 60e3) return; // nothing to do, or a stale job
  running = true;
  try {
    if (!await waitFor(() => document.querySelector('#side'), 90e3)) throw new Error('WhatsApp Web is not logged in in this browser. Log in, then press Send again.');
    status(job, 'opening', 'Opening your own chat…');
    if (job.stage !== 'opening') {
      const r = await page('wa-open');
      if (r.fallback) {
        const me = ownNumber();
        if (!me) throw new Error('Could not find your own number in WhatsApp Web.');
        await chrome.storage.local.set({ job: { ...job, stage: 'opening' } });
        location.href = `https://web.whatsapp.com/send?phone=${me}`; // a fresh load here opens only the own chat
        return;
      }
    } else if (!await waitFor(() => document.querySelector('#main footer [contenteditable="true"]'), 45e3)) {
      throw new Error('Could not open your own chat in WhatsApp Web; nothing was sent.');
    }
    await chrome.storage.local.remove('job'); // from here on, this job is never sent twice
    for (const [i, item] of job.items.entries()) {
      status(job, 'sending', `Sending ${i + 1} of ${job.items.length}…`);
      await page('wa-item', item.kind === 'files' ? { kind: 'files', files: await asFiles(item.files) } : item);
    }
    status(job, 'done', 'Sent');
  } catch (e) {
    await chrome.storage.local.remove('job');
    status(job, 'error', e.message || String(e));
  } finally {
    running = false;
  }
}

chrome.runtime.onMessage.addListener(m => { if (m?.drs === 'start') start(); });
start();
