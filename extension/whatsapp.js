// Runs on WhatsApp Web. Opens the logged-in account's own chat ("You", or however the language names it) by
// going to /send?phone=<own number>, then hands each item to whatsapp-page.js, one after another, in order.
// The own number comes from WhatsApp Web's storage, so it works for any account and any chat label.
const sleep = ms => new Promise(r => setTimeout(r, ms));
const waitFor = async (fn, ms) => { for (const end = Date.now() + ms; ;) { const v = fn(); if (v) return v; if (Date.now() > end) return null; await sleep(250); } };
const status = (job, state, message) => chrome.runtime.sendMessage({ drs: 'status', id: job.id, state, message }).catch(() => {});

function ownNumber() {
  for (const k of ['last-wid-md', 'last-wid']) {
    try { const v = JSON.parse(localStorage.getItem(k) || 'null'); if (v) return String(v).split(/[:@]/)[0]; } catch { /* next key */ }
  }
  return null;
}

// Hand one item to the page-world script and wait until WhatsApp has taken it.
async function toPage(item) {
  const id = crypto.randomUUID();
  const files = item.kind === 'files' ? await Promise.all(item.files.map(async f => new File([await (await fetch(`data:${f.type};base64,${f.data}`)).blob()], f.name, { type: f.type }))) : null;
  return new Promise((ok, fail) => {
    const listen = e => {
      if (e.source !== window || e.data?.drs !== 'wa-done' || e.data.id !== id) return;
      removeEventListener('message', listen);
      e.data.error ? fail(new Error(e.data.error)) : ok();
    };
    addEventListener('message', listen);
    window.postMessage({ drs: 'wa-item', id, item: { kind: item.kind, text: item.text, files } }, location.origin);
  });
}

let running = false;
async function start() {
  if (running) return;
  const { job } = await chrome.storage.local.get('job');
  if (!job || Date.now() - job.created > 10 * 60e3) return; // nothing to do, or a stale job
  running = true;
  try {
    if (!await waitFor(() => document.querySelector('#side'), 90e3)) throw new Error('WhatsApp Web is not logged in in this browser. Log in, then press Send again.');
    const me = ownNumber();
    if (!me) throw new Error('Could not find your own number in WhatsApp Web.');
    if (job.stage !== 'opening') {
      // A fresh load at /send?phone=<own number> opens exactly the own chat and nothing else.
      await chrome.storage.local.set({ job: { ...job, stage: 'opening' } });
      status(job, 'opening', 'Opening your own chat…');
      location.href = `https://web.whatsapp.com/send?phone=${me}`;
      return;
    }
    await chrome.storage.local.remove('job'); // whatever happens next, this job is never sent twice
    if (!await waitFor(() => document.querySelector('#main footer [contenteditable="true"]'), 45e3)) throw new Error('Could not open your own chat in WhatsApp Web; nothing was sent.');
    for (const [i, item] of job.items.entries()) {
      status(job, 'sending', `Sending ${i + 1} of ${job.items.length}…`);
      await toPage(item);
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
