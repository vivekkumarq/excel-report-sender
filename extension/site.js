// Runs on the Daily Report Sender page. The page posts { drs: 'send', id, items }; the files are stored for the
// WhatsApp Web tab, which the background script opens or brings forward. Progress flows back as { drs: 'status' }.
document.documentElement.dataset.drsHelper = chrome.runtime.getManifest().version;

const toBase64 = file => new Promise((ok, fail) => {
  const r = new FileReader();
  r.onload = () => ok(String(r.result).slice(String(r.result).indexOf(',') + 1));
  r.onerror = () => fail(r.error);
  r.readAsDataURL(file);
});

window.addEventListener('message', async e => {
  if (e.source !== window || e.data?.drs !== 'send') return;
  const { id, items } = e.data;
  const status = (state, message) => window.postMessage({ drs: 'status', id, state, message }, location.origin);
  try {
    status('preparing', 'Preparing the files…');
    const job = {
      id, created: Date.now(), stage: 'new',
      items: await Promise.all(items.map(async it => it.kind === 'text' ? { kind: 'text', text: it.text }
        : { kind: 'files', files: await Promise.all(it.files.map(async f => ({ name: f.name, type: f.type, data: await toBase64(f) }))) })),
    };
    await chrome.storage.local.set({ job });
    const r = await chrome.runtime.sendMessage({ drs: 'open' });
    if (r?.error) throw new Error(r.error);
    status('opening', 'Opening WhatsApp Web…');
  } catch (err) {
    status('error', err.message || String(err));
  }
});

chrome.runtime.onMessage.addListener(m => {
  if (m?.drs === 'status') window.postMessage(m, location.origin);
});
