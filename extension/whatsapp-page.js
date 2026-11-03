// Runs in WhatsApp Web's own page context. Sends one item into the open chat: text is pasted into the message
// box; images or PDFs are pasted as files, which opens WhatsApp's preview, and its send button is pressed.
// Buttons are found by WhatsApp's icon names, which do not change with the interface language.
const sleep = ms => new Promise(r => setTimeout(r, ms));
const waitFor = async (fn, ms) => { for (const end = Date.now() + ms; ;) { const v = fn(); if (v) return v; if (Date.now() > end) return null; await sleep(200); } };
const composer = () => document.querySelector('#main footer [contenteditable="true"]');
const sendButtons = () => [...document.querySelectorAll('svg title')].filter(t => /(^|-)send(-filled)?$/.test(t.textContent)).map(t => t.closest('[role="button"], button')).filter(Boolean);
const previewSend = () => sendButtons().find(b => !b.closest('#main footer'));

function paste(dt) {
  const box = composer();
  if (!box) throw new Error('The chat is not open.');
  box.focus();
  box.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: dt }));
}

async function sendText(text) {
  const dt = new DataTransfer();
  dt.setData('text/plain', text);
  paste(dt);
  if (!await waitFor(() => composer()?.innerText.trim(), 5000)) throw new Error('Could not put the text into the message box.');
  const btn = await waitFor(() => sendButtons().find(b => b.closest('#main footer')), 5000);
  if (!btn) throw new Error('WhatsApp send button not found.');
  btn.click();
  await waitFor(() => !composer()?.innerText.trim(), 10000);
  await sleep(600);
}

async function sendFiles(files) {
  for (let k = 0; k < files.length; k += 30) { // WhatsApp takes up to 30 files per send
    const dt = new DataTransfer();
    files.slice(k, k + 30).forEach(f => dt.items.add(f));
    paste(dt);
    const btn = await waitFor(previewSend, 30000);
    if (!btn) throw new Error('WhatsApp did not open the file preview.');
    await sleep(1000); // let the preview finish loading every file
    btn.click();
    if (!await waitFor(() => !previewSend(), 60000)) throw new Error('WhatsApp did not finish sending the files.');
    await sleep(1500);
  }
}

window.addEventListener('message', async e => {
  if (e.source !== window || e.data?.drs !== 'wa-item') return;
  const { id, item } = e.data;
  try {
    if (item.kind === 'text') await sendText(item.text); else await sendFiles(item.files);
    window.postMessage({ drs: 'wa-done', id }, location.origin);
  } catch (err) {
    window.postMessage({ drs: 'wa-done', id, error: String(err?.message || err) }, location.origin);
  }
});
