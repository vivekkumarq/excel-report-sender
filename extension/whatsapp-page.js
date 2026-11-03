// Runs in WhatsApp Web's own page context.
// - 'wa-open': opens the logged-in account's own chat directly through WhatsApp Web's own modules (instant, no
//   reload, no "Starting chat" lookup). The account is found by its own ids, so the chat's label does not matter.
// - 'wa-item': sends one item into that chat. Text is pasted into the message box; images or PDFs are pasted as
//   files, which opens WhatsApp's preview, and its send button is pressed. Before every item the open chat is
//   checked again to be the own chat.
// Buttons are found by WhatsApp's icon names, which do not change with the interface language.
const sleep = ms => new Promise(r => setTimeout(r, ms));
const waitFor = async (fn, ms) => { for (const end = Date.now() + ms; ;) { const v = fn(); if (v) return v; if (Date.now() > end) return null; await sleep(150); } };
const composer = () => document.querySelector('#main footer [contenteditable="true"]');
const sendButtons = () => [...document.querySelectorAll('svg title')].filter(t => /(^|-)send(-filled)?$/.test(t.textContent)).map(t => t.closest('[role="button"], button')).filter(Boolean);
const previewSend = () => sendButtons().find(b => !b.closest('#main footer'));

// WhatsApp Web's internal modules; null if a future version renames them (then the caller uses the slower route).
function wa() {
  try {
    return { me: require('WAWebUserPrefsMeUser'), chats: require('WAWebChatCollection').ChatCollection, cmd: require('WAWebCmd').Cmd, find: require('WAWebFindChatAction') };
  } catch { return null; }
}
const myIds = w => [w.me.getMaybeMeLidUser?.(), w.me.getMaybeMePnUser?.()].filter(Boolean);
const isMine = (w, chat) => !!chat?.id && myIds(w).some(id => chat.id.equals ? chat.id.equals(id) : chat.id._serialized === id._serialized);

async function openOwnChat() {
  const w = wa();
  if (!w || !myIds(w).length) return { fallback: true };
  let chat = myIds(w).map(id => w.chats.get(id)).find(Boolean);
  for (const id of chat ? [] : myIds(w)) { // never messaged yourself yet: let WhatsApp create the chat
    try { const r = await w.find.findOrCreateLatestChat(id); chat = r?.chat ?? r; } catch { /* try the next id */ }
    if (chat) break;
  }
  if (!chat) return { fallback: true };
  if (!isMine(w, w.chats.getActive?.())) await w.cmd.openChatBottom({ chat });
  if (!await waitFor(() => isMine(w, w.chats.getActive?.()) && composer(), 15000)) throw new Error('Could not open your own chat in WhatsApp Web.');
  return {};
}

function checkOwnChat() {
  const w = wa();
  if (w && !isMine(w, w.chats.getActive?.())) throw new Error('Your own chat is no longer the open chat, so sending stopped.');
}

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
  await sleep(300);
}

async function sendFiles(files) {
  for (let k = 0; k < files.length; k += 30) { // WhatsApp takes up to 30 files per send
    const dt = new DataTransfer();
    files.slice(k, k + 30).forEach(f => dt.items.add(f));
    paste(dt);
    const btn = await waitFor(previewSend, 30000);
    if (!btn) throw new Error('WhatsApp did not open the file preview.');
    await sleep(600); // let the preview take in every file
    btn.click();
    if (!await waitFor(() => !previewSend(), 60000)) throw new Error('WhatsApp did not finish sending the files.');
    await sleep(800);
  }
}

window.addEventListener('message', async e => {
  if (e.source !== window || !/^wa-(open|item)$/.test(e.data?.drs)) return;
  const { id, drs, item } = e.data;
  try {
    let result = {};
    if (drs === 'wa-open') result = await openOwnChat();
    else { checkOwnChat(); if (item.kind === 'text') await sendText(item.text); else await sendFiles(item.files); }
    window.postMessage({ drs: 'wa-done', id, ...result }, location.origin);
  } catch (err) {
    window.postMessage({ drs: 'wa-done', id, error: String(err?.message || err) }, location.origin);
  }
});
