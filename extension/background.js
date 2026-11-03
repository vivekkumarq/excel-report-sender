// Opens (or brings forward) WhatsApp Web for a send job, and relays progress back to the report page.
chrome.runtime.onMessage.addListener((m, sender, reply) => {
  if (m?.drs === 'open') {
    open(sender.tab.id).then(() => reply({}), e => reply({ error: e.message }));
    return true; // reply is async
  }
  if (m?.drs === 'status') {
    chrome.storage.local.get('site').then(({ site }) => site != null && chrome.tabs.sendMessage(site, m).catch(() => {}));
  }
});

async function open(siteTab) {
  await chrome.storage.local.set({ site: siteTab });
  const [tab] = await chrome.tabs.query({ url: 'https://web.whatsapp.com/*' });
  if (!tab) { await chrome.tabs.create({ url: 'https://web.whatsapp.com/' }); return; }
  await chrome.tabs.update(tab.id, { active: true });
  await chrome.windows.update(tab.windowId, { focused: true });
  // A tab opened before the helper was installed has no helper script yet: reload it.
  await chrome.tabs.sendMessage(tab.id, { drs: 'start' }).catch(() => chrome.tabs.reload(tab.id));
}
