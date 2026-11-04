// Helper extension: shows whether it is installed (header pill and the helper card), and builds its .zip from
// this site's own files on demand. The extension marks the page with data-drs-helper="<version>".
import { zipSync } from './vendor/fflate.js';

const FILES = ['manifest.json', 'background.js', 'site.js', 'whatsapp.js', 'whatsapp-page.js', 'icon16.png', 'icon48.png', 'icon128.png'];
const state = document.querySelector('#helper-state'), box = document.querySelector('#helper'), pill = document.querySelector('#helper-pill');
const phone = navigator.userAgentData?.mobile ?? matchMedia('(pointer: coarse) and (max-width: 820px)').matches; // no extensions there

const latest = fetch('extension/manifest.json').then(r => r.json()).then(m => m.version).catch(() => null);
const newer = (a, b) => { const x = a.split('.').map(Number), y = b.split('.').map(Number); for (let i = 0; i < 3; i++) if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); return false; };

async function refresh() {
  const v = document.documentElement.dataset.drsHelper, l = await latest;
  const old = v && l && newer(l, v);
  state.textContent = !v ? 'The helper is not installed in this browser.'
    : old ? `Helper ${v} is installed; version ${l} is newer. Download it again, unzip it over the old folder, then press the reload icon on the helper's card in chrome://extensions.`
      : `Helper installed (version ${v}).`;
  state.className = v && !old ? 'note ok' : old ? 'note err' : 'note';
  pill.className = `pill ${v ? (old ? 'old' : 'on') : ''}`;
  pill.querySelector('.txt').textContent = !v ? (phone ? 'Share mode' : 'Helper not installed') : old ? 'Helper update' : 'Helper on';
  pill.title = state.textContent;
  try {
    if ((!v || old) && !phone && !sessionStorage.getItem('helper-seen')) box.open = true;
    sessionStorage.setItem('helper-seen', '1');
  } catch { /* storage blocked: leave the card as it is */ }
  dispatchEvent(new Event('drs-helper'));
}
refresh();

document.querySelector('#helper-zip').onclick = async () => {
  const entries = await Promise.all(FILES.map(async f => [`report-helper/${f}`, new Uint8Array(await (await fetch(`extension/${f}`)).arrayBuffer())]));
  const url = URL.createObjectURL(new Blob([zipSync(Object.fromEntries(entries))], { type: 'application/zip' }));
  Object.assign(document.createElement('a'), { href: url, download: 'report-helper.zip' }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
};
