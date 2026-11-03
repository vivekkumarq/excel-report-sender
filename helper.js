// Helper extension: shows whether it is installed, and builds its .zip from this site's own files on demand.
import { zipSync } from './vendor/fflate.js';

const FILES = ['manifest.json', 'background.js', 'site.js', 'whatsapp.js', 'whatsapp-page.js', 'icon16.png', 'icon48.png', 'icon128.png'];
const state = document.querySelector('#helper-state'), box = document.querySelector('#helper');

const latest = fetch('extension/manifest.json').then(r => r.json()).then(m => m.version).catch(() => null);
const newer = (a, b) => { const x = a.split('.').map(Number), y = b.split('.').map(Number); for (let i = 0; i < 3; i++) if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); return false; };

async function refresh() {
  const v = document.documentElement.dataset.drsHelper, l = await latest;
  const old = v && l && newer(l, v);
  state.textContent = !v ? 'Helper not installed in this browser.'
    : old ? `Helper ${v} is installed; version ${l} is faster. Download it again, unzip it over the old folder, then press the reload icon on the helper's card in chrome://extensions.`
      : `Helper installed (version ${v}).`;
  state.className = v && !old ? 'note ok' : old ? 'note err' : 'note';
  if ((!v || old) && !sessionStorage.getItem('helper-seen')) box.open = true;
  sessionStorage.setItem('helper-seen', '1');
}
refresh();
addEventListener('message', e => { if (e.source === window && e.data?.drs === 'ready') refresh(); });

document.querySelector('#helper-zip').onclick = async () => {
  const entries = await Promise.all(FILES.map(async f => [`report-helper/${f}`, new Uint8Array(await (await fetch(`extension/${f}`)).arrayBuffer())]));
  const url = URL.createObjectURL(new Blob([zipSync(Object.fromEntries(entries))], { type: 'application/zip' }));
  Object.assign(document.createElement('a'), { href: url, download: 'report-helper.zip' }).click();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
};
