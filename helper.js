// Helper extension: shows whether it is installed, and builds its .zip from this site's own files on demand.
import { zipSync } from './vendor/fflate.js';

const FILES = ['manifest.json', 'background.js', 'site.js', 'whatsapp.js', 'whatsapp-page.js', 'icon16.png', 'icon48.png', 'icon128.png'];
const state = document.querySelector('#helper-state'), box = document.querySelector('#helper');

function refresh() {
  const v = document.documentElement.dataset.drsHelper;
  state.textContent = v ? `Helper installed (version ${v}).` : 'Helper not installed in this browser.';
  state.className = v ? 'note ok' : 'note';
  if (!v && !sessionStorage.getItem('helper-seen')) box.open = true;
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
