// Service worker: lets the installed app open with no network, and receives workbooks shared to it from other
// apps on a phone (Web Share Target). Pages and code are fetched fresh when online (network first), so a new
// version of the site is picked up on the next visit; the copy kept here is only used when offline. Vendored
// files (fonts, libraries) never change once published, so those come from the cache first.
const CACHE = 'drs-v1', SHARE = 'drs-share';

self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));

self.addEventListener('fetch', e => {
  const req = e.request, url = new URL(req.url);
  if (url.origin !== location.origin) return;
  if (req.method === 'POST' && url.pathname.endsWith('/share')) return e.respondWith(receive(req));
  if (req.method !== 'GET' || /\.(xlsx|xlsm)$/i.test(url.pathname)) return; // never keep workbooks
  const key = url.origin + url.pathname; // one copy per file, whatever the query string
  const save = res => { if (res.ok && res.type === 'basic') { const copy = res.clone(); e.waitUntil(caches.open(CACHE).then(c => c.put(key, copy))); } return res; };
  const offline = () => caches.match(key).then(hit => hit || (req.mode === 'navigate' ? caches.match(new URL('./', self.registration.scope).href) : null)).then(hit => hit || Response.error());
  e.respondWith(url.pathname.includes('/vendor/')
    ? caches.match(key).then(hit => hit || fetch(req).then(save))
    : fetch(req).then(save, offline));
});

// The share sheet POSTs the file here; keep it for the page, then open the page, which picks it up.
async function receive(req) {
  const file = (await req.formData()).get('file');
  if (file instanceof File) {
    await (await caches.open(SHARE)).put('shared', new Response(file, { headers: { 'content-type': file.type || 'application/octet-stream', 'x-name': encodeURIComponent(file.name) } }));
  }
  return Response.redirect(new URL('./?shared=1', self.registration.scope).href, 303);
}
