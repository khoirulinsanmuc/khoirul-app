self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
// Chrome mensyaratkan handler fetch agar app dianggap bisa diinstal
self.addEventListener('fetch', e => {
  e.respondWith(fetch(e.request));
});