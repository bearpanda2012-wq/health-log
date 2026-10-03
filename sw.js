/* ดันดี — service worker: เปิดแอปได้แม้ไม่มีเน็ต
   - หน้าแอป (index.html ฯลฯ): ลองเน็ตก่อน ไม่มีเน็ตใช้ของที่เก็บไว้
   - ไลบรารี/ฟอนต์จาก CDN: ใช้ของที่เก็บไว้ก่อน
   - Google Apps Script (ข้อมูลสุขภาพ): ไม่เก็บเลย */
const CACHE = 'dundee-v2';
const SHELL = ['./', './index.html', './manifest.webmanifest', './icon.svg', './icon-192.png', './icon-512.png',
  'https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js'];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE).then(c => Promise.all(SHELL.map(u => c.add(u).catch(() => {})))).then(() => self.skipWaiting()));
});
self.addEventListener('activate', e => {
  e.waitUntil(caches.keys().then(ks => Promise.all(ks.filter(k => k !== CACHE).map(k => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener('fetch', e => {
  const req = e.request;
  if (req.method !== 'GET') return;
  const u = new URL(req.url);
  if (/script\.google(usercontent)?\.com$/.test(u.hostname)) return; // ข้อมูลจริง ไม่แคช
  const sameOrigin = u.origin === self.location.origin;
  const cdn = /(^|\.)(cdnjs\.cloudflare\.com|fonts\.googleapis\.com|fonts\.gstatic\.com)$/.test(u.hostname);
  if (sameOrigin) {
    e.respondWith(fetch(req).then(res => {
      if (res.ok) { const cp = res.clone(); caches.open(CACHE).then(c => c.put(req, cp)); }
      return res;
    }).catch(() => caches.match(req, { ignoreSearch: true }).then(r => r || caches.match('./index.html'))));
  } else if (cdn) {
    e.respondWith(caches.match(req).then(r => r || fetch(req).then(res => {
      const cp = res.clone(); caches.open(CACHE).then(c => c.put(req, cp)); return res;
    })));
  }
});
