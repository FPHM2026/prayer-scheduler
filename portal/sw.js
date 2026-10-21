/* Minister portal service worker (added 2026-10-21).
   Only job: show push notifications sent by the intake Worker's
   /api/push routes (see cloudflare-worker/src/index.js) and open the
   portal when one is tapped. Deliberately no offline caching - the
   portal always loads fresh from GitHub Pages, same as before. */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));

self.addEventListener('push', (event) => {
  let data = {};
  try { data = event.data ? event.data.json() : {}; } catch (e) {}
  event.waitUntil(self.registration.showNotification(data.title || 'My FPHM Schedule', {
    body: data.body || '',
    tag: data.tag || undefined,
    icon: 'icons/icon-192.png',
    badge: 'icons/icon-192.png'
  }));
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const url = self.registration.scope;
  event.waitUntil((async () => {
    const wins = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    const open = wins.find((w) => w.url.startsWith(url));
    if (open) { await open.focus(); open.navigate && open.navigate(url).catch(() => {}); return; }
    await self.clients.openWindow(url);
  })());
});
