/* Scheduler service worker (added 2026-10-21).
   Lets Chrome on Android/tablets offer "Install app" for the Scheduler,
   the same as the minister portal, and (added 2026-10-22) shows the admin
   push notifications the intake Worker sends when a new intake form is
   submitted (see /api/push/admin in cloudflare-worker/src/index.js).
   Tapping one opens that form. Still no offline caching and no fetch
   handling, so the app always loads fresh from GitHub Pages. */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));

self.addEventListener('push', (event) => {
  let data = {};
  try { data = event.data ? event.data.json() : {}; } catch (e) {}
  event.waitUntil(Promise.all([
    self.registration.showNotification(data.title || 'FPHM Admin', {
      body: data.body || '',
      tag: data.tag || undefined,
      icon: 'icons/icon-192.png',
      badge: 'icons/icon-192.png',
      data: { intakeId: data.intakeId || null }
    }),
    // Lets an open Scheduler refresh its Messages tab right away.
    self.clients.matchAll({ type: 'window' }).then((wins) => wins.forEach((w) => w.postMessage({ type: 'push' })))
  ]));
});

self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  const intakeId = (event.notification.data || {}).intakeId;
  const url = self.registration.scope + (intakeId ? '?intake=' + intakeId : '');
  event.waitUntil((async () => {
    const wins = await self.clients.matchAll({ type: 'window', includeUncontrolled: true });
    // Exact page match: at the repo root this worker's scope also covers
    // /portal/, /intake/ and /demo/, which must not be picked.
    const page = (u) => u.split(/[?#]/)[0].replace(/index\.html$/, '');
    const open = wins.find((w) => page(w.url) === self.registration.scope);
    if (open) {
      // An open Scheduler reloads its data and opens the form itself, so
      // the admin stays signed in and keeps their place.
      await open.focus();
      if (intakeId) open.postMessage({ type: 'open-intake', intakeId });
      return;
    }
    await self.clients.openWindow(url);
  })());
});
