/* Scheduler service worker (added 2026-10-21).
   Only here so Chrome on Android/tablets offers "Install app" for the
   Scheduler, the same as the minister portal. It does nothing else: no
   offline caching and no fetch handling, so the app always loads fresh
   from GitHub Pages. */
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', (event) => event.waitUntil(self.clients.claim()));
