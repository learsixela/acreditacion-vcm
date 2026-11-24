// Service worker mínimo: permite instalar la app. No guarda nada en caché,
// así la página siempre usa la versión publicada y los datos en vivo.
self.addEventListener('install', () => self.skipWaiting());
self.addEventListener('activate', e => e.waitUntil(self.clients.claim()));
self.addEventListener('fetch', () => {});
