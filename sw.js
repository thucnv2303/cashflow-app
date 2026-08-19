// CashFlow Service Worker - updates, offline fallback and notifications
const CACHE_NAME = 'cashflow-v3.33';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys
        .filter(key => key.startsWith('cashflow-') && key !== CACHE_NAME)
        .map(key => caches.delete(key))))
      .then(() => clients.claim())
  );
});

// Network-first keeps installed iPhone PWAs current while retaining offline fallback.
self.addEventListener('fetch', (event) => {
  const request = event.request;
  if (request.method !== 'GET') return;

  const url = new URL(request.url);
  const isAppAsset = url.origin === self.location.origin
    && (request.mode === 'navigate' || /\.(?:html|js|css|json)$/.test(url.pathname));
  if (!isAppAsset) return;

  event.respondWith(
    fetch(new Request(request, { cache: 'reload' }))
      .then(response => {
        if (!response || !response.ok) return response;
        const copy = response.clone();
        return caches.open(CACHE_NAME)
          .then(cache => cache.put(request, copy))
          .then(() => response);
      })
      .catch(() => caches.match(request)
        .then(cached => cached || (request.mode === 'navigate' ? caches.match('./') : null))
        .then(cached => cached || Response.error()))
  );
});

// Listen for messages from the main app
self.addEventListener('message', (event) => {
  if (event.data && event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
    return;
  }
  if (event.data && event.data.type === 'CHECK_REMINDER') {
    checkAndNotify(event.data);
  }
});

function checkAndNotify(data) {
  const { memberName, avatarImg } = data;
  self.registration.showNotification('CashFlow - Nhắc nhở 📝', {
    body: `${memberName} ơi, hôm nay chưa nhập giao dịch nào. Đừng quên ghi chép nhé!`,
    icon: avatarImg || '/favicon.ico',
    badge: '/favicon.ico',
    tag: 'cashflow-reminder',
    requireInteraction: true
  });
}

// Handle notification click - open the app
self.addEventListener('notificationclick', (event) => {
  event.notification.close();
  event.waitUntil(
    clients.matchAll({ type: 'window', includeUncontrolled: true }).then((clientList) => {
      for (const client of clientList) {
        if (client.url.includes('cashflow') && 'focus' in client) {
          return client.focus();
        }
      }
      return clients.openWindow('./');
    })
  );
});
