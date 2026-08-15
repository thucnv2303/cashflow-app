// CashFlow Service Worker - Background notification support
const CACHE_NAME = 'cashflow-v2';

self.addEventListener('install', (event) => {
  self.skipWaiting();
});

self.addEventListener('activate', (event) => {
  event.waitUntil(clients.claim());
});

// Listen for messages from the main app
self.addEventListener('message', (event) => {
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
