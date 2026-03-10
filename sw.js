// ===== PUSH NOTIFICATION HANDLING =====
// Chrome auto-displays notifications from FCM messages that include a
// webpush.notification field. No push event listener needed.
// The notificationclick handler below runs when the user taps the notification.

// When user taps the notification — open the app with auto-capture flag
self.addEventListener('notificationclick', event => {
  console.log('[SW] Notification clicked');
  const captureDate = event.notification.data?.date || '';
  const captureType = event.notification.data?.captureType || 'midnight';
  event.notification.close();

  event.waitUntil(
    self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clients => {
      // If app is already open, focus it and tell it to capture GPS
      for (const client of clients) {
        if (client.url.includes('midnight') && 'focus' in client) {
          client.postMessage({ type: 'midnight-gps-capture', captureType: captureType, date: captureDate });
          return client.focus();
        }
      }
      // Otherwise open the app with capture flag + date + type
      const url = './?capture=' + captureType + '&date=' + captureDate;
      return self.clients.openWindow(url);
    })
  );
});

// ===== CACHING (PWA offline support) =====
const CACHE_NAME = 'midnight-tracker-v25';
const ASSETS = [
  './index.html',
  './manifest.json',
  'https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap'
];

self.addEventListener('install', e => {
  e.waitUntil(caches.open(CACHE_NAME).then(c => c.addAll(ASSETS)));
  self.skipWaiting();
});

self.addEventListener('activate', e => {
  e.waitUntil(
    caches.keys().then(keys =>
      Promise.all(keys.filter(k => k !== CACHE_NAME).map(k => caches.delete(k)))
    )
  );
  self.clients.claim();
});

self.addEventListener('fetch', e => {
  // Don't cache Firebase or external API requests
  if (e.request.url.includes('firebasejs') ||
      e.request.url.includes('googleapis') ||
      e.request.url.includes('nominatim') ||
      e.request.url.includes('gstatic')) {
    return;
  }

  e.respondWith(
    fetch(e.request)
      .then(r => {
        const clone = r.clone();
        caches.open(CACHE_NAME).then(c => c.put(e.request, clone));
        return r;
      })
      .catch(() => caches.match(e.request))
  );
});
