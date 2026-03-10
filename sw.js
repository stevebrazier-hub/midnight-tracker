// ===== PUSH NOTIFICATION HANDLING =====
// FCM messages with a top-level notification field are auto-displayed by Chrome.
// This push listener is a safety net: if auto-display doesn't fire for any
// reason, we show the notification ourselves so Chrome never falls back to
// the generic "Tap to copy the URL" message.
self.addEventListener('push', event => {
  console.log('[SW] Push event received');

  let data = {};
  let title = 'Midnight Tracker';
  let body = 'Tap to log your midnight location';

  try {
    const payload = event.data?.json();
    // FCM puts auto-display fields in notification, custom fields in data
    if (payload.notification) {
      title = payload.notification.title || title;
      body = payload.notification.body || body;
    }
    data = payload.data || {};
  } catch(e) {
    try { body = event.data?.text() || body; } catch(e2) {}
  }

  const captureType = data.captureType || 'midnight';
  const captureDate = data.date || '';

  event.waitUntil(
    self.registration.showNotification(title, {
      body: body,
      tag: captureType === 'midnight' ? 'midnight-gps' : 'bracket-gps-' + captureType,
      renotify: true,
      requireInteraction: captureType === 'midnight',
      data: { action: 'capture-gps', captureType: captureType, date: captureDate, timestamp: Date.now() }
    })
  );
});

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
const CACHE_NAME = 'midnight-tracker-v26';
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
