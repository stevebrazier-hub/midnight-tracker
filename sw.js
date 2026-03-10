// Import Firebase Messaging for background push handling
importScripts('https://www.gstatic.com/firebasejs/10.8.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.8.0/firebase-messaging-compat.js');

// Firebase config (must match the app)
firebase.initializeApp({
  apiKey: "AIzaSyAjZaVwXku1n1niJtkxvcKjXDHibSHHIRc",
  authDomain: "midnight-tracker-steve.firebaseapp.com",
  databaseURL: "https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "midnight-tracker-steve",
  storageBucket: "midnight-tracker-steve.firebasestorage.app",
  messagingSenderId: "860824921259",
  appId: "1:860824921259:web:b1c462b04dc3bf18ae03ee"
});

const messaging = firebase.messaging();

// Raw push event listener — this ALWAYS fires for every push, regardless of
// whether Firebase's onBackgroundMessage handles it. We use this as the primary
// handler to guarantee notification display.
let pushHandled = false;
self.addEventListener('push', event => {
  console.log('[SW] Push event received');
  pushHandled = true;

  let data = {};
  try {
    const payload = event.data?.json();
    // FCM wraps data-only messages: payload.data contains our fields
    data = payload?.data || payload || {};
  } catch(e) {
    try { data = { body: event.data?.text() || '' }; } catch(e2) {}
  }

  const title = data.title || 'Midnight Tracker';
  const body = data.body || 'Tap to log your midnight location';
  const captureDate = data.date || '';
  const captureType = data.captureType || 'midnight';

  event.waitUntil(
    self.registration.showNotification(title, {
      body: body,
      icon: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" rx="20" fill="%231b2838"/><text x="50" y="65" font-size="50" text-anchor="middle" fill="white">🌙</text></svg>',
      badge: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><circle cx="50" cy="50" r="50" fill="%2300b8a9"/></svg>',
      tag: captureType === 'midnight' ? 'midnight-gps' : 'bracket-gps-' + captureType,
      renotify: true,
      requireInteraction: captureType === 'midnight',
      data: { action: 'capture-gps', captureType: captureType, date: captureDate, timestamp: Date.now() }
    })
  );
});

// Keep Firebase onBackgroundMessage as backup (may not fire for data-only messages)
messaging.onBackgroundMessage(payload => {
  console.log('[SW] onBackgroundMessage received (pushHandled=' + pushHandled + ')');
  // If our push handler already showed the notification, skip
  if (pushHandled) { pushHandled = false; return; }

  const title = payload.data?.title || payload.notification?.title || 'Midnight Tracker';
  const body = payload.data?.body || payload.notification?.body || 'Tap to log your midnight location';
  const captureDate = payload.data?.date || '';
  const captureType = payload.data?.captureType || 'midnight';

  return self.registration.showNotification(title, {
    body: body,
    icon: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><rect width="100" height="100" rx="20" fill="%231b2838"/><text x="50" y="65" font-size="50" text-anchor="middle" fill="white">🌙</text></svg>',
    badge: 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><circle cx="50" cy="50" r="50" fill="%2300b8a9"/></svg>',
    tag: captureType === 'midnight' ? 'midnight-gps' : 'bracket-gps-' + captureType,
    renotify: true,
    requireInteraction: captureType === 'midnight',
    data: { action: 'capture-gps', captureType: captureType, date: captureDate, timestamp: Date.now() }
  });
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
const CACHE_NAME = 'midnight-tracker-v23';
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
