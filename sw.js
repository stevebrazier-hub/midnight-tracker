// ===== PUSH NOTIFICATION HANDLING =====
// IMPORTANT: Register our push listener BEFORE importing Firebase.
// Firebase's messaging SDK also registers a push listener, and we need
// ours to fire first to guarantee notification display.
self.addEventListener('push', event => {
  console.log('[SW] Push event received');

  let data = {};
  let title = 'Midnight Tracker';
  let body = 'Tap to log your midnight location';

  try {
    const payload = event.data?.json();
    // FCM messages with top-level notification field
    if (payload.notification) {
      title = payload.notification.title || title;
      body = payload.notification.body || body;
    }
    // FCM data payload
    data = payload.data || {};
    // Fallback: data-only messages might have title/body in data
    if (!payload.notification) {
      title = data.title || title;
      body = data.body || body;
    }
  } catch(e) {
    try { body = event.data?.text() || body; } catch(e2) {}
  }

  const captureType = data.captureType || 'midnight';
  const captureDate = data.date || '';

  event.waitUntil(
    // Show notification AND immediately tell any open app tabs to capture GPS.
    // If no tabs are open, silently open the app so GPS capture triggers automatically.
    Promise.all([
      self.registration.showNotification(title, {
        body: body,
        tag: captureType === 'midnight' ? 'midnight-gps' : 'bracket-gps-' + captureType,
        renotify: true,
        // Sticky for midnight and evening so a delayed push (e.g. cron firing
        // at 22:50 BST = 23:50 Italy on 2026-05-27) stays on the lock screen
        // until tapped, instead of clearing while the user is asleep.
        requireInteraction: captureType === 'midnight' || captureType === 'evening',
        data: { action: 'capture-gps', captureType: captureType, date: captureDate, timestamp: Date.now() }
      }),
      self.clients.matchAll({ type: 'window', includeUncontrolled: true }).then(clients => {
        const appClients = clients.filter(c => c.url.includes('midnight'));
        if (appClients.length > 0) {
          // App is open — tell it to capture GPS
          for (const client of appClients) {
            console.log('[SW] Auto-triggering GPS capture in open tab');
            client.postMessage({ type: 'midnight-gps-capture', captureType: captureType, date: captureDate });
          }
        } else {
          // No app tabs open — silently open the app with capture flag.
          // The app will auto-capture GPS on load via URL parameters.
          // This is the key improvement for overnight GPS reliability.
          console.log('[SW] No open tabs — silently opening app for GPS capture');
          const captureUrl = './?capture=' + captureType + '&date=' + captureDate + '&silent=1';
          return self.clients.openWindow(captureUrl).catch(err => {
            console.warn('[SW] Failed to open window for GPS capture:', err.message);
          });
        }
      })
    ])
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

// ===== FIREBASE MESSAGING SDK =====
// Required by firebase.messaging().getToken() in the app — without this,
// FCM token registration fails and no pushes are delivered.
importScripts('https://www.gstatic.com/firebasejs/10.8.0/firebase-app-compat.js');
importScripts('https://www.gstatic.com/firebasejs/10.8.0/firebase-messaging-compat.js');

firebase.initializeApp({
  apiKey: "AIzaSyAjZaVwXku1n1niJtkxvcKjXDHibSHHIRc",
  authDomain: "midnight-tracker-steve.firebaseapp.com",
  databaseURL: "https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "midnight-tracker-steve",
  storageBucket: "midnight-tracker-steve.firebasestorage.app",
  messagingSenderId: "860824921259",
  appId: "1:860824921259:web:b1c462b04dc3bf18ae03ee"
});

// Initialize messaging — this registers Firebase's own push handler,
// but our listener above was registered first so it runs first.
firebase.messaging();

// ===== CACHING (PWA offline support) =====
const CACHE_NAME = 'midnight-tracker-v32';
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
