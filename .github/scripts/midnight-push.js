/**
 * Midnight Push Notification Script
 *
 * Sends a push notification via FCM to all registered devices,
 * prompting the app to capture GPS for the midnight location.
 *
 * Triggered by GitHub Actions at midnight CET/CEST.
 */

const admin = require('firebase-admin');

// Parse service account from GitHub secret
const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  databaseURL: 'https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app'
});

const db = admin.database();

async function sendMidnightPush() {
  console.log('Fetching FCM tokens from Firebase...');

  const snapshot = await db.ref('fcmTokens').once('value');
  const tokens = snapshot.val();

  if (!tokens) {
    console.log('No FCM tokens registered. Exiting.');
    process.exit(0);
  }

  const tokenList = Object.values(tokens).map(t => t.token).filter(Boolean);
  console.log(`Found ${tokenList.length} registered device(s)`);

  if (tokenList.length === 0) {
    console.log('No valid tokens. Exiting.');
    process.exit(0);
  }

  // Determine today's date in CET/CEST
  const now = new Date();
  const cetOffset = now.getTimezoneOffset(); // UTC offset in minutes
  // Just use UTC date since we fire at ~midnight CET
  const dateStr = now.toISOString().slice(0, 10);

  const message = {
    notification: {
      title: '🌙 Midnight Location Check',
      body: 'Tap to log where you are right now'
    },
    data: {
      action: 'capture-gps',
      date: dateStr,
      timestamp: String(Date.now())
    },
    webpush: {
      notification: {
        tag: 'midnight-gps',
        renotify: true,
        requireInteraction: true
      },
      fcmOptions: {
        link: 'https://midnight.cancomo.com/?capture=midnight'
      }
    }
  };

  // Send to each token individually (multicast not available for web push)
  let sent = 0;
  let failed = 0;
  const staleTokens = [];

  for (const token of tokenList) {
    try {
      await admin.messaging().send({ ...message, token });
      sent++;
      console.log(`Sent to: ${token.slice(0, 20)}...`);
    } catch (err) {
      failed++;
      console.warn(`Failed: ${token.slice(0, 20)}... — ${err.code || err.message}`);
      // Remove stale tokens
      if (err.code === 'messaging/registration-token-not-registered' ||
          err.code === 'messaging/invalid-registration-token') {
        staleTokens.push(token);
      }
    }
  }

  // Clean up stale tokens
  if (staleTokens.length > 0) {
    console.log(`Removing ${staleTokens.length} stale token(s)...`);
    const allTokens = snapshot.val();
    for (const [uid, data] of Object.entries(allTokens)) {
      if (staleTokens.includes(data.token)) {
        await db.ref('fcmTokens/' + uid).remove();
        console.log(`Removed token for ${data.email || uid}`);
      }
    }
  }

  console.log(`Done. Sent: ${sent}, Failed: ${failed}`);
  process.exit(0);
}

sendMidnightPush().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
