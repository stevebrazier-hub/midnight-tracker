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

  // Determine the date that just ENDED at midnight in UK time (GMT/BST).
  // The cron fires at or just after midnight, so new Date() gives the new day.
  // Subtract 2 minutes to get the day that just finished — that's the day
  // whose "midnight location" we're recording (e.g. Sunday 22nd, not Monday 23rd).
  const now = new Date();
  const justBeforeMidnight = new Date(now.getTime() - 120000);
  const ukDate = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Europe/London', year: 'numeric', month: '2-digit', day: '2-digit'
  }).formatToParts(justBeforeMidnight);
  const dateStr = ukDate.find(p => p.type === 'year').value + '-' +
    ukDate.find(p => p.type === 'month').value + '-' +
    ukDate.find(p => p.type === 'day').value;

  const message = {
    notification: {
      title: '🌙 Midnight Location Check',
      body: 'GPS captured automatically — tap only if you need to correct it'
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
