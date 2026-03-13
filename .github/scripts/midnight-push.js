/**
 * Midnight Push Notification Script — Hourly Smart Scheduler
 *
 * Runs every hour via GitHub Actions. Checks current UK time, determines
 * which push notifications should have been sent by now, and sends any
 * that are missing. Firebase dedup ensures no duplicates.
 *
 * Daily cycle for a given date D:
 *   - Evening (10pm UK on day D)     → stored as brackets.evening on day D
 *   - Midnight (00:00 UK, start D+1) → stored as autoGps on day D
 *   - Morning (7am UK on day D+1)    → stored as brackets.morning on day D
 *
 * With 24 runs per day, even if GitHub delays crons by hours,
 * the next run catches up. Much more reliable than 6 specific crons.
 */

const admin = require('firebase-admin');

const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  databaseURL: 'https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app'
});

const db = admin.database();

// Get current UK date/time parts
function getUKTime() {
  const now = new Date();
  const ukStr = now.toLocaleString('en-GB', { timeZone: 'Europe/London' });
  // Format: "13/03/2026, 14:30:00"
  const [datePart, timePart] = ukStr.split(', ');
  const [day, month, year] = datePart.split('/');
  const [hour, minute] = timePart.split(':');
  return {
    year: parseInt(year), month: parseInt(month), day: parseInt(day),
    hour: parseInt(hour), minute: parseInt(minute),
    dateStr: `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`
  };
}

// Get yesterday's date string in UK time
function getUKYesterday() {
  const now = new Date();
  const yesterday = new Date(now.getTime() - 86400000);
  const ukDate = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Europe/London', year: 'numeric', month: '2-digit', day: '2-digit'
  }).formatToParts(yesterday);
  return ukDate.find(p => p.type === 'year').value + '-' +
    ukDate.find(p => p.type === 'month').value + '-' +
    ukDate.find(p => p.type === 'day').value;
}

// Notification text
const titles = {
  midnight: '\u{1F319} Midnight Location Check',
  evening: '\u{1F306} Evening Location Bracket',
  morning: '\u{1F305} Morning Location Bracket'
};
const bodies = {
  midnight: 'GPS captured automatically \u2014 tap only if you need to correct it',
  evening: '10pm GPS bracket \u2014 confirms where you are before midnight',
  morning: 'Morning GPS bracket \u2014 confirms where you woke up'
};

async function sendPush(captureType, dateStr, tokenList) {
  console.log(`  Sending ${captureType} push for ${dateStr}...`);

  const message = {
    notification: {
      title: titles[captureType],
      body: bodies[captureType]
    },
    data: {
      action: 'capture-gps',
      captureType: captureType,
      date: dateStr,
      timestamp: String(Date.now())
    },
    webpush: {
      notification: {
        tag: captureType === 'midnight' ? 'midnight-gps' : 'bracket-gps-' + captureType,
        renotify: 'true',
        requireInteraction: String(captureType === 'midnight')
      },
      fcmOptions: {
        link: 'https://midnight.cancomo.com/?capture=' + captureType + '&date=' + dateStr
      }
    }
  };

  let sent = 0, failed = 0;
  const staleTokens = [];

  for (const token of tokenList) {
    try {
      await admin.messaging().send({ ...message, token });
      sent++;
      console.log(`  Sent to: ${token.slice(0, 20)}...`);
    } catch (err) {
      failed++;
      console.warn(`  Failed: ${token.slice(0, 20)}... — ${err.code || err.message}`);
      if (err.code === 'messaging/registration-token-not-registered' ||
          err.code === 'messaging/invalid-registration-token') {
        staleTokens.push(token);
      }
    }
  }

  console.log(`  Result: sent=${sent}, failed=${failed}`);
  return staleTokens;
}

async function main() {
  // Manual override: if CAPTURE_TYPE is set, send that specific type
  const manualType = process.env.CAPTURE_TYPE;
  if (manualType) {
    console.log(`Manual trigger: capture type = ${manualType}`);
  }

  // Get tokens
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

  const uk = getUKTime();
  const ukYesterday = getUKYesterday();
  console.log(`UK time: ${uk.hour}:${String(uk.minute).padStart(2, '0')} on ${uk.dateStr}`);
  console.log(`UK yesterday: ${ukYesterday}`);

  // Determine what should have been sent by now
  const toSend = [];

  if (manualType) {
    // Manual: send the requested type for the appropriate date
    const date = manualType === 'evening' ? uk.dateStr :
                 manualType === 'midnight' ? ukYesterday :
                 manualType === 'morning' ? ukYesterday : uk.dateStr;
    toSend.push({ type: manualType, date: date });
  } else {
    // Auto-detect based on UK time:

    // Evening bracket: due after 10pm UK, attributed to today
    if (uk.hour >= 22) {
      toSend.push({ type: 'evening', date: uk.dateStr });
    }

    // Midnight capture: due after midnight UK, attributed to yesterday
    // (hour 0-6 means we just passed midnight, so yesterday needs midnight)
    if (uk.hour >= 0 && uk.hour <= 8) {
      toSend.push({ type: 'midnight', date: ukYesterday });
    }

    // Morning bracket: due after 7am UK, attributed to yesterday
    if (uk.hour >= 7 && uk.hour <= 12) {
      toSend.push({ type: 'morning', date: ukYesterday });
    }

    // Also check: did yesterday's evening bracket get sent?
    // (if it's now past midnight and evening was missed)
    if (uk.hour >= 0 && uk.hour <= 4) {
      toSend.push({ type: 'evening', date: ukYesterday });
    }
  }

  if (toSend.length === 0) {
    console.log('Nothing due at this hour. Exiting.');
    process.exit(0);
  }

  console.log(`Captures to check: ${toSend.map(s => s.type + '(' + s.date + ')').join(', ')}`);

  // Check Firebase for each and send what's missing
  let allStaleTokens = [];
  let totalSent = 0;

  for (const { type, date } of toSend) {
    const entrySnap = await db.ref('locations/' + date).once('value');
    const entry = entrySnap.val();

    let alreadyDone = false;
    if (type === 'midnight' && entry?.autoGps) alreadyDone = true;
    if (type === 'evening' && entry?.brackets?.evening?.lat) alreadyDone = true;
    if (type === 'morning' && entry?.brackets?.morning?.lat) alreadyDone = true;

    if (alreadyDone) {
      console.log(`${type} for ${date}: already captured — skipping.`);
      continue;
    }

    const stale = await sendPush(type, date, tokenList);
    allStaleTokens.push(...stale);
    totalSent++;
  }

  // Clean up stale tokens
  if (allStaleTokens.length > 0) {
    console.log(`Removing ${allStaleTokens.length} stale token(s)...`);
    const allTokenData = snapshot.val();
    for (const [uid, data] of Object.entries(allTokenData)) {
      if (allStaleTokens.includes(data.token)) {
        await db.ref('fcmTokens/' + uid).remove();
        console.log(`Removed token for ${data.email || uid}`);
      }
    }
  }

  console.log(`Done. Sent ${totalSent} push notification(s).`);
  process.exit(0);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
