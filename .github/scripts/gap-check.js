/**
 * Gap Check Script
 *
 * Runs each morning (08:00 UK) and checks if yesterday's midnight location
 * is missing. If so, sends a Teams alert asking Steve to fill it in.
 * Also checks for any gaps in the last 7 days as a safety net.
 *
 * Triggered by GitHub Actions cron.
 */

const admin = require('firebase-admin');

const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
const teamsWebhook = process.env.TEAMS_WEBHOOK_URL;

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  databaseURL: 'https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app'
});

const db = admin.database();

function ukDateStr(d) {
  const opts = { timeZone: 'Europe/London', year: 'numeric', month: '2-digit', day: '2-digit' };
  const parts = new Intl.DateTimeFormat('en-GB', opts).formatToParts(d);
  return parts.find(p => p.type === 'year').value + '-' +
    parts.find(p => p.type === 'month').value + '-' +
    parts.find(p => p.type === 'day').value;
}

function formatDate(ds) {
  const d = new Date(ds + 'T00:00:00');
  const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  return days[d.getDay()] + ' ' + d.getDate() + ' ' + months[d.getMonth()];
}

async function sendTeamsAlert(message) {
  if (!teamsWebhook) {
    console.log('No Teams webhook configured — skipping alert');
    return;
  }
  await fetch(teamsWebhook, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      '@type': 'MessageCard',
      '@context': 'http://schema.org/extensions',
      themeColor: 'd97706',
      summary: 'Midnight Tracker — Gap Alert',
      sections: [{
        activityTitle: 'Midnight Tracker — Data Gap',
        activityImage: 'https://img.icons8.com/fluency/48/moon-satellite.png',
        text: message
      }]
    })
  });
  console.log('Teams alert sent');
}

async function checkGaps() {
  console.log('Checking for location data gaps...');

  // Get all location data
  const snapshot = await db.ref('locations').once('value');
  const locations = snapshot.val() || {};

  // Also check for Teams webhook in Firebase settings
  let webhook = teamsWebhook;
  if (!webhook) {
    const whSnap = await db.ref('settings/teamsWebhook').once('value');
    webhook = whSnap.val();
  }

  const now = new Date();
  const gaps = [];

  // Check last 7 days (not including today — today might not have data yet)
  for (let i = 1; i <= 7; i++) {
    const checkDate = new Date(now.getTime() - i * 86400000);
    const ds = ukDateStr(checkDate);
    const entry = locations[ds];
    const hasData = entry && (entry.city || entry.country);
    if (!hasData) {
      gaps.push(ds);
    }
  }

  if (gaps.length === 0) {
    console.log('No gaps found in the last 7 days. All good!');
    process.exit(0);
  }

  console.log('Found ' + gaps.length + ' gap(s): ' + gaps.join(', '));

  // Build alert message
  const yesterday = ukDateStr(new Date(now.getTime() - 86400000));
  const isYesterdayMissing = gaps.includes(yesterday);

  let msg = '';
  if (isYesterdayMissing && gaps.length === 1) {
    msg = '**Yesterday (' + formatDate(yesterday) + ')** has no midnight location recorded.<br><br>' +
      'Please open <a href="https://midnight.cancomo.com">midnight.cancomo.com</a> and log where you were.';
  } else {
    msg = '**' + gaps.length + ' day(s)** in the last week are missing midnight location data:<br><br>';
    gaps.forEach(ds => {
      msg += '• **' + formatDate(ds) + '** — no data<br>';
    });
    msg += '<br>Please open <a href="https://midnight.cancomo.com">midnight.cancomo.com</a> and fill in the gaps.';
  }

  await sendTeamsAlert(msg);
  process.exit(0);
}

checkGaps().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
