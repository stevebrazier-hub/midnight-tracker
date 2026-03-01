/**
 * ONE-OFF Cleanup Script
 *
 * Removes bogus future "manual" entries that were not created by the user.
 * These are entries with no autoGps, no autoBooking, no lat/lon — just
 * pre-filled home address data for future dates.
 *
 * Safe: only deletes entries AFTER today that have no GPS or booking data.
 *
 * Run once via: node .github/scripts/cleanup-future-manual.js
 */

const admin = require('firebase-admin');

const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  databaseURL: 'https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app'
});
const db = admin.database();

function fmtDate(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

async function main() {
  const today = fmtDate(new Date());
  console.log('=== Cleanup bogus future manual entries ===');
  console.log('Today:', today);

  const snapshot = await db.ref('locations').once('value');
  const locations = snapshot.val() || {};

  const removals = {};
  let count = 0;

  for (const [dateStr, entry] of Object.entries(locations)) {
    // Only future dates (after today)
    if (dateStr <= today) continue;

    // Only "manual" entries — no GPS, no booking sync
    if (entry.autoGps || entry.autoBooking || entry.gpsConfirmed || entry.lat) continue;

    // Only entries with no meaningful user data (no flights, no notes)
    if (entry.flights || entry.notes) continue;

    console.log(`  DELETE ${dateStr}: ${entry.place || ''}, ${entry.city || ''}, ${entry.country || ''}`);
    removals['locations/' + dateStr] = null;
    count++;
  }

  if (count > 0) {
    await db.ref().update(removals);
    console.log(`\nRemoved ${count} bogus future entries.`);
  } else {
    console.log('\nNo bogus entries found.');
  }

  process.exit(0);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
