/**
 * One-time cleanup v2: Fix two issues in Firebase
 * 1. March 20 — remove fake flights (BA586, BA591, OX4, UB7 from Airbnb email)
 * 2. April 6-10 — fix place "Confirmation" → "Villa d Este", clear city "April"
 * Run with: node .github/scripts/fix-march20.js
 */
const admin = require('firebase-admin');

const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  databaseURL: 'https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app'
});
const db = admin.database();

async function main() {
  console.log('=== Firebase Cleanup v2 ===\n');

  const snap = await db.ref('locations').once('value');
  const all = snap.val() || {};
  const updates = {};
  let changes = 0;

  // --- Fix 1: March 20 — clear fake flights and booking source ---
  const mar20 = all['2026-03-20'];
  if (mar20) {
    console.log('March 20 current data:');
    console.log('  flights:', mar20.flights || '(none)');
    console.log('  bookingSource:', mar20.bookingSource || '(none)');
    console.log('  autoBooking:', mar20.autoBooking || '(none)');

    if (mar20.flights) {
      updates['locations/2026-03-20/flights'] = null;
      console.log('  → Clearing flights');
      changes++;
    }
    if (mar20.bookingSource) {
      updates['locations/2026-03-20/bookingSource'] = null;
      console.log('  → Clearing bookingSource');
      changes++;
    }
    if (mar20.autoBooking) {
      updates['locations/2026-03-20/autoBooking'] = null;
      console.log('  → Clearing autoBooking');
      changes++;
    }
  } else {
    console.log('March 20: no entry found');
  }

  // --- Fix 2: April 6-10 — fix place and city from Villa d'Este email ---
  const villaFixDates = ['2026-04-06', '2026-04-07', '2026-04-08', '2026-04-09', '2026-04-10'];
  for (const dateStr of villaFixDates) {
    const entry = all[dateStr];
    if (!entry) {
      console.log(`${dateStr}: no entry found`);
      continue;
    }
    console.log(`\n${dateStr} current data:`);
    console.log('  place:', entry.place || '(empty)');
    console.log('  city:', entry.city || '(empty)');

    if (entry.place === 'Confirmation' || !entry.place) {
      updates[`locations/${dateStr}/place`] = 'Villa d Este';
      console.log('  → Setting place to "Villa d Este"');
      changes++;
    }
    if (entry.city === 'April' || !entry.city) {
      updates[`locations/${dateStr}/city`] = 'Cernobbio';
      console.log('  → Setting city to "Cernobbio"');
      changes++;
    }
  }

  console.log(`\n${changes} updates to apply.`);

  if (changes === 0) {
    console.log('Nothing to do.');
    process.exit(0);
  }

  console.log('Updates:', JSON.stringify(updates, null, 2));
  await db.ref().update(updates);

  console.log('\nDone.');
  process.exit(0);
}

main().catch(err => { console.error(err); process.exit(1); });
