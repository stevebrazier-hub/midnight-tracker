/**
 * One-time cleanup: Remove bogus hotel entries created by run #118.
 * These entries have place="and late check out upon availability" and city="Tire"
 * from the Villa d'Este email full-body parsing issue.
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
  const snap = await db.ref('locations').once('value');
  const all = snap.val() || {};

  const updates = {};
  let cleaned = 0;

  for (const [dateStr, data] of Object.entries(all)) {
    // Identify bogus entries: place contains "late check out upon availability" or city is "Tire"
    const isBogus = (data.place && data.place.includes('late check out upon availability')) ||
                    (data.city === 'Tire' && data.autoBooking);

    if (!isBogus) continue;

    if (data.autoGps || data.gpsConfirmed || data.brackets) {
      // GPS-confirmed entry — only clean the booking fields, keep GPS data
      console.log(`  ${dateStr}: Cleaning booking fields from GPS-confirmed entry`);
      if (data.place && data.place.includes('late check out')) updates[`locations/${dateStr}/place`] = null;
      if (data.city === 'Tire') updates[`locations/${dateStr}/city`] = data.autoGps ? data.city : null;
      if (data.bookingSource) updates[`locations/${dateStr}/bookingSource`] = null;
      if (data.autoBooking) updates[`locations/${dateStr}/autoBooking`] = null;
      // Remove bogus flights that came from the same email
      if (data.flights && /\b(DX4|UB7)\b/.test(data.flights)) {
        const cleaned = data.flights.split(/,\s*/).filter(f => !['DX4', 'UB7'].includes(f)).join(', ');
        updates[`locations/${dateStr}/flights`] = cleaned || null;
      }
    } else {
      // Pure booking entry with no GPS — delete entirely
      console.log(`  ${dateStr}: Removing bogus booking entry entirely`);
      updates[`locations/${dateStr}`] = null;
    }
    cleaned++;
  }

  if (cleaned === 0) {
    console.log('No bogus entries found.');
    process.exit(0);
  }

  console.log(`\nCleaning ${cleaned} bogus entries...`);
  await db.ref().update(updates);

  console.log('Done — cleaned up bogus entries from run #118.');
  process.exit(0);
}

main().catch(err => { console.error(err); process.exit(1); });
