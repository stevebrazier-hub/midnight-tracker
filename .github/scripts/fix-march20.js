/**
 * One-time cleanup: Remove false AT09 flight and wrong bookingSource from March 20, 2026.
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
  const ref = db.ref('locations/2026-03-20');
  const snap = await ref.once('value');
  const data = snap.val();

  if (!data) {
    console.log('No data for 2026-03-20');
    process.exit(0);
  }

  console.log('Before:', JSON.stringify(data, null, 2));

  // Remove the false flight and wrong booking source
  const updates = {};
  if (data.flights && data.flights.includes('AT09')) {
    // Remove AT09, keep any other flights
    const remaining = (data.flights || '').split(/,\s*/).filter(f => f !== 'AT09').join(', ');
    updates.flights = remaining || null;
  }
  if (data.bookingSource && data.bookingSource.includes('Airbnb')) {
    // Remove the Airbnb-related booking source
    const parts = (data.bookingSource || '').split(' | ').filter(p => !p.includes('Airbnb'));
    updates.bookingSource = parts.join(' | ') || null;
  }

  if (Object.keys(updates).length === 0) {
    console.log('Nothing to clean up.');
    process.exit(0);
  }

  console.log('Updates:', updates);
  await ref.update(updates);

  const after = await ref.once('value');
  console.log('After:', JSON.stringify(after.val(), null, 2));

  console.log('Done — cleaned up March 20.');
  process.exit(0);
}

main().catch(err => { console.error(err); process.exit(1); });
