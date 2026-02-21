/**
 * Restore Backup Script
 *
 * Restores Firebase data from a JSON backup file.
 * Usage: node restore-backup.js [path-to-backup.json]
 *
 * If no path given, restores from backups/latest.json
 *
 * Environment variables required:
 *   FIREBASE_SERVICE_ACCOUNT - Firebase service account JSON
 */

const admin = require('firebase-admin');
const fs = require('fs');
const path = require('path');

const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  databaseURL: 'https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app'
});
const db = admin.database();

async function main() {
  const backupPath = process.argv[2] || path.join(process.cwd(), 'backups', 'latest.json');

  if (!fs.existsSync(backupPath)) {
    console.error('Backup file not found:', backupPath);
    process.exit(1);
  }

  console.log('=== Midnight Tracker — RESTORE ===');
  console.log('Restoring from:', backupPath);

  const data = JSON.parse(fs.readFileSync(backupPath, 'utf8'));
  const locationCount = data.locations ? Object.keys(data.locations).length : 0;
  const presetCount = Array.isArray(data.presets) ? data.presets.length : 0;

  console.log(`Data: ${locationCount} locations, ${presetCount} presets`);
  console.log('\n⚠️  This will OVERWRITE all current Firebase data.');
  console.log('Press Ctrl+C within 5 seconds to cancel...\n');

  await new Promise(resolve => setTimeout(resolve, 5000));

  console.log('Restoring...');
  await db.ref().set(data);
  console.log('✅ Restore complete.');
  console.log(`Restored ${locationCount} locations, ${presetCount} presets, and all settings.`);

  process.exit(0);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
