/**
 * Daily Backup Script
 *
 * Exports all Firebase data to:
 *   1. backups/latest.json — full Firebase snapshot (for restore)
 *   2. backups/latest.xlsx — human-readable spreadsheet (for audit)
 *   3. backups/history/backup-YYYY-MM-DD.json — dated archive (keep 5 days)
 *
 * Runs daily via GitHub Actions.
 */

const admin = require('firebase-admin');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// ===== FIREBASE INIT =====
const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  databaseURL: 'https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app'
});
const db = admin.database();

const BACKUP_DIR = path.join(process.cwd(), 'backups');
const HISTORY_DIR = path.join(BACKUP_DIR, 'history');
const MAX_HISTORY = 5;

function fmtDate(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

async function main() {
  const today = fmtDate(new Date());
  console.log('=== Midnight Tracker — Daily Backup ===');
  console.log('Date:', today);

  // Create directories
  if (!fs.existsSync(BACKUP_DIR)) fs.mkdirSync(BACKUP_DIR, { recursive: true });
  if (!fs.existsSync(HISTORY_DIR)) fs.mkdirSync(HISTORY_DIR, { recursive: true });

  // 1. Export full Firebase snapshot
  console.log('\nReading Firebase data...');
  const snapshot = await db.ref().once('value');
  const allData = snapshot.val() || {};
  const locations = allData.locations || {};
  const settings = allData.settings || {};
  const presets = allData.presets || [];

  const entryCount = Object.keys(locations).length;
  console.log(`Found ${entryCount} location entries`);

  // 2. Write JSON backup
  const jsonData = JSON.stringify(allData, null, 2);
  fs.writeFileSync(path.join(BACKUP_DIR, 'latest.json'), jsonData);
  fs.writeFileSync(path.join(HISTORY_DIR, `backup-${today}.json`), jsonData);
  console.log('Wrote latest.json and history/' + `backup-${today}.json`);

  // 3. Build XLSX spreadsheet
  const wb = XLSX.utils.book_new();

  // Locations sheet
  const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const rows = Object.keys(locations).sort().map(dateStr => {
    const e = locations[dateStr];
    const d = new Date(dateStr + 'T00:00:00');
    return {
      'Date': dateStr,
      'Day': dayNames[d.getDay()],
      'Place': e.place || '',
      'City': e.city || '',
      'Country': e.country || '',
      'Flights': e.flights || '',
      'Notes': e.notes || '',
      'Latitude': e.lat || '',
      'Longitude': e.lon || '',
      'Working': e.working ? 'Yes' : '',
      'Source': e.gpsConfirmed ? 'GPS + Manual' : e.autoBooking ? 'Booking' : e.autoGps ? 'GPS' : (e.city ? 'Manual' : ''),
      'Country conflict': e.countryConflict || '',
      'Booking source': e.bookingSource || '',
      'Auto GPS': e.autoGps ? 'Yes' : '',
      'Auto booking': e.autoBooking ? 'Yes' : '',
      'GPS confirmed': e.gpsConfirmed ? 'Yes' : '',
    };
  });

  if (rows.length) {
    const ws = XLSX.utils.json_to_sheet(rows);

    // Set column widths
    ws['!cols'] = [
      { wch: 12 }, // Date
      { wch: 10 }, // Day
      { wch: 30 }, // Place
      { wch: 18 }, // City
      { wch: 12 }, // Country
      { wch: 15 }, // Flights
      { wch: 20 }, // Notes
      { wch: 12 }, // Lat
      { wch: 12 }, // Lon
      { wch: 8 },  // Working
      { wch: 14 }, // Source
      { wch: 30 }, // Conflict
      { wch: 50 }, // Booking source
      { wch: 8 },  // Auto GPS
      { wch: 10 }, // Auto booking
      { wch: 12 }, // GPS confirmed
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Locations');
  }

  // Stats summary sheet
  const countryCounts = {};
  let ukNights = 0;
  let italyDays = 0;
  let workDays = 0;
  Object.values(locations).forEach(e => {
    const c = e.country || 'Unrecorded';
    countryCounts[c] = (countryCounts[c] || 0) + 1;
    if (e.country === 'UK') ukNights++;
    if (e.country === 'Italy') italyDays++;
    if (e.working) workDays++;
  });
  const statsRows = Object.entries(countryCounts)
    .sort((a, b) => b[1] - a[1])
    .map(([country, count]) => ({ 'Country': country, 'Nights': count }));
  statsRows.push({});
  statsRows.push({ 'Country': 'UK nights (max 90)', 'Nights': ukNights });
  statsRows.push({ 'Country': 'Italy days (target 183+)', 'Nights': italyDays });
  statsRows.push({ 'Country': 'UK work days (max 30)', 'Nights': workDays });
  statsRows.push({ 'Country': 'Total entries', 'Nights': entryCount });
  statsRows.push({});
  statsRows.push({ 'Country': 'Backup date', 'Nights': today });

  const statsWs = XLSX.utils.json_to_sheet(statsRows);
  statsWs['!cols'] = [{ wch: 25 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, statsWs, 'Summary');

  // Write XLSX
  XLSX.writeFile(wb, path.join(BACKUP_DIR, 'latest.xlsx'));
  console.log('Wrote latest.xlsx');

  // 4. Clean up old history files (keep only MAX_HISTORY most recent)
  const historyFiles = fs.readdirSync(HISTORY_DIR)
    .filter(f => f.startsWith('backup-') && f.endsWith('.json'))
    .sort()
    .reverse();

  if (historyFiles.length > MAX_HISTORY) {
    const toDelete = historyFiles.slice(MAX_HISTORY);
    toDelete.forEach(f => {
      fs.unlinkSync(path.join(HISTORY_DIR, f));
      console.log(`Deleted old backup: ${f}`);
    });
  }

  console.log(`\nBackup complete. ${historyFiles.length > MAX_HISTORY ? historyFiles.length - MAX_HISTORY : 0} old backups cleaned up.`);
  console.log(`History: ${Math.min(historyFiles.length, MAX_HISTORY)} backups retained.`);

  process.exit(0);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
