/**
 * ONE-OFF Backfill Script
 *
 * Scans Outlook calendar and Hotels/Flights email folders back to 6 April 2025
 * (start of current UK tax year) to populate historical location data.
 *
 * Run manually once via GitHub Actions, then delete or ignore.
 *
 * Environment variables required:
 *   FIREBASE_SERVICE_ACCOUNT - Firebase service account JSON
 *   MS_TENANT_ID            - Azure AD tenant ID
 *   MS_CLIENT_ID            - Azure AD app client ID
 *   MS_CLIENT_SECRET        - Azure AD app client secret
 *   MS_USER_EMAIL           - Outlook mailbox to read (steveb@canapii.com)
 */

const admin = require('firebase-admin');
const https = require('https');

// ===== CONFIG =====
const USER_EMAIL = process.env.MS_USER_EMAIL || 'steveb@canapii.com';
const TAX_YEAR_START = '2025-04-06';  // UK tax year start
const HOTEL_FOLDERS = ['Hotels', 'Hotel'];
const FLIGHT_FOLDERS = ['Flights', 'Flight'];

// Known airports → city/country mapping
const AIRPORTS = {
  'LHR': { city: 'London', country: 'UK' }, 'LGW': { city: 'London', country: 'UK' },
  'STN': { city: 'London', country: 'UK' }, 'LTN': { city: 'London', country: 'UK' },
  'LCY': { city: 'London', country: 'UK' }, 'MXP': { city: 'Milan', country: 'Italy' },
  'FCO': { city: 'Rome', country: 'Italy' }, 'BKK': { city: 'Bangkok', country: 'Thailand' },
  'DMK': { city: 'Bangkok', country: 'Thailand' }, 'CDG': { city: 'Paris', country: 'France' },
  'ORY': { city: 'Paris', country: 'France' }, 'AMS': { city: 'Amsterdam', country: 'Netherlands' },
  'FRA': { city: 'Frankfurt', country: 'Germany' }, 'MUC': { city: 'Munich', country: 'Germany' },
  'BCN': { city: 'Barcelona', country: 'Spain' }, 'MAD': { city: 'Madrid', country: 'Spain' },
  'ZRH': { city: 'Zurich', country: 'Switzerland' }, 'GVA': { city: 'Geneva', country: 'Switzerland' },
  'IST': { city: 'Istanbul', country: 'Turkey' }, 'DXB': { city: 'Dubai', country: 'UAE' },
  'SIN': { city: 'Singapore', country: 'Singapore' }, 'HKG': { city: 'Hong Kong', country: 'Hong Kong' },
  'NRT': { city: 'Tokyo', country: 'Japan' }, 'HND': { city: 'Tokyo', country: 'Japan' },
  'ICN': { city: 'Seoul', country: 'South Korea' }, 'TPE': { city: 'Taipei', country: 'Taiwan' },
  'DEL': { city: 'Delhi', country: 'India' }, 'BOM': { city: 'Mumbai', country: 'India' },
  'JFK': { city: 'New York', country: 'USA' }, 'LAX': { city: 'Los Angeles', country: 'USA' },
  'SFO': { city: 'San Francisco', country: 'USA' }, 'ORD': { city: 'Chicago', country: 'USA' },
  'SYD': { city: 'Sydney', country: 'Australia' }, 'MEL': { city: 'Melbourne', country: 'Australia' },
  'YYZ': { city: 'Toronto', country: 'Canada' }, 'LIS': { city: 'Lisbon', country: 'Portugal' },
  'ATH': { city: 'Athens', country: 'Greece' }, 'VCE': { city: 'Venice', country: 'Italy' },
  'NAP': { city: 'Naples', country: 'Italy' }, 'BGY': { city: 'Milan', country: 'Italy' },
  'LIN': { city: 'Milan', country: 'Italy' }, 'PMO': { city: 'Palermo', country: 'Italy' },
  'CTA': { city: 'Catania', country: 'Italy' }, 'BHX': { city: 'Birmingham', country: 'UK' },
  'MAN': { city: 'Manchester', country: 'UK' }, 'EDI': { city: 'Edinburgh', country: 'UK' },
  'OXF': { city: 'Oxford', country: 'UK' },
};

// ===== FIREBASE INIT =====
const serviceAccount = JSON.parse(process.env.FIREBASE_SERVICE_ACCOUNT);
admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  databaseURL: 'https://midnight-tracker-steve-default-rtdb.europe-west1.firebasedatabase.app'
});
const db = admin.database();

// ===== MICROSOFT GRAPH API =====

async function getGraphToken() {
  const tenantId = process.env.MS_TENANT_ID;
  const clientId = process.env.MS_CLIENT_ID;
  const clientSecret = process.env.MS_CLIENT_SECRET;

  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default'
  }).toString();

  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': body.length }
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        const json = JSON.parse(data);
        if (json.access_token) resolve(json.access_token);
        else reject(new Error('Token error: ' + JSON.stringify(json)));
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function graphGet(token, path) {
  return new Promise((resolve, reject) => {
    const url = new URL('https://graph.microsoft.com/v1.0' + path);
    const req = https.request({
      hostname: url.hostname,
      path: url.pathname + url.search,
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch(e) { reject(new Error('Parse error: ' + data.slice(0, 200))); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// Paginated graph GET — follows @odata.nextLink to get all results
async function graphGetAll(token, path) {
  const allValues = [];
  let currentPath = path;
  let page = 1;
  while (currentPath) {
    const result = currentPath.startsWith('https://')
      ? await graphGetFullUrl(token, currentPath)
      : await graphGet(token, currentPath);
    if (result.error) {
      console.error('Graph error:', result.error.message);
      break;
    }
    const values = result.value || [];
    allValues.push(...values);
    console.log(`  Page ${page}: ${values.length} items (total: ${allValues.length})`);
    currentPath = result['@odata.nextLink'] || null;
    page++;
  }
  return allValues;
}

async function graphGetFullUrl(token, fullUrl) {
  return new Promise((resolve, reject) => {
    const url = new URL(fullUrl);
    const req = https.request({
      hostname: url.hostname,
      path: url.pathname + url.search,
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch(e) { reject(new Error('Parse error: ' + data.slice(0, 200))); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// ===== DATE HELPERS =====
function fmtDate(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function parseDate(str) {
  if (!str) return null;
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

function dateRange(start, end) {
  const dates = [];
  const d = new Date(start);
  const e = new Date(end);
  while (d <= e) {
    dates.push(fmtDate(d));
    d.setDate(d.getDate() + 1);
  }
  return dates;
}

// ===== PARSING =====

function extractFlights(text) {
  if (!text) return [];
  const pattern = /\b([A-Z]{2})\s*(\d{1,4})\b/g;
  const flights = [];
  let m;
  while ((m = pattern.exec(text)) !== null) {
    flights.push(m[1] + m[2]);
  }
  return [...new Set(flights)];
}

function extractAirports(text) {
  if (!text) return [];
  const found = [];
  for (const code of Object.keys(AIRPORTS)) {
    const pattern = new RegExp('\\b' + code + '\\b', 'g');
    if (pattern.test(text)) found.push(code);
  }
  return found;
}

function extractDestination(text) {
  if (!text) return null;
  const patterns = [
    /\b([A-Z]{3})\s*(?:to|→|->|>|–|—)\s*([A-Z]{3})\b/gi,
    /(?:arriving|arr\.?|destination)\s*:?\s*([A-Z]{3})\b/gi,
  ];
  for (const pat of patterns) {
    const m = pat.exec(text);
    if (m) {
      const dest = m[2] || m[1];
      if (AIRPORTS[dest.toUpperCase()]) return AIRPORTS[dest.toUpperCase()];
    }
  }
  const airports = extractAirports(text.toUpperCase());
  if (airports.length >= 2) return AIRPORTS[airports[airports.length - 1]];
  if (airports.length === 1) return AIRPORTS[airports[0]];
  return null;
}

function extractHotelName(text) {
  if (!text) return null;
  const patterns = [
    /(?:booking|reservation|confirmation)\s+(?:at|for)\s+(.+?)(?:\s*[-–|,]|\s+in\s+|\s+on\s+|$)/i,
    /(?:hotel|resort|inn|lodge|hostel|apartment|residence|suites?)\s*:?\s*(.+?)(?:\s*[-–|,]|$)/i,
    /(?:your stay at|check.?in at|welcome to)\s+(.+?)(?:\s*[-–|,]|\s+on\s+|$)/i,
  ];
  for (const pat of patterns) {
    const m = pat.exec(text);
    if (m && m[1].trim().length > 2 && m[1].trim().length < 80) {
      return m[1].trim();
    }
  }
  return null;
}

function extractCity(text) {
  if (!text) return null;
  const cityPattern = /\bin\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)/;
  const m = cityPattern.exec(text);
  if (m) return m[1];
  return null;
}

// ===== FOLDER SEARCH =====

async function findFolder(token, folderName) {
  const result = await graphGet(token, `/users/${USER_EMAIL}/mailFolders?$top=50`);
  if (result.error) {
    console.error('Folder error:', result.error.message);
    return null;
  }

  for (const folder of (result.value || [])) {
    if (folder.displayName === folderName) return folder.id;
    const children = await graphGet(token, `/users/${USER_EMAIL}/mailFolders/${folder.id}/childFolders?$top=50`);
    for (const child of (children.value || [])) {
      if (child.displayName === folderName) return child.id;
    }
  }
  return null;
}

async function findFirstFolder(token, candidates) {
  for (const name of candidates) {
    const id = await findFolder(token, name);
    if (id) {
      console.log(`Found folder: "${name}"`);
      return id;
    }
  }
  return null;
}

// ===== CALENDAR PROCESSING (BACKFILL) =====

async function processCalendar(token) {
  const startDate = new Date(TAX_YEAR_START + 'T00:00:00Z');
  const endDate = new Date(); // up to today

  console.log(`\nReading calendar from ${fmtDate(startDate)} to ${fmtDate(endDate)}...`);

  const path = `/users/${USER_EMAIL}/calendarview?startDateTime=${startDate.toISOString()}&endDateTime=${endDate.toISOString()}&$top=250&$select=subject,bodyPreview,start,end,location,categories`;
  const events = await graphGetAll(token, path);
  console.log(`Total calendar events: ${events.length}`);

  const bookings = [];

  for (const event of events) {
    const subject = event.subject || '';
    const body = event.bodyPreview || '';
    const location = event.location?.displayName || '';
    const allText = subject + ' ' + body + ' ' + location;
    const allTextUpper = allText.toUpperCase();

    // Skip car rentals in calendar too
    if (/\b(car\s*rental|hertz|avis|europcar|sixt|enterprise|rent.?a.?car|pick.?up.*drop.?off|vehicle\s*collect)/i.test(allText)) continue;

    const isFlight = /\b(flight|fly|depart|arrive|airport|boarding|BA\d|EK\d|LH\d|AF\d|AZ\d|FR\d|U2\d|QR\d|EY\d|SQ\d|CX\d|TK\d)/i.test(allText);
    const isHotel = /\b(hotel|check.?in|check.?out|booking|reservation|stay|accommodation|airbnb)/i.test(allText);

    if (!isFlight && !isHotel) continue;

    const evStart = parseDate(event.start?.dateTime || event.start?.date);
    const evEnd = parseDate(event.end?.dateTime || event.end?.date);
    if (!evStart) continue;

    // Only include dates within tax year
    if (fmtDate(evStart) < TAX_YEAR_START) continue;

    if (isFlight) {
      const flights = extractFlights(allTextUpper);
      const dest = extractDestination(allTextUpper);
      bookings.push({
        type: 'flight',
        date: fmtDate(evStart),
        flights: flights.join(', '),
        city: dest?.city || extractCity(allText) || '',
        country: dest?.country || '',
        place: '',
        source: 'calendar',
        raw: subject
      });
      console.log(`  ✈ ${fmtDate(evStart)} | ${flights.join(', ') || '?'} | ${dest?.city || '?'} | ${subject.slice(0, 60)}`);
    }

    if (isHotel && evEnd) {
      const hotelName = extractHotelName(allText) || '';
      const nights = dateRange(evStart, new Date(evEnd.getTime() - 86400000));
      const nightsInTaxYear = nights.filter(d => d >= TAX_YEAR_START);

      for (const dateStr of nightsInTaxYear) {
        bookings.push({
          type: 'hotel',
          date: dateStr,
          flights: '',
          city: extractCity(allText) || location || '',
          country: '',
          place: hotelName,
          source: 'calendar',
          raw: subject
        });
      }
      if (nightsInTaxYear.length) {
        console.log(`  🏨 ${nightsInTaxYear[0]}→${nightsInTaxYear[nightsInTaxYear.length-1]} | ${hotelName.slice(0, 40)} | ${nightsInTaxYear.length} nights`);
      }
    }
  }

  return bookings;
}

// ===== EMAIL PROCESSING (BACKFILL) =====

async function processEmails(token) {
  const allBookings = [];

  console.log('\nLooking for hotel email folder...');
  const hotelFolderId = await findFirstFolder(token, HOTEL_FOLDERS);
  if (hotelFolderId) {
    const hotelBookings = await processEmailsFromFolder(token, hotelFolderId, 'hotel');
    allBookings.push(...hotelBookings);
  } else {
    console.log('No hotel folder found.');
  }

  console.log('\nLooking for flights email folder...');
  const flightFolderId = await findFirstFolder(token, FLIGHT_FOLDERS);
  if (flightFolderId) {
    const flightBookings = await processEmailsFromFolder(token, flightFolderId, 'flight');
    allBookings.push(...flightBookings);
  } else {
    console.log('No flights folder found.');
  }

  return allBookings;
}

async function processEmailsFromFolder(token, folderId, folderType) {
  const since = new Date(TAX_YEAR_START + 'T00:00:00Z');

  const filter = `receivedDateTime ge ${since.toISOString()}`;
  const path = `/users/${USER_EMAIL}/mailFolders/${folderId}/messages?$filter=${encodeURIComponent(filter)}&$top=100&$select=subject,bodyPreview,receivedDateTime,from&$orderby=receivedDateTime desc`;
  const messages = await graphGetAll(token, path);
  console.log(`Total emails in ${folderType} folder since ${TAX_YEAR_START}: ${messages.length}`);

  const bookings = [];

  for (const msg of messages) {
    const subject = msg.subject || '';
    const body = msg.bodyPreview || '';
    const allText = subject + ' ' + body;
    const allTextUpper = allText.toUpperCase();

    // Skip car rental emails
    if (/\b(car\s*rental|hertz|avis|europcar|sixt|enterprise|rent.?a.?car|pick.?up.*drop.?off|vehicle\s*collect)/i.test(allText)) {
      console.log(`  🚗 SKIP car rental: ${subject.slice(0, 60)}`);
      continue;
    }

    const isFlight = folderType === 'flight' ||
                     /\b(flight|itinerary|boarding|e-?ticket|airline)/i.test(allText) ||
                     extractFlights(allTextUpper).length > 0;
    const isHotel = folderType === 'hotel' ||
                    /\b(hotel|reservation|check.?in|booking|stay|accommodation|airbnb|nights?)/i.test(allText);

    const datePatterns = [
      /(\d{4}-\d{2}-\d{2})/g,
      /(\d{1,2})\s+(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{4})/gi,
      /(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2}),?\s+(\d{4})/gi,
    ];

    const extractedDates = [];
    for (const pat of datePatterns) {
      let m;
      while ((m = pat.exec(allText)) !== null) {
        const d = parseDate(m[0]);
        if (d && d.getFullYear() >= 2025 && d.getFullYear() <= 2028) {
          extractedDates.push(d);
        }
      }
    }

    extractedDates.sort((a, b) => a - b);
    // Filter to tax year
    const validDates = extractedDates.filter(d => fmtDate(d) >= TAX_YEAR_START);

    if (isFlight && validDates.length > 0) {
      const flights = extractFlights(allTextUpper);
      const dest = extractDestination(allTextUpper);
      bookings.push({
        type: 'flight',
        date: fmtDate(validDates[0]),
        flights: flights.join(', '),
        city: dest?.city || '',
        country: dest?.country || '',
        place: '',
        source: 'email',
        raw: subject
      });
      console.log(`  ✈ ${fmtDate(validDates[0])} | ${flights.join(', ') || '?'} | ${dest?.city || '?'} | ${subject.slice(0, 60)}`);
    }

    if (isHotel && validDates.length > 0) {
      const hotelName = extractHotelName(allText) || '';
      const checkIn = validDates[0];
      const checkOut = validDates.length > 1 ? validDates[validDates.length - 1] : new Date(checkIn.getTime() + 86400000);
      const nights = dateRange(checkIn, new Date(checkOut.getTime() - 86400000));
      const nightsInTaxYear = nights.filter(d => d >= TAX_YEAR_START);

      for (const dateStr of nightsInTaxYear) {
        bookings.push({
          type: 'hotel',
          date: dateStr,
          flights: '',
          city: extractCity(allText) || '',
          country: '',
          place: hotelName,
          source: 'email',
          raw: subject
        });
      }
      if (nightsInTaxYear.length) {
        console.log(`  🏨 ${nightsInTaxYear[0]}→${nightsInTaxYear[nightsInTaxYear.length-1]} | ${hotelName.slice(0, 40)} | ${nightsInTaxYear.length} nights`);
      }
    }
  }

  return bookings;
}

// ===== FIREBASE UPDATE =====

async function updateFirebase(bookings) {
  if (!bookings.length) {
    console.log('\nNo bookings to update.');
    return;
  }

  console.log(`\nProcessing ${bookings.length} booking entries...`);

  const snapshot = await db.ref('locations').once('value');
  const existing = snapshot.val() || {};

  const updates = {};
  let newCount = 0;
  let mergedCount = 0;
  let skippedCount = 0;

  for (const booking of bookings) {
    const dateStr = booking.date;
    const current = existing[dateStr];

    // Don't overwrite manually-set entries (no autoGps, no autoBooking = manual)
    if (current && current.city && !current.autoGps && !current.autoBooking) {
      skippedCount++;
      continue;
    }

    // Build booking source audit trail
    const sourceInfo = booking.source + ': ' + (booking.raw || '').slice(0, 120);
    const existingSource = current?.bookingSource || '';
    const combinedSource = existingSource
      ? (existingSource.includes(sourceInfo) ? existingSource : existingSource + ' | ' + sourceInfo)
      : sourceInfo;

    const entry = {
      place: current?.place || booking.place || '',
      city: current?.city || booking.city || '',
      country: current?.country || booking.country || '',
      flights: mergeFlights(current?.flights, booking.flights),
      notes: current?.notes || '',
      autoBooking: true,
      bookingSource: combinedSource
    };

    if (current?.lat) entry.lat = current.lat;
    if (current?.lon) entry.lon = current.lon;
    if (current?.working) entry.working = current.working;
    if (current?.autoGps) entry.autoGps = current.autoGps;

    const hasNew = (!current) ||
                   (!current.place && entry.place) ||
                   (!current.city && entry.city) ||
                   (!current.flights && entry.flights) ||
                   (entry.flights && entry.flights !== current.flights) ||
                   (!current.bookingSource && entry.bookingSource);

    if (hasNew) {
      updates['locations/' + dateStr] = entry;
      if (!current) newCount++;
      else mergedCount++;
    } else {
      skippedCount++;
    }
  }

  if (Object.keys(updates).length > 0) {
    await db.ref().update(updates);
    console.log(`\nUpdated Firebase: ${newCount} new, ${mergedCount} merged, ${skippedCount} skipped`);
  } else {
    console.log(`\nNo updates needed (${skippedCount} skipped)`);
  }

  await db.ref('settings/lastBackfill').set(new Date().toISOString());
}

function mergeFlights(existing, newFlights) {
  if (!existing && !newFlights) return '';
  if (!existing) return newFlights;
  if (!newFlights) return existing;
  const all = new Set([...existing.split(/[,\s]+/), ...newFlights.split(/[,\s]+/)].filter(Boolean));
  return [...all].join(', ');
}

// ===== MAIN =====

async function main() {
  console.log('=== Midnight Tracker — BACKFILL ===');
  console.log('Tax year start:', TAX_YEAR_START);
  console.log('Time:', new Date().toISOString());

  console.log('\nAuthenticating with Microsoft Graph...');
  const token = await getGraphToken();
  console.log('Authenticated.');

  // Process calendar
  const calendarBookings = await processCalendar(token);
  console.log(`\nCalendar: ${calendarBookings.length} booking entries found`);

  // Process emails
  const emailBookings = await processEmails(token);
  console.log(`\nEmail: ${emailBookings.length} booking entries found`);

  // Combine — calendar takes priority
  const allBookings = [...calendarBookings, ...emailBookings];

  // Deduplicate by date+type
  const seen = new Set();
  const deduplicated = [];
  for (const b of allBookings) {
    const key = b.date + '|' + b.type;
    if (!seen.has(key)) {
      seen.add(key);
      deduplicated.push(b);
    } else {
      const existing = deduplicated.find(d => d.date === b.date);
      if (existing && b.flights) {
        existing.flights = mergeFlights(existing.flights, b.flights);
      }
    }
  }

  console.log(`\n=== SUMMARY ===`);
  console.log(`Total unique entries: ${deduplicated.length}`);
  console.log('\nAll entries:');
  deduplicated.sort((a, b) => a.date.localeCompare(b.date));
  deduplicated.forEach(b => {
    console.log(`  ${b.date} | ${b.type.padEnd(6)} | ${(b.place || b.flights || '-').slice(0, 30).padEnd(30)} | ${(b.city || '?').padEnd(15)} | ${b.country || '?'} | ${b.source}`);
  });

  // Update Firebase
  await updateFirebase(deduplicated);

  console.log('\nBackfill complete.');
  process.exit(0);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
