/**
 * Booking Sync Script
 *
 * Reads Outlook calendar events and emails from Hotels/Bookings folder
 * via Microsoft Graph API, extracts flight and hotel details, and
 * updates Firebase Realtime Database.
 *
 * Triggered by GitHub Actions on a schedule (every 6 hours).
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
const DAYS_AHEAD = 90;  // Look 90 days ahead for calendar events
const DAYS_BACK = 7;    // Look 7 days back for recent emails
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

// ===== DATE HELPERS =====
function fmtDate(d) {
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
}

function parseDate(str) {
  if (!str) return null;
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

// Get all dates between two dates (inclusive)
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

// Extract flight numbers from text (e.g., BA123, EK456, LH1234)
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

// Extract airport codes from text
function extractAirports(text) {
  if (!text) return [];
  const found = [];
  for (const code of Object.keys(AIRPORTS)) {
    // Look for the 3-letter code as a standalone word
    const pattern = new RegExp('\\b' + code + '\\b', 'g');
    if (pattern.test(text)) found.push(code);
  }
  return found;
}

// Determine destination from flight context
// e.g., "LHR to MXP" → destination is MXP
function extractDestination(text) {
  if (!text) return null;
  // Patterns: "X to Y", "X → Y", "X - Y", "X>Y", "departing X arriving Y"
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
  // Fall back to last airport mentioned
  const airports = extractAirports(text.toUpperCase());
  if (airports.length >= 2) return AIRPORTS[airports[airports.length - 1]];
  if (airports.length === 1) return AIRPORTS[airports[0]];
  return null;
}

// Extract hotel name from text
function extractHotelName(text) {
  if (!text) return null;
  // Common patterns in hotel confirmation subjects/bodies
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

// Extract city from text
function extractCity(text) {
  if (!text) return null;
  // Look for "in <City>" or "<City>, <Country>" patterns
  const cityPattern = /\bin\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)/;
  const m = cityPattern.exec(text);
  if (m) return m[1];
  return null;
}

// ===== CALENDAR PROCESSING =====

async function processCalendar(token) {
  const now = new Date();
  const startDate = new Date(now);
  startDate.setDate(startDate.getDate() - 3); // Include a few days back
  const endDate = new Date(now);
  endDate.setDate(endDate.getDate() + DAYS_AHEAD);

  const start = startDate.toISOString();
  const end = endDate.toISOString();

  console.log(`Reading calendar events from ${fmtDate(startDate)} to ${fmtDate(endDate)}...`);

  const path = `/users/${USER_EMAIL}/calendarview?startDateTime=${start}&endDateTime=${end}&$top=100&$select=subject,bodyPreview,start,end,location,categories`;
  const result = await graphGet(token, path);

  if (result.error) {
    console.error('Calendar error:', result.error.message);
    return [];
  }

  const events = result.value || [];
  console.log(`Found ${events.length} calendar events`);

  const bookings = [];

  for (const event of events) {
    const subject = event.subject || '';
    const body = event.bodyPreview || '';
    const location = event.location?.displayName || '';
    const allText = subject + ' ' + body + ' ' + location;
    const allTextUpper = allText.toUpperCase();

    // Skip car rentals
    if (/\b(car\s*rental|hertz|avis|europcar|sixt|enterprise|rent.?a.?car|pick.?up.*drop.?off|vehicle\s*collect)/i.test(allText)) continue;

    // Skip events that don't look like travel
    const isFlight = /\b(flight|fly|depart|arrive|airport|boarding|BA\d|EK\d|LH\d|AF\d|AZ\d|FR\d|U2\d|QR\d|EY\d|SQ\d|CX\d|TK\d)/i.test(allText);
    const isHotel = /\b(hotel|check.?in|check.?out|booking|reservation|stay|accommodation|airbnb)/i.test(allText);

    if (!isFlight && !isHotel) continue;

    const startDate = parseDate(event.start?.dateTime || event.start?.date);
    const endDate = parseDate(event.end?.dateTime || event.end?.date);
    if (!startDate) continue;

    if (isFlight) {
      const flights = extractFlights(allTextUpper);
      const dest = extractDestination(allTextUpper);
      bookings.push({
        type: 'flight',
        date: fmtDate(startDate),
        flights: flights.join(', '),
        city: dest?.city || extractCity(allText) || '',
        country: dest?.country || '',
        place: '',
        source: 'calendar',
        raw: subject
      });
    }

    if (isHotel && endDate) {
      const hotelName = extractHotelName(allText) || '';
      const nights = dateRange(startDate, new Date(endDate.getTime() - 86400000)); // Exclude checkout day

      for (const dateStr of nights) {
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
    }
  }

  return bookings;
}

// ===== EMAIL PROCESSING =====

async function findFolder(token, folderName) {
  // Search mail folders (including nested ones)
  const result = await graphGet(token, `/users/${USER_EMAIL}/mailFolders?$top=50`);
  if (result.error) {
    console.error('Folder error:', result.error.message);
    return null;
  }

  const topLevel = (result.value || []).map(f => f.displayName);
  console.log('  Top-level folders:', topLevel.join(', '));

  for (const folder of (result.value || [])) {
    if (folder.displayName === folderName) return folder.id;

    // Check child folders
    const children = await graphGet(token, `/users/${USER_EMAIL}/mailFolders/${folder.id}/childFolders?$top=50`);
    const childNames = (children.value || []).map(c => c.displayName);
    if (childNames.length) console.log('  ' + folder.displayName + ' → children:', childNames.join(', '));
    for (const child of (children.value || [])) {
      if (child.displayName === folderName) return child.id;
    }
  }
  console.log('  Folder "' + folderName + '" not found');
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

async function processEmails(token) {
  const allBookings = [];

  // Process hotel folders
  console.log('Looking for hotel email folder...');
  const hotelFolderId = await findFirstFolder(token, HOTEL_FOLDERS);
  if (hotelFolderId) {
    const hotelBookings = await processEmailsFromFolder(token, hotelFolderId, 'hotel');
    allBookings.push(...hotelBookings);
  } else {
    console.log('No hotel folder found.');
  }

  // Process flight folders
  console.log('Looking for flights email folder...');
  const flightFolderId = await findFirstFolder(token, FLIGHT_FOLDERS);
  if (flightFolderId) {
    const flightBookings = await processEmailsFromFolder(token, flightFolderId, 'flight');
    allBookings.push(...flightBookings);
  } else {
    console.log('No flights folder found.');
  }

  if (!allBookings.length) {
    console.log('No booking folders found. Skipping email processing.');
  }

  return allBookings;
}

async function processEmailsFromFolder(token, folderId, folderType) {
  const since = new Date();
  since.setDate(since.getDate() - DAYS_BACK);

  const filter = `receivedDateTime ge ${since.toISOString()}`;
  const path = `/users/${USER_EMAIL}/mailFolders/${folderId}/messages?$filter=${encodeURIComponent(filter)}&$top=50&$select=subject,bodyPreview,receivedDateTime,from`;
  const result = await graphGet(token, path);

  if (result.error) {
    console.error('Email error:', result.error.message);
    return [];
  }

  const messages = result.value || [];
  console.log(`Found ${messages.length} recent emails in ${folderType} folder`);

  const bookings = [];

  for (const msg of messages) {
    const subject = msg.subject || '';
    const body = msg.bodyPreview || '';
    const from = msg.from?.emailAddress?.address || '';
    const allText = subject + ' ' + body;
    const allTextUpper = allText.toUpperCase();

    // Skip car rental emails
    if (/\b(car\s*rental|hertz|avis|europcar|sixt|enterprise|rent.?a.?car|pick.?up.*drop.?off|vehicle\s*collect)/i.test(allText)) {
      console.log(`  🚗 SKIP car rental: ${subject.slice(0, 60)}`);
      continue;
    }

    // Use folder type as hint — emails in Hotels folder are hotels, Flights folder are flights
    const isFlight = folderType === 'flight' ||
                     /\b(flight|itinerary|boarding|e-?ticket|airline)/i.test(allText) ||
                     extractFlights(allTextUpper).length > 0;
    const isHotel = folderType === 'hotel' ||
                    /\b(hotel|reservation|check.?in|booking|stay|accommodation|airbnb|nights?)/i.test(allText);

    // Try to extract dates from the email
    // Common patterns: "Check-in: 15 March 2026", "Date: 2026-03-15", "March 15, 2026"
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

    // Sort dates and take first as check-in, last as check-out
    extractedDates.sort((a, b) => a - b);

    if (isFlight && extractedDates.length > 0) {
      const flights = extractFlights(allTextUpper);
      const dest = extractDestination(allTextUpper);
      bookings.push({
        type: 'flight',
        date: fmtDate(extractedDates[0]),
        flights: flights.join(', '),
        city: dest?.city || '',
        country: dest?.country || '',
        place: '',
        source: 'email',
        raw: subject
      });
    }

    if (isHotel && extractedDates.length >= 2) {
      const hotelName = extractHotelName(allText) || '';
      const checkIn = extractedDates[0];
      const checkOut = extractedDates[extractedDates.length - 1];
      const nights = dateRange(checkIn, new Date(checkOut.getTime() - 86400000));

      for (const dateStr of nights) {
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
    }
  }

  return bookings;
}

// ===== FIREBASE UPDATE =====

async function updateFirebase(bookings) {
  if (!bookings.length) {
    console.log('No bookings to update.');
    return;
  }

  console.log(`\nProcessing ${bookings.length} booking entries...`);

  // Read existing entries
  const snapshot = await db.ref('locations').once('value');
  const existing = snapshot.val() || {};

  const updates = {};
  let newCount = 0;
  let mergedCount = 0;
  let skippedCount = 0;

  for (const booking of bookings) {
    const dateStr = booking.date;
    const current = existing[dateStr];

    // Don't overwrite manually-set entries
    if (current && current.city && !current.autoGps && !current.autoBooking) {
      skippedCount++;
      continue;
    }

    // Merge: existing data takes priority, booking data fills gaps
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

    // Preserve existing fields
    if (current?.lat) entry.lat = current.lat;
    if (current?.lon) entry.lon = current.lon;
    if (current?.working) entry.working = current.working;
    if (current?.autoGps) entry.autoGps = current.autoGps;

    // Check for country conflict — GPS vs booking disagree on country
    if (current?.autoGps && current?.country && booking.country &&
        current.country !== booking.country) {
      entry.countryConflict = 'GPS says ' + current.country + ', booking says ' + booking.country;
      console.log(`  ⚠ COUNTRY CONFLICT on ${dateStr}: GPS=${current.country}, Booking=${booking.country}`);
    }

    // Only update if we're adding new information
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
    console.log(`Updated Firebase: ${newCount} new, ${mergedCount} merged, ${skippedCount} skipped`);
  } else {
    console.log(`No updates needed (${skippedCount} skipped)`);
  }

  // Log a sync timestamp
  await db.ref('settings/lastBookingSync').set(new Date().toISOString());
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
  console.log('=== Midnight Tracker — Booking Sync ===');
  console.log('Time:', new Date().toISOString());

  // Get Graph API token
  console.log('\nAuthenticating with Microsoft Graph...');
  const token = await getGraphToken();
  console.log('Authenticated.');

  // Process calendar
  const calendarBookings = await processCalendar(token);
  console.log(`Calendar: ${calendarBookings.length} booking entries found`);

  // Process emails
  const emailBookings = await processEmails(token);
  console.log(`Email: ${emailBookings.length} booking entries found`);

  // Combine (calendar takes priority for same date)
  const allBookings = [...calendarBookings, ...emailBookings];

  // Deduplicate by date (keep first occurrence, which is calendar)
  const seen = new Set();
  const deduplicated = [];
  for (const b of allBookings) {
    const key = b.date + '|' + b.type;
    if (!seen.has(key)) {
      seen.add(key);
      deduplicated.push(b);
    } else {
      // Merge flight numbers if same date
      const existing = deduplicated.find(d => d.date === b.date);
      if (existing && b.flights) {
        existing.flights = mergeFlights(existing.flights, b.flights);
      }
    }
  }

  console.log(`\nTotal unique bookings: ${deduplicated.length}`);
  deduplicated.forEach(b => {
    console.log(`  ${b.date} | ${b.type} | ${b.place || b.flights || '-'} | ${b.city} | ${b.source}`);
  });

  // Update Firebase
  await updateFirebase(deduplicated);

  console.log('\nDone.');
  process.exit(0);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
