/**
 * Booking Sync Script
 *
 * Reads calendar events and emails from two sources:
 *   1. Microsoft Outlook (steveb@canapii.com) — via Graph API (client credentials)
 *   2. Google (swbrazier@gmail.com) — via Calendar & Gmail APIs (OAuth2 refresh token)
 *
 * Extracts flight and hotel details, deduplicates, and updates Firebase.
 *
 * Triggered by GitHub Actions on a schedule (every 6 hours).
 *
 * Environment variables required:
 *   FIREBASE_SERVICE_ACCOUNT  - Firebase service account JSON
 *   MS_TENANT_ID              - Azure AD tenant ID
 *   MS_CLIENT_ID              - Azure AD app client ID
 *   MS_CLIENT_SECRET          - Azure AD app client secret
 *   MS_USER_EMAIL             - Outlook mailbox to read (steveb@canapii.com)
 *   GOOGLE_CLIENT_ID          - Google OAuth2 client ID (optional, skipped if missing)
 *   GOOGLE_CLIENT_SECRET      - Google OAuth2 client secret
 *   GOOGLE_REFRESH_TOKEN      - Google OAuth2 refresh token (from one-time consent)
 */

const SYNC_VERSION = '1.5.0';

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

// Normalise country names so variants map to canonical short forms
function normalizeCountry(c) {
  if (!c) return c;
  const map = {
    'United States': 'USA',
    'United States of America': 'USA',
    'United Kingdom': 'UK',
    'Great Britain': 'UK',
    'United Arab Emirates': 'UAE',
    'Republic of China': 'Taiwan',
    'Korea, Republic of': 'South Korea',
    'Republic of Korea': 'South Korea',
  };
  return map[c] || c;
}

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

// ===== GOOGLE API =====

const GOOGLE_ENABLED = !!(process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET && process.env.GOOGLE_REFRESH_TOKEN);

async function getGoogleToken() {
  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    client_id: process.env.GOOGLE_CLIENT_ID,
    client_secret: process.env.GOOGLE_CLIENT_SECRET,
    refresh_token: process.env.GOOGLE_REFRESH_TOKEN
  }).toString();

  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'oauth2.googleapis.com',
      path: '/token',
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded', 'Content-Length': Buffer.byteLength(body) }
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try {
          const json = JSON.parse(data);
          if (json.access_token) resolve(json.access_token);
          else reject(new Error('Google token error: ' + JSON.stringify(json)));
        } catch(e) { reject(new Error('Google token parse error: ' + data.slice(0, 200))); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function googleGet(token, url) {
  const u = new URL(url);
  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: u.hostname,
      path: u.pathname + u.search,
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/json' }
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        try { resolve(JSON.parse(data)); }
        catch(e) { reject(new Error('Google parse error: ' + data.slice(0, 200))); }
      });
    });
    req.on('error', reject);
    req.end();
  });
}

// ===== GOOGLE CALENDAR =====

async function processGoogleCalendar(token) {
  const now = new Date();
  const startDate = new Date(now);
  startDate.setDate(startDate.getDate() - 3);
  const endDate = new Date(now);
  endDate.setDate(endDate.getDate() + DAYS_AHEAD);

  console.log(`Reading Google Calendar events from ${fmtDate(startDate)} to ${fmtDate(endDate)}...`);

  const url = `https://www.googleapis.com/calendar/v3/calendars/primary/events?timeMin=${startDate.toISOString()}&timeMax=${endDate.toISOString()}&maxResults=250&singleEvents=true&orderBy=startTime`;
  const result = await googleGet(token, url);

  if (result.error) {
    console.error('Google Calendar error:', result.error.message);
    return [];
  }

  const events = result.items || [];
  console.log(`Found ${events.length} Google Calendar events`);

  const bookings = [];

  for (const event of events) {
    const subject = event.summary || '';
    const body = event.description || '';
    const location = event.location || '';
    const allText = subject + ' ' + body + ' ' + location;
    const allTextUpper = allText.toUpperCase();

    // Skip car rentals
    if (/\b(car\s*rental|hertz|avis|europcar|sixt|enterprise|rent.?a.?car|pick.?up.*drop.?off|vehicle\s*collect)/i.test(allText)) continue;

    // Skip events that don't look like travel
    const isFlight = /\b(flight|fly|depart|arrive|airport|boarding|BA\d|EK\d|LH\d|AF\d|AZ\d|FR\d|U2\d|QR\d|EY\d|SQ\d|CX\d|TK\d)/i.test(allText);
    const isHotel = /\b(hotel|check.?in|check.?out|booking|reservation|stay|accommodation|airbnb)/i.test(allText);

    if (!isFlight && !isHotel) continue;

    // Google Calendar uses date or dateTime
    const startStr = event.start?.dateTime || event.start?.date;
    const endStr = event.end?.dateTime || event.end?.date;
    const startDt = parseDate(startStr);
    const endDt = parseDate(endStr);
    if (!startDt) continue;

    if (isFlight) {
      const flights = extractFlights(allTextUpper);
      const dest = extractDestination(allTextUpper);
      bookings.push({
        type: 'flight',
        date: fmtDate(startDt),
        flights: flights.join(', '),
        city: dest?.city || extractCity(allText) || '',
        country: dest?.country || '',
        place: '',
        source: 'google-calendar',
        raw: subject
      });
    }

    if (isHotel && endDt) {
      const hotelName = extractHotelName(subject) || location || extractHotelName(allText) || '';
      const nights = dateRange(startDt, new Date(endDt.getTime() - 86400000));

      for (const dateStr of nights) {
        bookings.push({
          type: 'hotel',
          date: dateStr,
          flights: '',
          city: extractCity(allText) || location || '',
          country: '',
          place: hotelName,
          source: 'google-calendar',
          raw: subject
        });
      }
    }
  }

  return bookings;
}

// ===== GMAIL =====
// Scans specific Gmail folders (labels) for travel emails, matching Outlook approach.
// Gmail nested labels use "/" separator, e.g. "Inbox/Hotels" for a subfolder under Inbox.

const GMAIL_HOTEL_LABELS = ['Hotels', 'Hotel', 'Inbox/Hotels', 'Inbox/Hotel'];
const GMAIL_FLIGHT_LABELS = ['Flights', 'Flight', 'Inbox/Flights', 'Inbox/Flight'];

async function findGmailLabel(token, candidates) {
  // List all labels
  const result = await googleGet(token, 'https://gmail.googleapis.com/gmail/v1/users/me/labels');
  if (result.error) {
    console.error('Gmail labels error:', result.error.message);
    return null;
  }

  const labels = result.labels || [];
  const labelNames = labels.map(l => l.name);
  console.log('  Gmail labels:', labelNames.filter(n => !n.startsWith('CATEGORY_') && !n.startsWith('UNREAD') && !n.startsWith('STARRED')).join(', '));

  for (const candidate of candidates) {
    const match = labels.find(l => l.name === candidate || l.name.toLowerCase() === candidate.toLowerCase());
    if (match) {
      console.log(`  Found Gmail label: "${match.name}" (id: ${match.id})`);
      return match.id;
    }
  }
  return null;
}

async function processGmailFolder(token, labelId, folderType) {
  // Get recent messages with this label (last 7 days)
  const since = new Date();
  since.setDate(since.getDate() - DAYS_BACK);
  const q = `newer_than:${DAYS_BACK}d`;

  const searchUrl = `https://gmail.googleapis.com/gmail/v1/users/me/messages?labelIds=${labelId}&q=${encodeURIComponent(q)}&maxResults=20`;
  const searchResult = await googleGet(token, searchUrl);

  if (searchResult.error) {
    console.error('Gmail folder error:', searchResult.error.message);
    return [];
  }

  const messageIds = (searchResult.messages || []).map(m => m.id);
  console.log(`  Found ${messageIds.length} recent emails in Gmail ${folderType} folder`);

  const bookings = [];

  for (const msgId of messageIds) {
    const msgUrl = `https://gmail.googleapis.com/gmail/v1/users/me/messages/${msgId}?format=metadata&metadataHeaders=Subject&metadataHeaders=From`;
    const msg = await googleGet(token, msgUrl);
    if (msg.error) continue;

    const headers = msg.payload?.headers || [];
    const subject = headers.find(h => h.name === 'Subject')?.value || '';
    const from = headers.find(h => h.name === 'From')?.value || '';
    const snippet = msg.snippet || '';
    const allText = subject + ' ' + snippet;
    const allTextUpper = allText.toUpperCase();

    // Skip car rentals
    if (/\b(car\s*rental|hertz|avis|europcar|sixt|enterprise|rent.?a.?car)/i.test(allText)) {
      console.log(`  🚗 SKIP car rental: ${subject.slice(0, 60)}`);
      continue;
    }

    // Use folder type as hint (same as Outlook email processing)
    const isFlight = folderType === 'flight' ||
                     /\b(flight|itinerary|boarding|e-?ticket|airline)/i.test(allText) ||
                     extractFlights(allTextUpper).length > 0;
    const isHotel = folderType === 'hotel' ||
                    /\b(hotel|reservation|check.?in|booking|stay|accommodation|airbnb|nights?)/i.test(allText);

    // Extract dates from snippet text
    const extractedDates = extractRawDates(allText);
    extractedDates.sort((a, b) => a - b);

    if (isFlight && extractedDates.length > 0) {
      const flights = extractFlights(allTextUpper);
      const dest = extractDestination(allTextUpper);
      // Multi-leg fix: push one booking per unique date (up to 30 days apart).
      // A round-trip itinerary email contains both outbound and return dates;
      // previously only extractedDates[0] (the outbound) became a booking.
      const flightDates = uniqueFlightDates(extractedDates);
      for (const d of flightDates) {
        bookings.push({
          type: 'flight',
          date: fmtDate(d),
          flights: flights.join(', '),
          city: dest?.city || '',
          country: dest?.country || '',
          place: '',
          source: 'gmail',
          raw: subject
        });
      }
    }

    if (isHotel && extractedDates.length >= 2) {
      const hotelName = extractHotelName(allText) || '';
      const checkIn = extractedDates[0];
      const checkOut = extractedDates[extractedDates.length - 1];
      const daySpan = Math.round((checkOut - checkIn) / 86400000);
      if (daySpan > 30 || daySpan < 1) continue;
      const nights = dateRange(checkIn, new Date(checkOut.getTime() - 86400000));

      for (const dateStr of nights) {
        bookings.push({
          type: 'hotel',
          date: dateStr,
          flights: '',
          city: extractCity(allText) || '',
          country: '',
          place: hotelName,
          source: 'gmail',
          raw: subject
        });
      }
    }
  }

  return bookings;
}

async function processGmail(token) {
  const allBookings = [];

  // Process hotel folders
  console.log('Looking for Gmail hotel folder...');
  const hotelLabelId = await findGmailLabel(token, GMAIL_HOTEL_LABELS);
  if (hotelLabelId) {
    const hotelBookings = await processGmailFolder(token, hotelLabelId, 'hotel');
    allBookings.push(...hotelBookings);
  } else {
    console.log('  No hotel label found in Gmail.');
  }

  // Process flight folders
  console.log('Looking for Gmail flights folder...');
  const flightLabelId = await findGmailLabel(token, GMAIL_FLIGHT_LABELS);
  if (flightLabelId) {
    const flightBookings = await processGmailFolder(token, flightLabelId, 'flight');
    allBookings.push(...flightBookings);
  } else {
    console.log('  No flights label found in Gmail.');
  }

  return allBookings;
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

// Dedupe a sorted Date[] by YYYY-MM-DD and keep only entries within 30 days
// of the earliest. A single-leg flight email returns one date; a round-trip
// itinerary (outbound + return in the same email) returns one date per leg.
// The 30-day cap rejects unrelated noise like booking-creation timestamps.
function uniqueFlightDates(dates) {
  if (!dates || !dates.length) return [];
  const earliest = dates[0];
  const seen = new Set();
  const out = [];
  for (const d of dates) {
    const span = Math.round((d - earliest) / 86400000);
    if (span < 0 || span > 30) continue;
    const key = fmtDate(d);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(d);
  }
  return out;
}

// ===== SMART DATE EXTRACTION =====

// Date format patterns (reusable)
const DATE_FMTS = [
  /(\d{4}-\d{2}-\d{2})/,
  /(\d{1,2})\s+(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{4})/i,
  /(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2}),?\s+(\d{4})/i,
  /(?:Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2}),?\s+(\d{4})/i,
];

// Extract dates that appear after labeled keywords like "Arrival Date:", "Check-in:", etc.
// These are high-confidence dates from hotel/flight confirmations.
function extractLabeledDates(text) {
  if (!text) return [];
  const labels = [
    /(?:arrival|check.?in|depart(?:ure)?|check.?out|start|end|from|to)\s*(?:date)?\s*:?\s*/gi,
  ];
  const dates = [];
  for (const labelPat of labels) {
    let lm;
    while ((lm = labelPat.exec(text)) !== null) {
      // Look at the text right after the label (next 60 chars)
      const after = text.slice(lm.index + lm[0].length, lm.index + lm[0].length + 60);
      for (const datePat of DATE_FMTS) {
        const dm = datePat.exec(after);
        if (dm) {
          const d = parseDate(dm[0]);
          if (d && d.getFullYear() >= 2025 && d.getFullYear() <= 2028) {
            dates.push(d);
          }
          break;
        }
      }
    }
  }
  return dates;
}

// Raw date extraction from text — used as fallback on short text (subject + preview).
// Scans for all date patterns in the text.
function extractRawDates(text) {
  if (!text) return [];
  const datePatterns = [
    /(\d{4}-\d{2}-\d{2})/g,
    /(\d{1,2})\s+(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{4})/gi,
    /(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+(\d{1,2}),?\s+(\d{4})/gi,
  ];
  const dates = [];
  for (const pat of datePatterns) {
    let m;
    while ((m = pat.exec(text)) !== null) {
      const d = parseDate(m[0]);
      if (d && d.getFullYear() >= 2025 && d.getFullYear() <= 2028) {
        dates.push(d);
      }
    }
  }
  return dates;
}

// ===== PARSING =====

// Extract flight numbers from text (e.g., BA123, EK456, LH1234)
// Filters out common false positives: words like "AT", "TO", "IN", "NO", "IF", "OR", "BY", "ON", "UP"
// followed by numbers (e.g., "at 0900", "in 2026", "no 123")
function extractFlights(text) {
  if (!text) return [];
  // Common 2-letter words that aren't airline codes
  const FALSE_PREFIXES = new Set(['AT','TO','IN','NO','IF','OR','BY','ON','UP','AN','AS','BE','DO','GO','HE','IS','IT','ME','MY','OF','SO','US','WE']);
  // UK postcode prefixes (e.g. OX4, UB7, SW1, EC2, WC1, SE1, NW3, etc.)
  const UK_POSTCODE_PREFIXES = new Set([
    'AB','AL','BA','BB','BD','BH','BL','BN','BR','BS','BT','CA','CB','CF','CH','CM','CO','CR','CT','CV','CW',
    'DA','DD','DE','DG','DH','DL','DN','DT','DY','EC','EH','EN','EX','FK','FY','GL','GU','GY',
    'HA','HD','HG','HP','HR','HS','HU','HX','IG','IM','IP','IV','JE','KA','KT','KW','KY',
    'LA','LD','LE','LL','LN','LS','LU','ME','MK','ML','NE','NG','NN','NP','NR','NW',
    'OL','OX','PA','PE','PH','PL','PO','PR','RG','RH','RM','SA','SE','SG','SK','SL','SM','SN','SO',
    'SP','SR','SS','ST','SW','SY','TA','TD','TF','TN','TQ','TR','TS','TW','UB',
    'WA','WC','WD','WF','WN','WR','WS','WV','YO','ZE'
  ]);
  const pattern = /\b([A-Z]{2})\s*(\d{1,4})\b/g;
  const flights = [];
  let m;
  while ((m = pattern.exec(text)) !== null) {
    const code = m[1];
    const num = m[2];
    // Skip common English words followed by numbers
    if (FALSE_PREFIXES.has(code)) continue;
    // Skip UK postcodes (2 letters + 1-2 digits)
    if (num.length <= 2 && UK_POSTCODE_PREFIXES.has(code)) continue;
    // Skip numbers that look like years (2025-2028)
    const numVal = parseInt(num);
    if (num.length === 4 && numVal >= 2024 && numVal <= 2030) continue;
    flights.push(code + num);
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
  // Generic words that are NOT hotel names
  const JUNK_NAMES = /^(confirmation|reservation|booking|receipt|itinerary|details|reminder|update|notice|alert|your|the|a|for|at)$/i;
  // Common patterns in hotel confirmation subjects/bodies
  const patterns = [
    /(?:reservation\s+at|stay\s+at|check.?in\s+at|welcome\s+to|booking\s+at)\s+(.+?)(?:\s*[!.\n]|\s+in\s+|\s+on\s+|$)/i,
    /(?:property|hotel|resort)\s*(?:name)?\s*:?\s*(.+?)(?:\s*[.\n,]|$)/i,
    /(?:check.?in|check.?out)\s+(.+?)(?:\s*[-–|,]|\s+on\s+|$)/i,
    /(?:booking|reservation|confirmation)\s+(?:at|for)\s+(.+?)(?:\s*[-–|,]|\s+in\s+|\s+on\s+|$)/i,
    /(?:your stay at|check.?in at|welcome to)\s+(.+?)(?:\s*[-–|,]|\s+on\s+|$)/i,
    /(?:hotel|resort|inn|lodge|hostel|apartment|residence|suites?|villa)\s+([A-Z][\w'']+(?:\s+[\w'']+){0,5})(?:\s*[-–|,.]|\s+on\s+|$)/i,
  ];
  for (const pat of patterns) {
    const m = pat.exec(text);
    if (m && m[1].trim().length > 2 && m[1].trim().length < 80) {
      const name = m[1].trim().replace(/\s+(your|the|a)\s*$/i, '').trim();
      // Skip generic/junk names
      if (JUNK_NAMES.test(name)) continue;
      return name;
    }
  }
  return null;
}

// Extract city from labeled fields in full body text
function extractLabeledCity(text) {
  if (!text) return null;
  const NOT_CITIES = /^(January|February|March|April|May|June|July|August|September|October|November|December|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Confirmation|Reservation|Booking|Arrival|Departure|Check|Date|Room|Guest|Night|Gmail|Outlook|Yahoo|Hotmail|Icloud|Email|Inbox)$/i;
  const patterns = [
    /(?:city|location|address|property\s+address)\s*:?\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)/,
    /\b(\d+[^,\n]{5,40}),\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\b/,
  ];
  for (const pat of patterns) {
    const m = pat.exec(text);
    const city = m ? (m[2] || m[1]) : null;
    if (city && !NOT_CITIES.test(city)) return city;
  }
  return null;
}

// Extract city from text
function extractCity(text) {
  if (!text) return null;
  // Words that look like cities but aren't (months, common nouns)
  const NOT_CITIES = /^(January|February|March|April|May|June|July|August|September|October|November|December|Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday|Confirmation|Reservation|Booking|Dear|Hello|Please|Thank|Thanks|Your|The|This|Arrival|Departure|Check|Date|Room|Guest|Total|Price|Rate|Night|Day|Gmail|Outlook|Yahoo|Hotmail|Icloud|Email|Inbox)$/i;
  // Look for "in <City>" or "<City>, <Country>" patterns
  const cityPattern = /\bin\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)/g;
  let m;
  while ((m = cityPattern.exec(text)) !== null) {
    if (!NOT_CITIES.test(m[1])) return m[1];
  }
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
      // For "check in" / "check out" events, the hotel name is usually in the location field
      const hotelName = extractHotelName(subject) || location || extractHotelName(allText) || '';
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
  const path = `/users/${USER_EMAIL}/mailFolders/${folderId}/messages?$filter=${encodeURIComponent(filter)}&$top=50&$select=subject,body,bodyPreview,receivedDateTime,from`;
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
    const preview = msg.bodyPreview || '';
    // Strip HTML from full body for structured date extraction only
    const rawBody = msg.body?.content || '';
    const fullBody = rawBody.replace(/<[^>]+>/g, ' ').replace(/&[a-z]+;/gi, ' ').replace(/\s+/g, ' ').trim();
    const from = msg.from?.emailAddress?.address || '';
    // Use subject + preview for keyword matching (avoids noise from full HTML body)
    const allText = subject + ' ' + preview;
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

    // --- DATE EXTRACTION ---
    // Strategy: First try labeled dates from full body (e.g. "Arrival Date: April 6, 2026").
    // Fall back to raw date extraction from subject + preview only (not full body, to avoid noise).
    let extractedDates = extractLabeledDates(fullBody);
    if (extractedDates.length < 2) {
      // Fall back to raw date scanning from subject + preview
      extractedDates = extractRawDates(allText);
    }

    // Sort dates and take first as check-in, last as check-out
    extractedDates.sort((a, b) => a - b);

    if (isFlight && extractedDates.length > 0) {
      const flights = extractFlights(allTextUpper);
      const dest = extractDestination(allTextUpper);
      // Multi-leg fix: push one booking per unique date (up to 30 days apart).
      // A round-trip itinerary email contains both outbound and return dates;
      // previously only extractedDates[0] (the outbound) became a booking,
      // silently dropping the return leg (e.g. BA591 May 23 outbound + May 25 return).
      const flightDates = uniqueFlightDates(extractedDates);
      for (const d of flightDates) {
        bookings.push({
          type: 'flight',
          date: fmtDate(d),
          flights: flights.join(', '),
          city: dest?.city || '',
          country: dest?.country || '',
          place: '',
          source: 'email',
          raw: subject
        });
      }
    }

    if (isHotel && extractedDates.length >= 2) {
      // Try preview first, then first 1500 chars of full body for hotel name
      const bodySnippet = fullBody.slice(0, 1500);
      const hotelName = extractHotelName(subject) || extractHotelName(preview) || extractHotelName(bodySnippet) || '';
      // Try preview, then labeled fields in full body for city
      const city = extractCity(preview) || extractLabeledCity(bodySnippet) || extractCity(bodySnippet) || '';
      const checkIn = extractedDates[0];
      const checkOut = extractedDates[extractedDates.length - 1];
      const daySpan = Math.round((checkOut - checkIn) / 86400000);
      if (daySpan > 30 || daySpan < 1) continue; // Sanity check
      const nights = dateRange(checkIn, new Date(checkOut.getTime() - 86400000));

      for (const dateStr of nights) {
        bookings.push({
          type: 'hotel',
          date: dateStr,
          flights: '',
          city: city,
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

    // Only protect GPS-confirmed entries — booking data can overwrite everything else
    if (current && (current.autoGps || current.gpsConfirmed) && current.city) {
      // GPS-confirmed: only merge flights, don't touch location
      if (booking.flights) {
        const merged = mergeFlights(current.flights, booking.flights);
        if (merged !== (current.flights || '')) {
          const sourceInfo = booking.source + ': ' + (booking.raw || '').slice(0, 120);
          const existingSource = current.bookingSource || '';
          const combinedSource = existingSource
            ? (existingSource.includes(sourceInfo) ? existingSource : existingSource + ' | ' + sourceInfo)
            : sourceInfo;
          updates['locations/' + dateStr + '/flights'] = merged;
          updates['locations/' + dateStr + '/bookingSource'] = combinedSource;
          mergedCount++;
          console.log(`  ✈ Merged flight ${booking.flights} into GPS-confirmed entry for ${dateStr}`);
        } else {
          skippedCount++;
        }
      } else {
        skippedCount++;
      }
      continue;
    }

    // Non-GPS entries: booking data can freely populate/overwrite
    // Build booking source audit trail
    const sourceInfo = booking.source + ': ' + (booking.raw || '').slice(0, 120);
    const existingSource = current?.bookingSource || '';
    const combinedSource = existingSource
      ? (existingSource.includes(sourceInfo) ? existingSource : existingSource + ' | ' + sourceInfo)
      : sourceInfo;

    const entry = {
      place: current?.place || booking.place || '',
      city: current?.city || booking.city || '',
      country: normalizeCountry(current?.country || booking.country || ''),
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
  console.log(`=== Midnight Tracker — Booking Sync v${SYNC_VERSION} ===`);
  console.log('Time:', new Date().toISOString());

  // --- Microsoft Outlook ---
  console.log('\nAuthenticating with Microsoft Graph...');
  const msToken = await getGraphToken();
  console.log('Authenticated.');

  const calendarBookings = await processCalendar(msToken);
  console.log(`Outlook Calendar: ${calendarBookings.length} booking entries found`);

  const emailBookings = await processEmails(msToken);
  console.log(`Outlook Email: ${emailBookings.length} booking entries found`);

  // --- Google (optional) ---
  let googleCalBookings = [];
  let gmailBookings = [];

  if (GOOGLE_ENABLED) {
    console.log('\nAuthenticating with Google...');
    try {
      const gToken = await getGoogleToken();
      console.log('Authenticated with Google.');

      googleCalBookings = await processGoogleCalendar(gToken);
      console.log(`Google Calendar: ${googleCalBookings.length} booking entries found`);

      gmailBookings = await processGmail(gToken);
      console.log(`Gmail: ${gmailBookings.length} booking entries found`);
    } catch(e) {
      console.error('Google API error (continuing with Outlook only):', e.message);
    }
  } else {
    console.log('\nGoogle not configured — skipping. Set GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN to enable.');
  }

  // Combine all sources (Outlook calendar > Google calendar > emails)
  const allBookings = [...calendarBookings, ...googleCalBookings, ...emailBookings, ...gmailBookings];

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
