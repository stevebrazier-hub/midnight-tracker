/**
 * One-time Google OAuth2 Setup Script
 *
 * Run this ONCE locally to get a refresh token for the sync script.
 *
 * Prerequisites:
 *   1. Go to https://console.cloud.google.com
 *   2. Create a project (e.g., "Midnight Tracker")
 *   3. Enable "Google Calendar API" and "Gmail API"
 *   4. Go to Credentials → Create Credentials → OAuth 2.0 Client ID
 *      - Application type: "Desktop app"
 *      - Name: "Midnight Tracker Sync"
 *   5. Download the JSON or note the Client ID and Client Secret
 *   6. Go to OAuth consent screen → Add test user: swbrazier@gmail.com
 *      (or publish the app if you want to skip test user mode)
 *
 * Usage:
 *   node google-oauth-setup.js <CLIENT_ID> <CLIENT_SECRET>
 *
 * It will:
 *   1. Print a URL — open it in your browser and sign in with swbrazier@gmail.com
 *   2. After consent, Google redirects to localhost with a code
 *   3. The script exchanges the code for a refresh token
 *   4. Add the refresh token as a GitHub secret: GOOGLE_REFRESH_TOKEN
 */

const http = require('http');
const https = require('https');
const { URL } = require('url');

const CLIENT_ID = process.argv[2];
const CLIENT_SECRET = process.argv[3];

if (!CLIENT_ID || !CLIENT_SECRET) {
  console.error('Usage: node google-oauth-setup.js <CLIENT_ID> <CLIENT_SECRET>');
  console.error('\nGet these from https://console.cloud.google.com → Credentials → OAuth 2.0 Client ID');
  process.exit(1);
}

const REDIRECT_URI = 'http://localhost:3847/callback';
const SCOPES = [
  'https://www.googleapis.com/auth/calendar.readonly',
  'https://www.googleapis.com/auth/gmail.readonly'
].join(' ');

// Build authorization URL
const authUrl = 'https://accounts.google.com/o/oauth2/v2/auth?' + new URLSearchParams({
  client_id: CLIENT_ID,
  redirect_uri: REDIRECT_URI,
  response_type: 'code',
  scope: SCOPES,
  access_type: 'offline',
  prompt: 'consent'
}).toString();

console.log('\n=== Google OAuth2 Setup for Midnight Tracker ===\n');
console.log('1. Open this URL in your browser:\n');
console.log(authUrl);
console.log('\n2. Sign in with swbrazier@gmail.com and grant access.\n');
console.log('Waiting for callback on localhost:3847...\n');

// Start a temporary local server to receive the OAuth callback
const server = http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://localhost:3847');
  if (url.pathname !== '/callback') {
    res.writeHead(404);
    res.end('Not found');
    return;
  }

  const code = url.searchParams.get('code');
  const error = url.searchParams.get('error');

  if (error) {
    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end('<h1>Error</h1><p>' + error + '</p>');
    console.error('OAuth error:', error);
    process.exit(1);
  }

  if (!code) {
    res.writeHead(400, { 'Content-Type': 'text/html' });
    res.end('<h1>Missing code</h1>');
    return;
  }

  // Exchange code for tokens
  try {
    const tokens = await exchangeCode(code);
    const refreshToken = tokens.refresh_token;

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end('<h1>Success!</h1><p>You can close this window. Check the terminal for your refresh token.</p>');

    console.log('=== SUCCESS ===\n');
    console.log('Refresh token:\n');
    console.log(refreshToken);
    console.log('\n\nNow add these THREE GitHub secrets to your repo:');
    console.log('  https://github.com/stevebrazier-hub/midnight-tracker/settings/secrets/actions\n');
    console.log(`  GOOGLE_CLIENT_ID     = ${CLIENT_ID}`);
    console.log(`  GOOGLE_CLIENT_SECRET = ${CLIENT_SECRET}`);
    console.log(`  GOOGLE_REFRESH_TOKEN = ${refreshToken}`);
    console.log('\nDone! The next sync-bookings run will include Google data.\n');

    setTimeout(() => process.exit(0), 1000);
  } catch(e) {
    res.writeHead(500, { 'Content-Type': 'text/html' });
    res.end('<h1>Error exchanging code</h1><p>' + e.message + '</p>');
    console.error('Token exchange error:', e);
    process.exit(1);
  }
});

server.listen(3847, () => {});

function exchangeCode(code) {
  const body = new URLSearchParams({
    code,
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    redirect_uri: REDIRECT_URI,
    grant_type: 'authorization_code'
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
          if (json.refresh_token) resolve(json);
          else reject(new Error('No refresh token in response: ' + JSON.stringify(json)));
        } catch(e) { reject(e); }
      });
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}
