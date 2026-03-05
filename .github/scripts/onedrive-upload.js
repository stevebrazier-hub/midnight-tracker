/**
 * OneDrive Backup Upload Script
 *
 * Uploads the latest backup files (JSON + XLSX) to Steve's OneDrive
 * so they sync to all machines automatically.
 *
 * Uses Microsoft Graph API with the same credentials as sync-bookings.
 *
 * Requires Azure AD app permission: Files.ReadWrite.All
 *
 * Environment variables:
 *   MS_TENANT_ID     - Azure AD tenant ID
 *   MS_CLIENT_ID     - Azure AD app client ID
 *   MS_CLIENT_SECRET - Azure AD app client secret
 *   MS_USER_EMAIL    - OneDrive owner (steveb@canapii.com)
 *   ONEDRIVE_FOLDER  - Target folder path (default: Midnight Tracker Backups)
 */

const https = require('https');
const fs = require('fs');
const path = require('path');

const USER_EMAIL = process.env.MS_USER_EMAIL || 'steveb@canapii.com';
const FOLDER = process.env.ONEDRIVE_FOLDER || 'Midnight Tracker Backups';
const BACKUP_DIR = path.join(process.cwd(), 'backups');

// ===== AUTH =====

async function getGraphToken() {
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: process.env.MS_CLIENT_ID,
    client_secret: process.env.MS_CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default'
  }).toString();

  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'login.microsoftonline.com',
      path: `/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`,
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

// ===== UPLOAD =====

async function uploadFile(token, filePath, fileName) {
  const fileData = fs.readFileSync(filePath);
  const encodedFolder = encodeURIComponent(FOLDER);
  const encodedFile = encodeURIComponent(fileName);

  // Use the simple upload endpoint (files < 4MB)
  // PUT /users/{email}/drive/root:/{folder}/{file}:/content
  const uploadPath = `/v1.0/users/${USER_EMAIL}/drive/root:/${encodedFolder}/${encodedFile}:/content`;

  return new Promise((resolve, reject) => {
    const req = https.request({
      hostname: 'graph.microsoft.com',
      path: uploadPath,
      method: 'PUT',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/octet-stream',
        'Content-Length': fileData.length
      }
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          const json = JSON.parse(data);
          console.log(`  ✓ Uploaded ${fileName} (${(fileData.length / 1024).toFixed(1)} KB)`);
          resolve(json);
        } else {
          console.error(`  ✗ Failed ${fileName}: ${res.statusCode}`);
          console.error('   ', data.slice(0, 300));
          reject(new Error(`Upload failed: ${res.statusCode}`));
        }
      });
    });
    req.on('error', reject);
    req.write(fileData);
    req.end();
  });
}

// ===== MAIN =====

async function main() {
  const today = new Date().toISOString().slice(0, 10);
  console.log('=== OneDrive Backup Upload ===');
  console.log('Date:', today);
  console.log('Target folder:', FOLDER);
  console.log('User:', USER_EMAIL);
  console.log('Client ID:', process.env.MS_CLIENT_ID ? process.env.MS_CLIENT_ID.slice(0, 8) + '...' : 'MISSING');
  console.log('Tenant ID:', process.env.MS_TENANT_ID ? process.env.MS_TENANT_ID.slice(0, 8) + '...' : 'MISSING');
  console.log('Client Secret:', process.env.MS_CLIENT_SECRET ? '***set***' : 'MISSING');

  const token = await getGraphToken();
  console.log('Authenticated with Graph API');
  console.log('Token preview:', token.slice(0, 20) + '...\n');

  // Upload latest.json as dated file + latest
  const jsonFile = path.join(BACKUP_DIR, 'latest.json');
  const xlsxFile = path.join(BACKUP_DIR, 'latest.xlsx');

  let uploaded = 0;

  if (fs.existsSync(jsonFile)) {
    console.log('Found latest.json (' + (fs.statSync(jsonFile).size / 1024).toFixed(1) + ' KB)');
    await uploadFile(token, jsonFile, `backup-${today}.json`);
    uploaded++;
    await uploadFile(token, jsonFile, 'latest.json');
    uploaded++;
  } else {
    console.log('WARNING: No latest.json found at ' + jsonFile);
  }

  if (fs.existsSync(xlsxFile)) {
    console.log('Found latest.xlsx (' + (fs.statSync(xlsxFile).size / 1024).toFixed(1) + ' KB)');
    await uploadFile(token, xlsxFile, `midnight-tracker-${today}.xlsx`);
    uploaded++;
    await uploadFile(token, xlsxFile, 'latest.xlsx');
    uploaded++;
  } else {
    console.log('WARNING: No latest.xlsx found at ' + xlsxFile);
  }

  console.log('\nDone. ' + uploaded + ' files uploaded to OneDrive/' + FOLDER);
  if (uploaded === 0) {
    console.error('ERROR: No files were uploaded!');
    process.exit(1);
  }
  process.exit(0);
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
