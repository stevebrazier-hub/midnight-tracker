# Midnight Tracker — Setup & Deployment

## Architecture
Same pattern as Chelsea Tickets:
- Single HTML file (`index.html`), no build step
- Firebase Realtime DB for sync
- Azure Static Web Apps for hosting, auto-deploy from GitHub `main` branch
- Custom domain: `midnight.cancomo.com`

## Step 1: Create Firebase Project

1. Go to https://console.firebase.google.com
2. Click **Add project**
3. Name it something like `midnight-tracker-steve` (must be globally unique)
4. Skip Google Analytics (not needed)
5. Once created, click the **web icon** (`</>`) to register a web app
6. Name it "Midnight Tracker", skip Firebase Hosting
7. **Copy the `firebaseConfig` object** it shows you — you need these values for `index.html`

## Step 2: Create Realtime Database

1. In the Firebase console left menu: **Build > Realtime Database**
2. Click **Create Database**
3. Choose location: **europe-west1** (Belgium) — same region as Chelsea Tickets
4. Select **Start in test mode** (we'll tighten rules later)
5. Once created, go to the **Rules** tab
6. Paste the contents of `database.rules.json` from this repo and **Publish**

## Step 3: Paste Config into index.html

Edit `index.html` and replace the placeholder `FIREBASE_CONFIG` block (~line 10 of the script) with your real values:

```js
const FIREBASE_CONFIG = {
  apiKey: "AIzaSy...",
  authDomain: "your-project.firebaseapp.com",
  databaseURL: "https://your-project-default-rtdb.europe-west1.firebasedatabase.app",
  projectId: "your-project",
  storageBucket: "your-project.firebasestorage.app",
  messagingSenderId: "123456789",
  appId: "1:123456789:web:abc123"
};
```

## Step 4: GitHub Repo

1. Create a new repo: `stevebrazier-hub/midnight-tracker`
2. Push this folder's contents (index.html, manifest.json, sw.js, database.rules.json)
3. The deployed file is `index.html` in the repo root — same as Chelsea Tickets

## Step 5: Azure Static Web App

1. Go to Azure Portal > Static Web Apps > Create
2. Link to the GitHub repo `stevebrazier-hub/midnight-tracker`
3. Branch: `main`
4. Build preset: **Custom**
5. App location: `/` (root)
6. Output location: leave blank
7. Skip API location
8. Azure will auto-deploy on every push to `main`

## Step 6: Custom Domain

1. In Azure Portal, go to your Static Web App > Custom domains
2. Add `midnight.cancomo.com`
3. In GoDaddy DNS, add a CNAME record:
   - Name: `midnight`
   - Value: the Azure-provided hostname (e.g. `nice-river-xxx.azurestaticapps.net` or similar)
4. Azure will auto-provision an SSL certificate

## Data Structure (Firebase Realtime DB)

```
/locations
  /2026-01-15
    place: "Hotel Negresco"
    city: "Nice"
    country: "France"
    flights: "BA569"
    notes: ""
    lat: 43.6947
    lon: 7.2562
  /2026-01-16
    ...
```

## Files

| File | Purpose |
|------|---------|
| `index.html` | The app (single file, deployed) |
| `manifest.json` | PWA manifest |
| `sw.js` | Service worker for offline support |
| `database.rules.json` | Firebase Realtime DB rules |
| `SETUP.md` | This file |
