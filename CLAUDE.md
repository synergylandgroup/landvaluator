# LandValuator — Developer Reference (CLAUDE.md)

**Synergy Land Investments | Production | April 2026**

> Complete reference for the LandValuator web app. Written for Claude Code and long-term maintainability — no prior context required.

---

## 1. System Overview

LandValuator is a single-page web app for land investors. Users draw pricing zones over county maps, connect Google Sheets containing parcel data exported from LandInsights (LI), and sync zone assignments back to their spreadsheet for mail campaign generation.

**Live URL:** https://landvaluator.app  
**GitHub:** synergylandgroup/landvaluator  
**Hosting:** Netlify (auto-deploys from GitHub main)  
**Auth + DB:** Supabase (project: dcrxczsgcuiwimwpokxo, region: us-west-1)  
**Map:** Mapbox GL JS v3.3.0  

---

## 2. File Inventory

| File | Lines | Role |
|------|-------|------|
| `index.html` | ~719 | All CSS, HTML structure, modals, map UI. No JS logic. |
| `app.js` | ~3,292 | All application + auth logic. Supabase client, DB adapter, map, zones, sheets. |
| `netlify/functions/sheets-read.js` | ~80 | Reads properties from Google Sheets via service account JWT. |
| `netlify/functions/sheets-write-zones.js` | ~65 | Writes zone assignments to Scrubbed and Priced tab. |
| `netlify/functions/sheets-write-pricing.js` | ~60 | Writes pricing tiers to Pricing Settings tab. |
| `netlify/functions/sheets-trigger-refresh.js` | ~30 | Calls GAS doPost to trigger refreshOfferPrices() in sheet. Currently wired but GAS_REFRESH_URL env var must be set. |
| `netlify/functions/auth-callback.js` | ~55 | PKCE password reset handler. Exchanges code, redirects with tokens. |
| `netlify.toml` | 8 | Build config + /auth/callback → auth-callback redirect rule. |
| `favicon.png` | — | 512px PNG favicon. |
| `GS_Apps_Script_.gs` | ~1,290 | Google Apps Script bound to each county spreadsheet. Handles data loading, pricing, zone sync. |

> **CRITICAL:** Never merge `index.html` and `app.js` into a single file. The split was introduced after a monolithic file caused silent truncation at ~26,000 chars during Claude file downloads, breaking deployments.

---

## 3. Supabase Configuration

| Setting | Value |
|---------|-------|
| Project URL | https://dcrxczsgcuiwimwpokxo.supabase.co |
| Anon Key | eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9... (publishable, safe for frontend) |
| Region | us-west-1 (North California) |
| Auth flow | Email/password. PKCE for password reset. |
| Email confirmation | OFF (disabled for frictionless onboarding) |
| Site URL | https://landvaluator.app |
| Redirect URLs | https://landvaluator.app, https://landvaluator.app/auth/callback, http://localhost:3000 |

### 3.1 Database Tables

| Table | Key Columns | Notes |
|-------|-------------|-------|
| `zones` | id, user_id, data jsonb, updated_at | All zone polygons as JSON array. Unique on user_id. |
| `unassigned_zones` | id, user_id, data jsonb, updated_at | Unassigned virtual zone pricing. Unique on user_id. |
| `sheet_configs` | id, user_id, configs jsonb, updated_at | Per-county sheet configs object. Unique on user_id. |
| `app_state` | id, user_id, state text, county text, updated_at | Last selected state/county. Unique on user_id. |
| `ui_state` | id, user_id, key text, value jsonb, updated_at | Accordion open/close, tooltip state. Unique on (user_id, key). |

All tables have Row Level Security (RLS) enabled. Unique constraint on user_id is required for upsert onConflict to work.

---

## 4. Netlify Environment Variables

| Variable | Purpose |
|----------|---------|
| `GOOGLE_SERVICE_ACCOUNT` | Full JSON string of Google service account credentials. **Never commit to GitHub.** |
| `GAS_REFRESH_URL` | Exec URL of Apps Script web app deployment. Used by sheets-trigger-refresh.js to auto-fire refreshOfferPrices(). Currently set but the doPost in GAS must be deployed as a web app to activate. |
| `NETLIFY_TOKEN` | Netlify API token. |
| `NETLIFY_SITE_ID` | Netlify site ID. |

**Service account:** landvaluator-sheets@landvaluator.iam.gserviceaccount.com (under synergylandinvestments@gmail.com GCP project)

---

## 5. Authentication System

### 5.1 Auth Flow

- User visits app → Supabase `getSession()` pre-fetches existing session before map loads
- No session: auth modal shown, map loads in background
- Valid session: `_currentUser` set, `map.on('load')` calls `_initAppAfterAuth()` directly
- `onAuthStateChange` fires on every session change — central hub for all auth state

### 5.2 Auth Functions (lines 37–175)

| Function | Purpose |
|----------|---------|
| `_authSwitchTab(tab)` | Toggles login/signup form visibility |
| `_authLogin()` | Email/password sign in |
| `_authSignup()` | Create account with first/last name in metadata |
| `_authSignOut()` | Signs out, clears all app state, shows auth modal |
| `_authShowReset()` | Switch from login modal to reset modal |
| `_authSendReset()` | Send password reset email |
| `_authSetNewPassword()` | Update password after reset flow, then init app |
| `_toggleUserMenu()` | Open/close user dropdown in header |
| `_updateUserUI(user)` | Update header: avatar, first name, full name, email |

### 5.3 Password Reset Flow (PKCE)

1. User clicks Forgot Password → reset email sent to landvaluator.app/auth/callback?type=recovery
2. Netlify routes /auth/callback → auth-callback.js function
3. Function exchanges ?code= param with Supabase, gets access_token + refresh_token
4. Function redirects to landvaluator.app?type=recovery#access_token=...&refresh_token=...
5. App detects ?type=recovery, sets `_passwordRecoveryMode = true`
6. App calls `supabase.auth.setSession()` with hash tokens
7. `onAuthStateChange` fires — `_passwordRecoveryMode` blocks normal login init
8. Set New Password modal shown, user enters new password
9. `supabase.auth.updateUser({password})` called, modal closes, app initializes normally

### 5.4 Key Auth Flags

| Variable | Purpose |
|----------|---------|
| `_currentUser` | Supabase user object. Set by `getSession()` pre-fetch and `onAuthStateChange`. |
| `_authAppReady` | True after `_initAppAfterAuth()` called. Reset to false on logout so re-login triggers fresh init. |
| `_mapLoadFired` | True after `map.on('load')`. Gates whether `_initAppAfterAuth` runs immediately. |
| `_passwordRecoveryMode` | True during PKCE reset flow. Blocks SIGNED_IN from initializing app. |

---

## 6. Database Adapter (DB object, lines ~265–365)

All persistence flows through the DB object. Methods are async and use Supabase. County list cache stays in localStorage (UI cache only, no user data).

| Method | Async? | Behavior |
|--------|--------|---------|
| `DB.saveZones(zones)` | Yes | Upserts zones array. Fire-and-forget OK. |
| `DB.loadZones()` | Yes | Fetches data column from zones table. |
| `DB.saveSheetConfigs(configs)` | Yes | Upserts full sheetConfigs object. |
| `DB.loadSheetConfigs()` | Yes | Fetches configs column from sheet_configs table. |
| `DB.saveAppState(state)` | Yes | Upserts {state, county} to app_state. |
| `DB.loadAppState()` | Yes | Fetches state + county columns. |
| `DB.saveUIState(key, value)` | No (sync+async) | Writes to `_uiStateCache` immediately, upserts to Supabase in background. |
| `DB.loadUIState(key, fallback)` | No (sync) | Reads from `_uiStateCache`. Returns fallback if key not found. |
| `DB.loadAllUIState()` | Yes | Fetches all ui_state rows, populates `_uiStateCache`. Called once after login. |
| `DB.saveUnassigned(entries)` | Yes | Upserts unassigned zone pricing to unassigned_zones table. |
| `DB.loadUnassigned()` | Yes | Fetches data column from unassigned_zones table. |
| `DB.saveCountyCache(abbr, counties)` | No | localStorage only. 30-day cache. |
| `DB.loadCountyCache(abbr)` | No | localStorage only. Returns parsed cache or null. |

**UI State Cache Strategy:** `renderPolygonList()` is synchronous and called 22+ times throughout the app. `DB.loadAllUIState()` fetches all rows in one query after login and populates `_uiStateCache`. All reads are then instant in-memory lookups.

---

## 7. Page Init Sequence

### 7.1 Pre-map (synchronous, on script parse)
- Supabase client created (`_supa`)
- Password recovery URL param detected
- `getSession()` pre-fetch — async IIFE sets `_currentUser` before map loads
- Auth functions defined, `onAuthStateChange` registered
- DB adapter defined with `_uiStateCache`
- STATES, STATE_FIPS, COLORS, map constants defined
- Mapbox map initialized

### 7.2 map.on('load')
- `_initDrawLayers()`, `_initPinLayer()` called
- `_mapLoadFired = true`
- If `_currentUser` already set: `_authAppReady = true`, `_initAppAfterAuth()` called immediately
- If no user: auth modal shown, wait for `onAuthStateChange`

### 7.3 _initAppAfterAuth() — runs after auth confirmed
1. `DB.loadAllUIState()` — populates `_uiStateCache` in one query
2. `_initTooltipToggle()` — reads tooltips_off from cache
3. `DB.loadSheetConfigs()` — populates sheetConfigs object
4. `loadZonesFromURL()` — checks ?zones= or ?share= params
5. `restoreZones()` — loads zones from Supabase, renders sidebar, draws layers
6. `setTimeout 600ms` — `_rebuildAllLabels()`, `_loadAllCountyBoundaries(true)`, `_mapInitComplete = true`
7. Deep-link check — reads ?state= and ?county= params
8. `DB.loadAppState()` — restores last selected state/county
9. `loadCounties().then()` — restores county dropdown, sheetConfig, reconnects sheets

---

## 8. Global State Variables

### 8.1 Auth & Init State
| Variable | Type | Description |
|----------|------|-------------|
| `_supa` | Object | Supabase client instance. |
| `_currentUser` | Object\|null | Supabase user. Set by getSession() and onAuthStateChange. |
| `_authAppReady` | Boolean | True after `_initAppAfterAuth()` called. Reset on logout. |
| `_mapLoadFired` | Boolean | True after map.on('load'). |
| `_passwordRecoveryMode` | Boolean | True during PKCE reset flow. |
| `_uiStateCache` | Object | In-memory mirror of ui_state table. |

### 8.2 Core Data
| Variable | Type | Description |
|----------|------|-------------|
| `polygons` | Array | All zone polygons. Shape: `{id, name, letter, stateAbbr, countyName, color, points[], description, pricingTiers[], propCount, labelMarker, handles[], _isRect, _bounds, _isUnassigned}` |
| `properties` | Array | Loaded property records. Shape: `{lat, lng, apn, county, state, acreage, liAcreage, parcelLink, ownerName, zone, rowIndex, marker: null}` |
| `sheetConfigs` | Object | Per-county sheet configs keyed by `"stateAbbr\|countyName"`. Persisted to Supabase. |
| `sheetConfig` | Object | Active config for currently selected county. |

### 8.3 Map State
| Variable | Type | Description |
|----------|------|-------------|
| `countySourceId` | String\|null | Mapbox source ID of active county boundary layer. |
| `_countyLayers` | Object | `"stateAbbr\|countyName"` → sourceId for all county boundary layers. |
| `_pendingCountyGeoJSON` | Object\|null | GeoJSON of selected county. Used to validate polygon draws. |
| `_countyGeoJSONCache` | Object | In-memory cache of county GeoJSON. |
| `_mapInitComplete` | Boolean | Set true 600ms after map load. Gates style.load county layer redraw. |
| `COUNTY_PILL_ZOOM` | Number | = 7. Below: county pills shown. At/above: zone labels shown. |
| `_pinsVisible` | Boolean | Whether property pins are currently visible. |

---

## 9. Google Sheets Integration

### 9.1 Sheet Config Object
Stored per county in `sheetConfigs["stateAbbr|countyName"]`:
```javascript
{
  sheetId: "1G1OQg8...",      // Google Sheets file ID
  sheetUrl: "https://...",    // Full URL (for pre-fill on disconnect)
  sheetTitle: "Newaygo...",   // Spreadsheet display name
  stateAbbr: "MI",
  countyName: "Newaygo",
  sheetName: "LI Raw Dataset",
  colLat: "Latitude",
  colLng: "Longitude",
  colAPN: "APN",
  colCity: "City",
  colCounty: "County",
  colState: "State",
  colZip: "ZIP",
  colZone: "County Zone",
}
```

### 9.2 Connection Flow (connectSheets)
1. Parse sheet ID from URL input (or saved config for current county only)
2. Check sheet ID not already used by a different county key
3. Fetch properties via sheets-read Netlify function
4. **County name check on raw data** — compare property county fields against selected county BEFORE filtering. Blocks import if mismatch detected with full detail toast (8 second duration).
5. **County boundary validation** — `_validatePropertiesInCounty()` fetches county GeoJSON and checks coordinates. Returns null if boundary fetch fails — import is blocked, not silently passed.
6. Save config to Supabase only on success (`_finishSheetConnect`)
7. Auto-assign properties to existing zones
8. Set propCount on each zone from assignment results
9. Persist zones, render sidebar, close modal

### 9.3 Disconnect Flow (disconnectSheet)
- Deletes sheet config from `sheetConfigs` and saves to Supabase
- Filters properties for that county out of `properties` array
- Removes virtual unassigned polygon for that county
- Resets propCount to 0 on all real zones for that county
- Updates stat counters
- Resets header connection indicator if that was the active county
- Calls `closeSheetsModal()` automatically

### 9.4 Delete County Flow (deleteCounty)
- Removes all zone polygons and map layers for that county
- Removes county boundary layer if no zones remain
- **Also clears sheet config from sheetConfigs and Supabase**
- **Filters out all properties for that county**
- **Removes virtual unassigned polygon**
- **Resets stat counters**
- **Resets header connection indicator if that was the active county**

### 9.5 Validation Safeguards
- **Modal open check:** On `openSheetsModal()`, if the saved config's sheetId is also registered under a different county key in sheetConfigs, the bad config is auto-cleared and the modal shows disconnected state.
- **Raw county check:** Runs on `data.properties` (raw API response) before `loadPropertiesFromFunction` filters anything. Catches wrong-county sheets before they connect.
- **On block:** Rolls back any saved config for that county, resets modal to disconnected, shows 8-second toast with full detail.

### 9.6 Netlify Functions

**sheets-read.js**
- Endpoint: `/.netlify/functions/sheets-read`
- Method: POST
- Request: `{sheetId, sheetName, colCounty, colAPN}`
- Response: `{spreadsheetTitle, properties[], scrubbedApns[], ownerMap{}}`
- Auth: Google service account JWT from `GOOGLE_SERVICE_ACCOUNT` env var

**sheets-write-zones.js**
- Endpoint: `/.netlify/functions/sheets-write-zones`
- Request: `{sheetId, sheetName, assignments[{apn, zone}]}`
- Response: `{success: true, updated: N}`

**sheets-write-pricing.js**
- Endpoint: `/.netlify/functions/sheets-write-pricing`
- Request: `{sheetId, tiers[{zone, minAcres, maxAcres, pricePerAcre}]}`
- Response: `{success: true}`

**sheets-trigger-refresh.js**
- Endpoint: `/.netlify/functions/sheets-trigger-refresh`
- Calls `GAS_REFRESH_URL` (Apps Script doPost) to trigger `refreshOfferPrices()` in the sheet
- Fire-and-forget from app.js after Save & Sync succeeds
- Fails silently if `GAS_REFRESH_URL` not set

---

## 10. Google Sheets Modal Design

The modal has two states:

**Disconnected state:** Shows 3 numbered steps:
1. Get the template (link to make-a-copy URL)
2. Share with service account (copy email button)
3. Paste sheet URL (input field)

**Connected state:** Steps hidden. Shows only:
- Green status box with sheet title and Open ↗ button
- Disconnect & re-enter URL link
- Cancel / Refresh & Load buttons

**Template URL:** `https://docs.google.com/spreadsheets/d/1HpTfSOxSMOPBzw8dZ0WD26Oo8kdMBNC5Rell0uucejs/copy`

**Service account email** is obfuscated in app.js as a char code array decoded at runtime. Current value: `landvaluator-sheets@landvaluator.iam.gserviceaccount.com`

---

## 11. Zone System

### 11.1 Zone Polygon Object
```javascript
{
  id: "uuid",
  name: "Zone A",
  letter: "A",
  stateAbbr: "MI",
  countyName: "Newaygo",
  color: "#e05252",
  points: [[lng, lat], ...],  // GeoJSON coordinate order
  description: "",
  pricingTiers: [{minAcres, maxAcres, pricePerAcre}],
  propCount: 485,
  labelMarker: MapboxMarker,
  handles: [],
  _isRect: false,
  _bounds: null,
  _isUnassigned: false,       // true for virtual unassigned polygon
}
```

### 11.2 Zone Layer System (1.2)
- Page restore: `skipLayers=true` in `_loadZone()` — zone data loaded but no fill/line layers drawn
- County select: `_restoreAllZoneLayers()` draws all fill/line layers
- Zoom < 7: county pills shown, fill layers hidden
- Zoom ≥ 7: zone labels and fill layers shown, county pills hidden
- `_mapInitComplete` flag prevents premature county layer redraw on first page load

### 11.3 propCount Behavior
- `propCount` is only restored from Supabase on page load if a sheet is connected for that county (`_hasSheet` check in `_loadZone`)
- Reset to 0 on disconnect, delete county, and logout
- Repopulated immediately on reconnect via `properties.filter(p => p.zone === poly.letter).length`

### 11.4 Unassigned Zone
- Virtual polygon with `id = "__unassigned__stateAbbr|countyName"` and `_isUnassigned: true`
- Created dynamically in `renderPolygonList()` when unassigned property count > 0
- Removed on disconnect and delete county
- Pricing stored separately in Supabase `unassigned_zones` table (one entry per user — backlog item to make per-county)

---

## 12. Key Function Reference

### 12.1 Auth (lines 37–175)
See Section 5.2

### 12.2 DB Adapter (lines 265–365)
See Section 6

### 12.3 Zone Layer Management
| Function | Line | Purpose |
|----------|------|---------|
| `_addZoneLayers(poly)` | ~763 | Adds fill+line Mapbox layers for one zone. |
| `_removeZoneLayers(id)` | ~776 | Removes fill+line layers and source. |
| `_restoreAllZoneLayers()` | ~783 | Calls `_addZoneLayers` for every non-unassigned polygon. |
| `_addZoneLabel(poly)` | ~818 | Creates Mapbox Marker with zone label HTML. |
| `_buildCountyPills()` | ~846 | Rebuilds county pill markers (zoomed-out view). |
| `_refreshLabelMode()` | ~946 | Shows/hides pills vs zone labels based on zoom vs COUNTY_PILL_ZOOM=7. |
| `_rebuildAllLabels()` | ~983 | `_buildCountyPills()` + `_refreshLabelMode()`. Call after any polygon change. |

### 12.4 Drawing & Zone Creation
| Function | Line | Purpose |
|----------|------|---------|
| `startDraw()` | ~1018 | Sets drawMode=polygon, shows cancel button and crosshair cursor. |
| `cancelDraw()` | ~996 | Resets all draw state, clears preview layers. |
| `undoLastDrawPoint()` | ~1006 | Removes last point; cancels if 0 points remain. |
| `_finishPolygon()` | ~1081 | Validates boundary + overlap, calls `createPolygonAuto()`. |
| `createPolygonAuto(pts, color)` | ~1160 | Creates polygon, assigns next letter, adds layers+label, persists. |
| `pointInPolygon(lat, lng, pts)` | ~2695 | Ray-casting algorithm. pts is [lng,lat] pairs. |

### 12.5 Google Sheets & Properties
| Function | Line | Purpose |
|----------|------|---------|
| `connectSheets()` | ~2483 | Full connect flow with county name check + boundary validation. |
| `disconnectSheet()` | ~2449 | Disconnect + clear properties + reset counters + close modal. |
| `saveAndSyncZone()` | ~1454 | Save pricing, assign properties, write zones + pricing to sheet. |
| `loadPropertiesFromFunction(props, ...)` | ~2700 | Filters, validates coords, stores property records in properties[]. |
| `_validatePropertiesInCounty(props, fips, countyName)` | ~1331 | Checks properties fall within county GeoJSON boundary. |
| `_finishSheetConnect({...})` | ~2580 | Saves config, sets connected, runs initial zone assignment with propCount. |
| `deleteCounty(stateAbbr, countyName, evt)` | ~2067 | Deletes all zones + clears sheet config + clears properties for county. |

### 12.6 Navigation & County Loading
| Function | Line | Purpose |
|----------|------|---------|
| `loadCounty()` | ~3052 | Loads county boundary GeoJSON, zooms map, draws boundary. |
| `loadCounties(silent)` | ~2906 | Fetches county list for state. Uses 30-day localStorage cache. |
| `loadCountyBoundaryOnly(sa, cn, co)` | ~1940 | Loads county GeoJSON without zooming. |
| `navigateToState(sa)` | ~1847 | Fetches state boundary, zooms, draws county boundaries for zones. |
| `_fetchCountyGeoJSON(fips, name)` | ~3015 | 4-level fallback: Census2020 URL1, URL2, TIGERweb, Nominatim OSM. |

### 12.7 Persistence
| Function | Line | Purpose |
|----------|------|---------|
| `persistZones()` | ~2178 | Awaits DB.saveZones + DB.saveUnassigned. Single call point. |
| `restoreZones()` | ~2212 | Loads zones from Supabase, renders sidebar, draws layers. |
| `saveAppState()` | ~3115 | Fire-and-forget DB.saveAppState with current dropdowns. |

---

## 13. CSS Design System (index.html)

### 13.1 CSS Variables
```css
--bg: #f6f7f9          /* page background */
--panel: #ffffff        /* card/modal surfaces */
--panel2: #eef0f4       /* input backgrounds */
--border: #dde1e9       /* borders */
--accent: #5b7fa6       /* primary blue */
--zone-blue: #2c5282    /* darker blue for zone text */
--accent-light: #edf2f8 /* light blue backgrounds */
--accent2: #b94040      /* red/danger */
--green: #2e8a5a
--red: #b94040
--yellow: #c49b2a
--text: #1a2332         /* primary text */
--muted: #6b7d95        /* secondary text */
--font: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif
```

### 13.2 Layout
- `body`: CSS grid, `grid-template-rows: 64px 1fr`, `grid-template-columns: 420px 1fr`
- Header: spans full width (grid-column 1/-1), 64px height
- Sidebar (`aside`): 420px wide, left column
- Map: right column, fills remaining space
- Mobile (≤700px): single column layout

### 13.3 Pill Counters
- `.zone-prop-count`: `font-size: 11px`, `background: #dce6f0`, `color: #2c5282`, `border-radius: 20px`, `padding: 1px 7px`, dynamic width (no min-width)
- `.county-zone-pill`: same style, `background: #b8cfe0`
- Both right-aligned via `margin-left: auto` on parent flex container

### 13.4 Modals
Six modals, all using `.modal-overlay` (fixed overlay) + `.modal` (content):
- `confirmModal` — generic yes/no confirmation
- `sheetsModal` — Google Sheets connection (redesigned step-based)
- `zoneEditorModal` — pricing tier editor (700px wide)
- `authModal` — email/password login + signup tabs
- `resetModal` — password reset email
- `newPasswordModal` — set new password after reset

---

## 14. Known Working Behaviors

### Session Restore Without Modal Flicker
`getSession()` is called in an async IIFE at the top of app.js before the map loads. If a session exists, `_currentUser` is set before `map.on('load')` fires, so the app initializes without ever showing the auth modal.

### Sidebar Always Renders on Login
`restoreZones()` calls `renderPolygonList()` immediately after zones are loaded. The 600ms setTimeout runs AFTER `restoreZones()` completes.

### Zone Counts After Disconnect/Delete
- Disconnect: clears properties, removes unassigned polygon, resets propCount to 0, updates stat counters immediately
- Delete county: same cleanup plus removes sheet config from Supabase
- Page refresh after disconnect: `_loadZone()` checks `_hasSheet` and loads propCount as 0 if no sheet connected

### Wrong-County Sheet Protection
Three-layer protection:
1. Raw data check before `loadPropertiesFromFunction` filters anything
2. Boundary validation blocks on fetch failure (not silent pass)
3. Modal open sanity check clears bad saved configs automatically

### showToast Duration
`showToast(msg, type, duration)` — duration defaults to 4500ms. Pass 8000 for long messages like wrong-county import blocks.

---

## 15. Apps Script (GS_Apps_Script_.gs)

Bound to each county's Google Spreadsheet. Key functions:

| Function | Purpose |
|----------|---------|
| `populateScrubbed()` | Loads LI Raw Dataset → Scrubbed and Priced with filters dialog |
| `applyFiltersAndLoad(prefs)` | Called from HTML dialog with filter preferences |
| `refreshOfferPrices(silent)` | Recalculates all offer price columns using Pricing Settings |
| `syncBlindOfferTab()` | Syncs to Blind Offer Mail Ready tab |
| `syncRangeOfferTab()` | Syncs to Range Offer Mail Ready tab |
| `removeFlaggedRows()` | Deletes yellow-highlighted rows |
| `removeLowSellerIQ()` | Deletes Low likelihood seller rows |
| `openCountyInApp()` | Opens LandValuator deep-link for this county |
| `doPost(e)` | Web app endpoint. Called by sheets-trigger-refresh Netlify function to auto-fire refreshOfferPrices(). Requires one-time deployment as web app. |
| `onOpen()` | Adds ✉️ Mailing Campaign Commands menu |

**Menu items:**
1. Load data into Scrubbed & Priced
2. Remove flagged rows (yellow)
3. Remove all Low likelihood sellers
4. Sync Blind Offer Mail Ready Tab
5. Sync Range Offer Mail Ready Tab
6. Refresh offer prices only
7. Open county in LandValuator

**appsscript.json oauthScopes:**
```json
["https://www.googleapis.com/auth/spreadsheets",
 "https://www.googleapis.com/auth/script.storage",
 "https://www.googleapis.com/auth/script.container.ui"]
```

---

## 16. Backlog (Phase 2)

| # | Feature | Notes | Est. |
|---|---------|-------|------|
| 1 | **Google OAuth login** | "Sign in with Google" as additional auth option. Supabase has built-in support. Add button to auth modal. | ~1-2 hrs |
| 2 | **Auto-refresh pricing via Netlify function** | Port `refreshOfferPrices()` into `sheets-refresh-prices.js`. Clean up `doPost`, `sheets-trigger-refresh.js`, and `GAS_REFRESH_URL` env var. Manual "6. Refresh" stays as fallback. | ~2-3 hrs |
| 3 | **Unassigned zone persistence** | Show unassigned zone as soon as first zone is drawn (before sheet connected). County pill and zone counts show dash when no sheet connected. Unassigned stays visible on disconnect. Pricing assignable before sheet connection. Requires Supabase schema change — `unassigned_zones` table currently one entry per user, needs per-county storage. | ~2-3 hrs |
| 4 | **KML parcel boundary outlines** | Viable if LandInsights can export all parcels in one KML per county (single-file-per-parcel is not viable). KML contains APN, address, owner, acreage. Would match to property records by APN and show clickable parcel outlines on map. Ray is checking with LandInsights. | ~1-2 hrs if viable |
| 5 | **Onboarding video** | Short screen recording showing how to connect Google Sheets. Embed link in connect modal. Content creation, no code. | — |
| 6 | **Email confirmation** | Currently OFF. Enable once SMTP is configured for production. | — |
| 8 | **Subscription billing** | Stripe integration. No monthly fee — 2.9% + 30¢ per transaction. | When ready to monetize |

---

## 17. Development Notes

- **No build step.** Vanilla JS and CSS. Deploy = push to GitHub main → Netlify auto-deploys.
- **File split is load-bearing.** index.html and app.js must stay separate. See Section 2 critical note.
- **Service account email** is char-code obfuscated in app.js. Decoded at DOMContentLoaded and injected into `#serviceEmailEl`. Do not store it as a plain string.
- **County key format:** `"stateAbbr|countyName"` e.g. `"MI|Newaygo"`. Used as key in sheetConfigs and _countyLayers.
- **renderPolygonList() is synchronous.** Cannot be made async. UI state reads must use `_uiStateCache` (sync) not DB calls (async).
- **GeoJSON coordinates are [lng, lat].** pointInPolygon takes (lat, lng, pts) where pts is [lng, lat] pairs — note the argument order difference.
- **Netlify production deploys cost 15 credits each** on the current credit-based plan (free tier = 300 credits/month). Batch commits before pushing.

---

*LandValuator — Synergy Land Investments — April 2026*
