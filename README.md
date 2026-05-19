# Engine Hours by Zone — MyGeotab Add-In

Vite + React + TypeScript + Zenith MyGeotab add-in that breaks a vehicle's
engine hours down by where they accumulated. Stops within a chosen radius
are clustered into a single "zone"; each zone reports total engine hours
accumulated, visit count, and the chronological timeline alternates zones
with transit legs between them.

Mirrors the same stack and conventions as `advanced_report_builder`.

## Project layout

```
engine_hours_zone/
├── package.json
├── tsconfig.json
├── vite.config.ts          ← base: "" produces clean relative asset paths
├── index.html              ← Vite entry shell
├── config.json             ← MyGeotab Add-In manifest
├── .nojekyll               ← keeps GitHub Pages from filtering _-prefixed files
└── src/
    ├── main.tsx            ← lifecycle wiring (geotab.addin.engineHoursByZone)
    ├── App.tsx             ← page: toolbar + summary + zones + timeline
    ├── types.ts            ← GeotabApi, Trip, StatusData, Cluster, Stop, ZoneReport
    ├── styles.css          ← imports Zenith CSS + GPSFMS brand
    ├── api/
    │   └── geotab.ts       ← promise wrappers, fetchDevices, fetchTripsAndEngineHours, fetchAddresses
    ├── utils/
    │   ├── cluster.ts      ← haversineMeters, buildStops, clusterStops, buildTimeline
    │   ├── format.ts       ← duration/time/hours/miles helpers
    │   └── export.ts       ← multi-sheet xlsx (Metadata / Zones / Stops)
    └── components/
        ├── Summary.tsx     ← KPI strip
        ├── ZoneTable.tsx   ← sorted zones table
        └── Timeline.tsx    ← day-grouped chronological timeline
```

## Install & build (Windows PowerShell)

```powershell
cd C:\Users\User\Projects\engine_hours_zone

# A partial node_modules was left behind by a sandbox install attempt.
# Remove it so npm starts clean.
Remove-Item -Recurse -Force node_modules -ErrorAction SilentlyContinue
Remove-Item -Force package-lock.json -ErrorAction SilentlyContinue

npm install
npm run build
```

Build output lands in `dist/`. The hashed asset filenames (e.g.
`assets/index-abc123.js`) prevent MyGeotab's iframe cache from going stale
silently.

To iterate on the UI without rebuilding each time:

```powershell
npm run dev
```

Vite serves at `http://localhost:5173`. The page renders a "Standalone preview"
banner when there's no MyGeotab API present — that's expected outside MyGeotab.

## Push to GitHub & enable Pages

The repo is already on GitHub at
`https://github.com/Austin-GPSFMS/engine_hours_zone.git`.

The current repo contents (the single-file standalone HTML deployed earlier
as proof-of-concept) will be replaced by this Vite project.

```powershell
cd C:\Users\User\Projects\engine_hours_zone

git init
git branch -M main
git remote add origin https://github.com/Austin-GPSFMS/engine_hours_zone.git
git add .
git commit -m "Refactor to Vite/React/TypeScript/Zenith"
# Force-push because we're replacing the previous standalone HTML in the repo.
git push -u origin main --force
```

Then enable GitHub Pages so the `dist/` folder is served:

1. Open https://github.com/Austin-GPSFMS/engine_hours_zone/settings/pages
2. Under **Source**, choose **Deploy from a branch**
3. Branch: `main`, Folder: `/ (root)`
4. **Save**

After ~30 seconds the page will be live at:

```
https://austin-gpsfms.github.io/engine_hours_zone/dist/index.html
```

(That URL is already baked into `config.json` — no further edits needed.)

## Install in MyGeotab

1. Open `config.json` in this repo, copy the entire contents.
2. In MyGeotab → **Administration → System… → System Settings → Add-Ins**
3. Remove the previous "Engine Hours by Zone" entry if present (the
   jsDelivr-served standalone version).
4. Click **New Add-In**, paste this `config.json`, **OK**.
5. **Save** at the top of System Settings, then hard refresh (Ctrl+Shift+R).
6. Menu item appears under **Activity → Engine Hours by Zone**.

## Iteration workflow once it's live

```powershell
# edit src/...
npm run build
git add . && git commit -m "..." && git push
```

GitHub Pages picks up the new `dist/` within a minute. MyGeotab will pick up
the new bundle on the next page load — the hashed asset filenames in `dist/`
mean we don't have to bust any caches manually.

## How the algorithm works

1. **Get Trips and StatusData (engine hours diagnostic) for the device + range.**
   Done in one `multiCall`, both arrays come back sorted.
2. **Build stops.** For each consecutive trip pair, the gap from `trip.stop`
   → `trip.nextTripStart` is the stop period. We interpolate engine hours at
   both timestamps and take the delta — that's how many engine hours
   accumulated while parked there.
3. **Cluster stops.** Greedy clustering: each stop joins the first cluster
   whose centroid is within the chosen radius (default 1 mile). Centroid is
   the running mean.
4. **Reverse-geocode cluster centers.** One `GetAddresses` call returns the
   formatted addresses; failures fall back to lat/lng strings.
5. **Build the timeline.** Weave trips (transit events) and stops (zone
   events) by `tripIndex` so we render a chronological day-by-day view.

## Engine hours source

Always uses `DiagnosticEngineHoursAdjustmentId` — the calibrated cumulative
value MyGeotab's built-in reports use. Returns seconds; we divide by 3600
for display.

If a customer's devices report engine hours under a different diagnostic
(some custom integrations do), change `ENGINE_HOURS_DIAGNOSTIC_ID` in
`src/api/geotab.ts`.

## Future iterations worth queuing

- **Match against MyGeotab Zones.** Today we auto-detect zones from raw
  coordinates. Add a pass: if a cluster center falls inside an existing
  MyGeotab Zone, label the cluster with that Zone's name. Requires `Get Zone`
  + point-in-polygon.
- **Multi-vehicle mode.** One row per vehicle in the zones table; click into
  a vehicle to see its timeline.
- **Inline map.** Drop a Zenith Map component (or Leaflet) and pin each
  cluster, sized by accumulated engine hours.
- **Group filter.** Replace the single-vehicle Dropdown with a
  `GroupsFilter` so the report runs over a group of devices.
- **Pin a tag/release** for production. Use `git tag v1.0.0 && git push --tags`,
  then optionally rev `config.json` to point at the tag for cache stability
  (`https://austin-gpsfms.github.io/engine_hours_zone/dist/index.html` already
  works; pinning is only relevant if we later move to jsDelivr `@vX.Y.Z`).
