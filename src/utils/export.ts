/**
 * ExcelJS xlsx export — multi-vehicle, trip-level breakdown with global zones.
 *
 *   Sheet 1: "Segments"     — primary chronological view per vehicle/day
 *                             with logo + metadata header + Origin/Destination
 *                             + Map hyperlink into MyGeotab.
 *   Sheet 2: "Vehicles"     — per-vehicle spot-check: the latest engine-hours
 *                             reading (matches the MyGeotab Asset edit page)
 *                             alongside the max Exit we report, with a
 *                             Match / Close / Review data-match flag and a
 *                             separate Install Health flag that catches
 *                             3-wire devices with a stuck-on ignition wire.
 *   Sheet 3: "Zone Summary" — per-zone breakdown grouped by vehicle: each
 *                             zone gets a bolded "Fleet total" header row
 *                             followed by one row per vehicle that visited
 *                             it (visits, stopped hrs, accumulated hrs).
 *   Sheet 4: "Vehicle Zones" — per-vehicle × per-zone breakdown, also pivot-
 *                              friendly when filtering one vehicle at a time.
 *   Sheet 5: "Report Metadata" — full provenance block.
 *   Sheet 6: "Failed Vehicles" — only emitted when any errors occurred.
 *
 * Brand color: GPSFMS navy `#25477B`.
 */

import ExcelJS from "exceljs";
import logoUrl from "../assets/gpsfms-logo.png";
import type {
  Cluster,
  GeotabSessionInfo,
  MultiVehicleReport,
  VehicleBucket,
} from "../types";
import { METRIC_LABEL } from "../types";
import { formatDateTime, KM_PER_MILE } from "./format";
import { mapUrlForPoint } from "./mapUrl";

const HEADER_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FF25477B" },
};
const HEADER_FONT: Partial<ExcelJS.Font> = {
  color: { argb: "FFFFFFFF" },
  bold: true,
};
const TITLE_FONT: Partial<ExcelJS.Font> = {
  size: 18,
  bold: true,
  color: { argb: "FF25477B" },
};
const LABEL_FONT: Partial<ExcelJS.Font> = {
  bold: true,
  color: { argb: "FF1C2B39" },
};
const LINK_FONT: Partial<ExcelJS.Font> = {
  color: { argb: "FF0084C2" },
  underline: true,
};
const STRIPE_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFF6F8FA" },
};

function styleHeaderRow(row: ExcelJS.Row) {
  row.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
  });
}

function secToHours(s: number | null | undefined): number | null {
  return s == null ? null : Number((s / 3600).toFixed(2));
}

function msToHours(ms: number | null | undefined): number | null {
  return ms == null ? null : Number((ms / 3600000).toFixed(2));
}

/**
 * Convert milliseconds to a fraction-of-a-day value that Excel can format
 * as a time (e.g. with `numFmt = "[h]:mm"`). Excel time values are
 * fractions of a 24-hour day — 0.5 = noon, 1.0 = 24 hours. Using `[h]:mm`
 * (bracketed h) keeps the hour count from rolling over after 24h, so a
 * 36-hour stop renders as `36:00` instead of `12:00`.
 */
function msToExcelTime(ms: number | null | undefined): number | null {
  if (ms == null) return null;
  return ms / (1000 * 60 * 60 * 24);
}

function kmToMi(km: number | null | undefined): number | null {
  return km == null ? null : Number((km / KM_PER_MILE).toFixed(2));
}

function formatLocation(
  zoneId: string | null | undefined,
  address: string | null | undefined
): string {
  const z = zoneId ?? "";
  const a = address ?? "";
  if (z && a) return `${z}: ${a}`;
  return z || a;
}

async function loadLogoBytes(): Promise<ArrayBuffer> {
  const res = await fetch(logoUrl);
  return res.arrayBuffer();
}

function userTimeZone(): string {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "(unknown)";
  } catch {
    return "(unknown)";
  }
}

export async function exportToXlsx(
  report: MultiVehicleReport,
  session: GeotabSessionInfo | null
): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "GPSFMS Engine Hours by Zone";
  wb.created = new Date();

  let logoImageId: number | null = null;
  try {
    const bytes = await loadLogoBytes();
    logoImageId = wb.addImage({ buffer: bytes, extension: "png" });
  } catch (err) {
    console.warn("[EHZ] Logo embed failed (continuing without):", err);
  }

  await writeSegmentsSheet(wb, report, logoImageId, session);
  writeVehiclesSheet(wb, report);
  writeZoneSummarySheet(wb, report, session);
  writeVehicleZonesSheet(wb, report, session);
  writeMetadataSheet(wb, report, session);
  writeFailuresSheet(wb, report);

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const dateRange = `${report.fromDate.slice(0, 10)}_to_${report.toDate.slice(0, 10)}`;
  a.download = `engine-hours-by-zone-${dateRange}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ----------------------------------------------------------------------
// Sheet 1: Segments — primary chronological view
// ----------------------------------------------------------------------

const SEGMENT_HEADERS = [
  "Vehicle",
  "Date",
  "Day",
  "Type",
  "Start",
  "End",
  "Duration (hh:mm)",
  "Distance (mi)",
  "Origin",
  "Destination",
  "Entry (hrs)",
  "Exit (hrs)",
  "Δ (hrs)",
  "Map",
];

const SEGMENT_COL_WIDTHS = [22, 12, 22, 10, 20, 20, 14, 14, 50, 50, 14, 14, 12, 14];
const SEGMENTS_HEADER_ROW = 8;
/** Excel time format that doesn't roll over at 24 hours. */
const DURATION_FMT = "[h]:mm";
/** Origin label shown for trips whose predecessor falls outside the report. */
const BEFORE_REPORT_LABEL = "(before report period)";

async function writeSegmentsSheet(
  wb: ExcelJS.Workbook,
  report: MultiVehicleReport,
  logoImageId: number | null,
  session: GeotabSessionInfo | null
) {
  const seg = wb.addWorksheet("Segments", {
    views: [{ state: "frozen", ySplit: SEGMENTS_HEADER_ROW, xSplit: 1 }],
  });

  SEGMENT_COL_WIDTHS.forEach((w, i) => {
    seg.getColumn(i + 1).width = w;
  });
  // Origin / Destination: text wrap so long addresses stay readable.
  seg.getColumn(9).alignment = { wrapText: true, vertical: "top" };
  seg.getColumn(10).alignment = { wrapText: true, vertical: "top" };
  // Duration column rendered as Excel time (hh:mm, no 24h rollover).
  seg.getColumn(7).numFmt = DURATION_FMT;

  for (let r = 1; r <= SEGMENTS_HEADER_ROW - 1; r++) {
    seg.getRow(r).height = 22;
  }

  if (logoImageId != null) {
    seg.addImage(logoImageId, {
      tl: { col: 0.2, row: 0.2 },
      ext: { width: 220, height: 116 },
    });
  }

  seg.mergeCells("C1:F1");
  const titleCell = seg.getCell("C1");
  titleCell.value = "Engine Hours by Zone";
  titleCell.font = TITLE_FONT;
  titleCell.alignment = { vertical: "middle" };

  // 5 rows (3-7) — keep row 8 free for the column header so its merges
  // don't collide with the data table headers. "Generated:" still appears
  // on the Report Metadata sheet for full provenance.
  const metaRows: Array<[string, string | number]> = [
    [
      "Period:",
      `${formatDateTime(report.fromDate)} — ${formatDateTime(report.toDate)}`,
    ],
    ["Time Zone:", userTimeZone()],
    ["Metric:", METRIC_LABEL[report.metric]],
    [
      "Database:",
      session?.database
        ? `${session.database} (${session.server})`
        : "(unknown — standalone export)",
    ],
    ["Cluster radius:", `${(report.radiusMeters / 1609.34).toFixed(2)} miles`],
  ];
  metaRows.forEach(([label, value], i) => {
    const r = i + 3;
    const labelCell = seg.getCell(`C${r}`);
    labelCell.value = label;
    labelCell.font = LABEL_FONT;
    seg.mergeCells(`D${r}:F${r}`);
    seg.getCell(`D${r}`).value = value;
  });

  const totalsRows: Array<[string, string | number]> = [
    [
      "Vehicles:",
      `${report.totals.successfulVehicleCount} of ${report.totals.vehicleCount}`,
    ],
    ["Distinct zones:", report.zones.length],
    ["Total segments:", report.totals.totalSegments],
    [
      "Stop hours:",
      `${(report.totals.totalStopSeconds / 3600).toFixed(2)} hrs`,
    ],
    [
      "Trip hours:",
      `${(report.totals.totalTripSeconds / 3600).toFixed(2)} hrs`,
    ],
  ];
  totalsRows.forEach(([label, value], i) => {
    const r = i + 3;
    const labelCell = seg.getCell(`H${r}`);
    labelCell.value = label;
    labelCell.font = LABEL_FONT;
    seg.getCell(`I${r}`).value = value;
  });

  const headerRow = seg.getRow(SEGMENTS_HEADER_ROW);
  SEGMENT_HEADERS.forEach((h, i) => {
    headerRow.getCell(i + 1).value = h;
  });
  styleHeaderRow(headerRow);
  headerRow.height = 18;

  let rowIdx = SEGMENTS_HEADER_ROW + 1;
  let zebra = false;
  for (const v of report.vehicles) {
    if (v.error) continue;
    for (const day of v.days) {
      for (const s of day.segments) {
        const r = seg.getRow(rowIdx++);
        zebra = !zebra;

        let origin = "";
        let destination = "";
        let url: string | null = null;

        if (s.type === "trip") {
          // First-trip-in-window case: predecessor stop is outside our data
          // visibility, so label the origin explicitly rather than leaving
          // it blank.
          origin =
            s.fromZoneId != null
              ? formatLocation(s.fromZoneId, s.fromAddress)
              : BEFORE_REPORT_LABEL;
          destination = formatLocation(s.toZoneId, s.toAddress);
          url = mapUrlForPoint(session, s.toLat, s.toLng, s.toAddress);
        } else {
          const loc = formatLocation(s.clusterId, s.address);
          origin = loc;
          destination = loc;
          url = mapUrlForPoint(session, s.lat, s.lng, s.address);
        }

        r.getCell(1).value = v.deviceName;
        r.getCell(2).value = day.day;
        r.getCell(3).value = day.dayLabel;
        r.getCell(4).value = s.type === "trip" ? "Trip" : "Stop";
        r.getCell(5).value = formatDateTime(s.start);
        r.getCell(6).value = formatDateTime(s.end);
        // Duration: Excel time value formatted via column's [h]:mm numFmt.
        r.getCell(7).value = msToExcelTime(s.durationMs);
        r.getCell(8).value =
          s.type === "trip" ? kmToMi(s.distanceKm) : null;
        r.getCell(9).value = origin;
        r.getCell(10).value = destination;
        r.getCell(11).value = secToHours(s.entrySeconds);
        r.getCell(12).value = secToHours(s.exitSeconds);
        r.getCell(13).value = secToHours(s.accumulatedSeconds);

        if (url) {
          const mapCell = r.getCell(14);
          mapCell.value = { text: "Open in Maps", hyperlink: url };
          mapCell.font = LINK_FONT;
        }

        if (zebra) {
          for (let c = 1; c <= SEGMENT_HEADERS.length; c++) {
            const cell = r.getCell(c);
            if (!cell.fill) cell.fill = STRIPE_FILL;
          }
        }
      }
    }
  }

  seg.autoFilter = {
    from: { row: SEGMENTS_HEADER_ROW, column: 1 },
    to: { row: SEGMENTS_HEADER_ROW, column: SEGMENT_HEADERS.length },
  };
}

// ----------------------------------------------------------------------
// Sheet 2: Vehicles — per-vehicle spot-check against the Asset page
// ----------------------------------------------------------------------

/**
 * One row per selected vehicle. Surfaces:
 *  - The latest `DiagnosticEngineHoursAdjustmentId` reading (same value the
 *    MyGeotab Asset edit page shows) and our computed max Exit, so an
 *    operator can verify the report against the platform in a glance.
 *  - The latest `DiagnosticIgnitionId` event and an install-health flag
 *    that catches 3-wire devices with a stuck-on ignition (the kind of
 *    wiring issue that inflates ignition-mode engine-hours numbers).
 *
 * Data-match status thresholds:
 *   |Δ| < 0.10 hrs  → "Match"   (effectively identical)
 *   |Δ| < 1.00 hrs  → "Close"   (within an hour — expected timing drift)
 *   otherwise        → "Review"  (worth a manual look)
 *
 * Install-health thresholds (on the latest ignition event):
 *   off                        → "OK"           (clean — engine is off)
 *   on AND age < 24 h          → "OK"           (running normally right now)
 *   on AND age >= 24 h         → "Stuck-on Xd"  (likely wiring failure)
 *   no ignition data on file   → "No ignition data"
 */
const STUCK_ON_THRESHOLD_HOURS = 24;

function writeVehiclesSheet(
  wb: ExcelJS.Workbook,
  report: MultiVehicleReport
) {
  const ws = wb.addWorksheet("Vehicles");
  ws.columns = [
    { header: "Vehicle", key: "vehicle", width: 24 },
    { header: "Asset Engine Hours (hrs)", key: "assetHrs", width: 22 },
    { header: "Asset Reading Time", key: "readTime", width: 22 },
    { header: "Max Exit — this report (hrs)", key: "maxExit", width: 28 },
    { header: "Δ (report − asset, hrs)", key: "delta", width: 22 },
    { header: "Data Match", key: "status", width: 14 },
    { header: "Segments", key: "segs", width: 10 },
    { header: "Trips", key: "trips", width: 8 },
    { header: "Stops", key: "stops", width: 8 },
    { header: "Stop hrs", key: "stopHrs", width: 12 },
    { header: "Trip hrs", key: "tripHrs", width: 12 },
    { header: "Last Ignition Event", key: "ignTime", width: 22 },
    { header: "Last Ignition State", key: "ignState", width: 16 },
    { header: "Install Health", key: "install", width: 18 },
    { header: "Error", key: "error", width: 40 },
  ];
  const COL_COUNT = 15;
  styleHeaderRow(ws.getRow(1));
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: COL_COUNT },
  };
  ws.views = [{ state: "frozen", ySplit: 1, xSplit: 1 }];

  const vehicles = report.vehicles
    .slice()
    .sort((a, b) => a.deviceName.localeCompare(b.deviceName));

  let rowIdx = 2;
  let zebra = false;
  for (const v of vehicles) {
    const row = ws.getRow(rowIdx++);
    zebra = !zebra;
    row.getCell(1).value = v.deviceName;

    // Install-health and last-ignition columns apply even to error rows,
    // because we may still have fetched the ignition event before failing.
    const ignTime = v.latestIgnition
      ? formatDateTime(v.latestIgnition.dateTime)
      : "(none on record)";
    const ignState = v.latestIgnition
      ? v.latestIgnition.on
        ? "On"
        : "Off"
      : "Unknown";
    const { label: installLabel, severity: installSeverity } =
      classifyInstallHealth(v.latestIgnition);

    if (v.error) {
      row.getCell(6).value = "Error";
      row.getCell(12).value = ignTime;
      row.getCell(13).value = ignState;
      row.getCell(14).value = installLabel;
      row.getCell(15).value = v.error;
      paintInstallCell(row.getCell(14), installSeverity);
      if (zebra) stripeRow(row, COL_COUNT);
      continue;
    }

    const assetHrs = v.anchor ? secToHours(v.anchor.value) : null;
    const readTime = v.anchor
      ? formatDateTime(v.anchor.dateTime)
      : "(no engine-hours reading)";

    let maxExitSec: number | null = null;
    for (const day of v.days) {
      for (const s of day.segments) {
        if (s.exitSeconds != null) {
          if (maxExitSec == null || s.exitSeconds > maxExitSec) {
            maxExitSec = s.exitSeconds;
          }
        }
      }
    }
    const maxExitHrs = secToHours(maxExitSec);

    let delta: number | null = null;
    let status: string;
    if (assetHrs == null) {
      status = "No anchor";
    } else if (maxExitHrs == null) {
      status = "No segments";
    } else {
      delta = Number((maxExitHrs - assetHrs).toFixed(2));
      const abs = Math.abs(delta);
      if (abs < 0.1) status = "Match";
      else if (abs < 1.0) status = "Close";
      else status = "Review";
    }

    row.getCell(2).value = assetHrs;
    row.getCell(3).value = readTime;
    row.getCell(4).value = maxExitHrs;
    row.getCell(5).value = delta;
    row.getCell(6).value = status;
    row.getCell(7).value = v.days.reduce(
      (sum, d) => sum + d.segments.length,
      0
    );
    row.getCell(8).value = v.totalTrips;
    row.getCell(9).value = v.totalStops;
    row.getCell(10).value = Number((v.totalStopSeconds / 3600).toFixed(2));
    row.getCell(11).value = Number((v.totalTripSeconds / 3600).toFixed(2));
    row.getCell(12).value = ignTime;
    row.getCell(13).value = ignState;
    row.getCell(14).value = installLabel;

    // Soft color hint on the Data Match cell so mismatches catch the eye.
    const statusCell = row.getCell(6);
    if (status === "Match") {
      statusCell.font = { color: { argb: "FF15803D" }, bold: true };
    } else if (status === "Close") {
      statusCell.font = { color: { argb: "FFB45309" }, bold: true };
    } else if (status === "Review") {
      statusCell.font = { color: { argb: "FFB91C1C" }, bold: true };
    }

    paintInstallCell(row.getCell(14), installSeverity);

    if (zebra) stripeRow(row, COL_COUNT);
  }
}

/** Severity of an install-health classification. Drives cell coloring. */
type InstallSeverity = "ok" | "warn" | "bad" | "neutral";

/**
 * Decide whether the device's latest ignition event indicates a healthy
 * install or a stuck-on wiring failure. See the threshold doc on
 * writeVehiclesSheet for the rules.
 */
function classifyInstallHealth(
  latest: { dateTime: string; on: boolean } | null | undefined
): { label: string; severity: InstallSeverity } {
  if (!latest) {
    return { label: "No ignition data", severity: "neutral" };
  }
  if (!latest.on) {
    return { label: "OK", severity: "ok" };
  }
  const ageMs = Date.now() - new Date(latest.dateTime).getTime();
  const ageHours = ageMs / (1000 * 60 * 60);
  if (ageHours < STUCK_ON_THRESHOLD_HOURS) {
    return { label: "OK", severity: "ok" };
  }
  const days = Math.floor(ageHours / 24);
  const remHours = Math.floor(ageHours - days * 24);
  const ageLabel = days >= 1 ? `${days}d ${remHours}h` : `${Math.floor(ageHours)}h`;
  return {
    label: `Stuck-on (${ageLabel})`,
    severity: ageHours >= 7 * 24 ? "bad" : "warn",
  };
}

function paintInstallCell(cell: ExcelJS.Cell, severity: InstallSeverity) {
  if (severity === "ok") {
    cell.font = { color: { argb: "FF15803D" }, bold: true };
  } else if (severity === "warn") {
    cell.font = { color: { argb: "FFB45309" }, bold: true };
  } else if (severity === "bad") {
    cell.font = { color: { argb: "FFB91C1C" }, bold: true };
  }
}

function stripeRow(row: ExcelJS.Row, colCount: number) {
  for (let c = 1; c <= colCount; c++) {
    const cell = row.getCell(c);
    if (!cell.fill) cell.fill = STRIPE_FILL;
  }
}

// ----------------------------------------------------------------------
// Sheet 3: Zone Summary — per-zone breakdown grouped by vehicle
// ----------------------------------------------------------------------

/** Bolded "Fleet total" header row per zone — subtle GPSFMS-tinted band. */
const ZONE_HEADER_FILL: ExcelJS.Fill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "FFEAF1FA" },
};

/**
 * Per-zone breakdown. Each zone gets a bolded "Fleet total" row with the
 * roll-up numbers followed by one row per vehicle that visited the zone
 * (visits, stopped hrs, accumulated hrs for that vehicle at this zone
 * specifically). Zone metadata (id/address/lat/lng) is repeated on every
 * row so pivot tables and filters still work cleanly.
 */
function writeZoneSummarySheet(
  wb: ExcelJS.Workbook,
  report: MultiVehicleReport,
  session: GeotabSessionInfo | null
) {
  const ws = wb.addWorksheet("Zone Summary");
  ws.columns = [
    { header: "Zone", key: "zone", width: 8 },
    { header: "Address", key: "address", width: 45 },
    { header: "Vehicle", key: "vehicle", width: 26 },
    { header: "Visits", key: "visits", width: 10 },
    { header: "Total Stopped (hrs)", key: "stopped", width: 18 },
    { header: "Hours Accumulated", key: "hours", width: 20 },
    { header: "Latitude", key: "lat", width: 12 },
    { header: "Longitude", key: "lng", width: 12 },
    { header: "Map", key: "map", width: 14 },
  ];
  const COL_COUNT = 9;
  styleHeaderRow(ws.getRow(1));
  ws.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: COL_COUNT },
  };
  ws.views = [{ state: "frozen", ySplit: 1, xSplit: 1 }];

  const deviceNameById = new Map<string, string>();
  for (const v of report.vehicles) {
    deviceNameById.set(v.deviceId, v.deviceName);
  }

  // Heaviest zones bubble to the top.
  const sortedZones: Cluster[] = report.zones
    .slice()
    .sort((a, b) => b.totalSeconds - a.totalSeconds);

  let rowIdx = 2;
  for (const c of sortedZones) {
    // Aggregate this zone's stops by deviceId.
    const byVehicle = new Map<
      string,
      { visits: number; stoppedMs: number; seconds: number }
    >();
    for (const s of c.stops) {
      let bucket = byVehicle.get(s.deviceId);
      if (!bucket) {
        bucket = { visits: 0, stoppedMs: 0, seconds: 0 };
        byVehicle.set(s.deviceId, bucket);
      }
      bucket.visits += 1;
      bucket.stoppedMs += s.durationMs;
      if (s.accumulatedSeconds != null) bucket.seconds += s.accumulatedSeconds;
    }

    const mapUrl = mapUrlForPoint(
      session,
      c.centerLat,
      c.centerLng,
      c.address
    );

    // Fleet-total header row — bolded and tinted so the zone groups are
    // visually distinct when scrolling through the sheet.
    const headerRow = ws.getRow(rowIdx++);
    headerRow.getCell(1).value = c.id;
    headerRow.getCell(2).value = c.address ?? "";
    headerRow.getCell(3).value = `Fleet total (${byVehicle.size} vehicle${
      byVehicle.size === 1 ? "" : "s"
    })`;
    headerRow.getCell(4).value = c.visits;
    headerRow.getCell(5).value = msToHours(c.totalStoppedMs);
    headerRow.getCell(6).value = secToHours(c.totalSeconds);
    headerRow.getCell(7).value = Number(c.centerLat.toFixed(6));
    headerRow.getCell(8).value = Number(c.centerLng.toFixed(6));
    if (mapUrl) {
      headerRow.getCell(9).value = { text: "Open in Maps", hyperlink: mapUrl };
      headerRow.getCell(9).font = { ...LINK_FONT, bold: true };
    }
    for (let col = 1; col <= COL_COUNT; col++) {
      const cell = headerRow.getCell(col);
      if (col !== 9) {
        // Preserve the LINK_FONT on the map cell; bold the rest.
        cell.font = { ...(cell.font ?? {}), bold: true };
      }
      cell.fill = ZONE_HEADER_FILL;
    }

    // Per-vehicle rows sorted by accumulated hours desc so the biggest
    // contributors sit right under the fleet-total row.
    const vehicleRows = Array.from(byVehicle.entries())
      .map(([deviceId, stats]) => ({
        deviceId,
        deviceName: deviceNameById.get(deviceId) ?? deviceId,
        ...stats,
      }))
      .sort((a, b) => b.seconds - a.seconds);

    for (const v of vehicleRows) {
      const row = ws.getRow(rowIdx++);
      row.getCell(1).value = c.id;
      row.getCell(2).value = c.address ?? "";
      row.getCell(3).value = v.deviceName;
      row.getCell(4).value = v.visits;
      row.getCell(5).value = msToHours(v.stoppedMs);
      row.getCell(6).value = secToHours(v.seconds);
      row.getCell(7).value = Number(c.centerLat.toFixed(6));
      row.getCell(8).value = Number(c.centerLng.toFixed(6));
      // Skip the map link on per-vehicle rows — it's identical to the
      // header row's link and adds visual noise on zones with many vehicles.
    }
  }
}

// ----------------------------------------------------------------------
// Sheet 4: Vehicle Zones — vehicle-first × zone breakdown (mirror of
// Sheet 3, sorted from the other direction for pivot flexibility)
// ----------------------------------------------------------------------

function writeVehicleZonesSheet(
  wb: ExcelJS.Workbook,
  report: MultiVehicleReport,
  session: GeotabSessionInfo | null
) {
  const ws = wb.addWorksheet("Vehicle Zones");
  ws.columns = [
    { header: "Vehicle", key: "vehicle", width: 22 },
    { header: "Zone", key: "zone", width: 8 },
    { header: "Address", key: "address", width: 45 },
    { header: "Latitude", key: "lat", width: 12 },
    { header: "Longitude", key: "lng", width: 12 },
    { header: "Visits", key: "visits", width: 8 },
    { header: "Total Stopped (hrs)", key: "stopped", width: 18 },
    { header: "Hours Accumulated", key: "hours", width: 22 },
    { header: "Map", key: "map", width: 14 },
  ];
  styleHeaderRow(ws.getRow(1));
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 9 } };
  ws.views = [{ state: "frozen", ySplit: 1, xSplit: 1 }];

  let rowIdx = 2;
  for (const v of report.vehicles) {
    if (v.error) continue;
    const sorted = v.zones
      .slice()
      .sort((a, b) => b.totalSeconds - a.totalSeconds);
    for (const z of sorted) {
      const row = ws.getRow(rowIdx++);
      row.getCell(1).value = v.deviceName;
      row.getCell(2).value = z.zone.id;
      row.getCell(3).value = z.zone.address ?? "";
      row.getCell(4).value = Number(z.zone.centerLat.toFixed(6));
      row.getCell(5).value = Number(z.zone.centerLng.toFixed(6));
      row.getCell(6).value = z.visits;
      row.getCell(7).value = msToHours(z.totalStoppedMs);
      row.getCell(8).value = secToHours(z.totalSeconds);
      const url = mapUrlForPoint(
        session,
        z.zone.centerLat,
        z.zone.centerLng,
        z.zone.address
      );
      if (url) {
        row.getCell(9).value = { text: "Open in Maps", hyperlink: url };
        row.getCell(9).font = LINK_FONT;
      }
    }
  }
}

// ----------------------------------------------------------------------
// Sheet 4: Metadata
// ----------------------------------------------------------------------

function writeMetadataSheet(
  wb: ExcelJS.Workbook,
  report: MultiVehicleReport,
  session: GeotabSessionInfo | null
) {
  const meta = wb.addWorksheet("Report Metadata");
  meta.columns = [
    { header: "Field", key: "field", width: 28 },
    { header: "Value", key: "value", width: 60 },
  ];
  styleHeaderRow(meta.getRow(1));
  meta.addRows([
    { field: "Report", value: "Engine Hours by Zone" },
    { field: "Metric", value: METRIC_LABEL[report.metric] },
    { field: "Database", value: session?.database ?? "(unknown)" },
    { field: "Server", value: session?.server ?? "(unknown)" },
    { field: "Time Zone", value: userTimeZone() },
    { field: "From", value: formatDateTime(report.fromDate) },
    { field: "To", value: formatDateTime(report.toDate) },
    { field: "Cluster radius (m)", value: report.radiusMeters },
    { field: "Distinct zones", value: report.zones.length },
    { field: "Vehicles selected", value: report.totals.vehicleCount },
    {
      field: "Vehicles with data",
      value: report.totals.successfulVehicleCount,
    },
    { field: "Total segments", value: report.totals.totalSegments },
    {
      field: "Total stop hours",
      value: (report.totals.totalStopSeconds / 3600).toFixed(2),
    },
    {
      field: "Total trip hours",
      value: (report.totals.totalTripSeconds / 3600).toFixed(2),
    },
    { field: "Generated", value: formatDateTime(new Date()) },
  ]);
}

// ----------------------------------------------------------------------
// Sheet 5: Failed Vehicles (only emitted if there's content)
// ----------------------------------------------------------------------

function writeFailuresSheet(wb: ExcelJS.Workbook, report: MultiVehicleReport) {
  const failed: VehicleBucket[] = report.vehicles.filter((v) => v.error);
  if (failed.length === 0) return;
  const errs = wb.addWorksheet("Failed Vehicles");
  errs.columns = [
    { header: "Vehicle", key: "vehicle", width: 22 },
    { header: "Error", key: "error", width: 60 },
  ];
  styleHeaderRow(errs.getRow(1));
  for (const v of failed) {
    errs.addRow({ vehicle: v.deviceName, error: v.error ?? "" });
  }
}
