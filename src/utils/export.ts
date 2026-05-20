/**
 * ExcelJS xlsx export — multi-vehicle, trip-level breakdown with global zones.
 *
 *   Sheet 1: "Segments"     — primary chronological view per vehicle/day
 *                             with logo + metadata header + Origin/Destination
 *                             + Map hyperlink into MyGeotab.
 *   Sheet 2: "Zone Summary" — global per-zone roll-up across ALL vehicles
 *                             (pivot-ready: each zone gets one row with total
 *                             visits, hours, vehicle count).
 *   Sheet 3: "Vehicle Zones" — per-vehicle × per-zone breakdown, also pivot-
 *                              friendly when filtering one vehicle at a time.
 *   Sheet 4: "Report Metadata" — full provenance block.
 *   Sheet 5: "Failed Vehicles" — only emitted when any errors occurred.
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
// Sheet 2: Zone Summary — global per-zone roll-up (pivot-ready)
// ----------------------------------------------------------------------

function writeZoneSummarySheet(
  wb: ExcelJS.Workbook,
  report: MultiVehicleReport,
  session: GeotabSessionInfo | null
) {
  const ws = wb.addWorksheet("Zone Summary");
  ws.columns = [
    { header: "Zone", key: "zone", width: 8 },
    { header: "Address", key: "address", width: 45 },
    { header: "Latitude", key: "lat", width: 12 },
    { header: "Longitude", key: "lng", width: 12 },
    { header: "Vehicles", key: "vehicles", width: 10 },
    { header: "Total Visits", key: "visits", width: 12 },
    { header: "Total Stopped (hrs)", key: "stopped", width: 18 },
    { header: "Total Hours Accumulated", key: "hours", width: 22 },
    { header: "Map", key: "map", width: 14 },
  ];
  styleHeaderRow(ws.getRow(1));
  ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 9 } };
  ws.views = [{ state: "frozen", ySplit: 1 }];

  // Sort by total hours descending so the heaviest zones bubble to the top.
  const sorted: Cluster[] = report.zones
    .slice()
    .sort((a, b) => b.totalSeconds - a.totalSeconds);

  let rowIdx = 2;
  for (const c of sorted) {
    const row = ws.getRow(rowIdx++);
    row.getCell(1).value = c.id;
    row.getCell(2).value = c.address ?? "";
    row.getCell(3).value = Number(c.centerLat.toFixed(6));
    row.getCell(4).value = Number(c.centerLng.toFixed(6));
    row.getCell(5).value = c.vehicleIds.size;
    row.getCell(6).value = c.visits;
    row.getCell(7).value = msToHours(c.totalStoppedMs);
    row.getCell(8).value = secToHours(c.totalSeconds);
    const url = mapUrlForPoint(session, c.centerLat, c.centerLng, c.address);
    if (url) {
      row.getCell(9).value = { text: "Open in Maps", hyperlink: url };
      row.getCell(9).font = LINK_FONT;
    }
  }
}

// ----------------------------------------------------------------------
// Sheet 3: Vehicle × Zone breakdown
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
