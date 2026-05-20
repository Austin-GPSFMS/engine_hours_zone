/**
 * ExcelJS xlsx export — multi-vehicle, trip-level breakdown.
 *
 *   Sheet 1: "Segments" — primary data sheet, header includes:
 *              - GPSFMS logo (top-left)
 *              - Title, period, timezone, cluster radius, generation time
 *              - Then the flat segment table with auto-filter + frozen
 *                header & vehicle column. Origin/Destination cells combine
 *                "Zx: address" for easy scanning, the rightmost Map cell
 *                opens the location in Google Maps.
 *   Sheet 2: "Zones"    — per-vehicle zone roll-up
 *   Sheet 3: "Report Metadata" — full provenance block
 *   Sheet 4: "Failed Vehicles" — only added if any vehicles errored out
 *
 * Brand color: GPSFMS navy `#25477B`.
 */

import ExcelJS from "exceljs";
import logoUrl from "../assets/gpsfms-logo.png";
import type {
  GeotabSessionInfo,
  MultiVehicleReport,
  VehicleBucket,
} from "../types";
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

function kmToMi(km: number | null | undefined): number | null {
  return km == null ? null : Number((km / KM_PER_MILE).toFixed(2));
}

/** "Z1: 3908 Veterans Pkwy, Garner, NC" — empty string when both pieces are missing. */
function formatLocation(
  zoneId: string | null | undefined,
  address: string | null | undefined
): string {
  const z = zoneId ?? "";
  const a = address ?? "";
  if (z && a) return `${z}: ${a}`;
  return z || a;
}

/** Fetch the bundled logo asset as raw bytes for embedding in the workbook. */
async function loadLogoBytes(): Promise<ArrayBuffer> {
  const res = await fetch(logoUrl);
  return res.arrayBuffer();
}

/** Resolve the user's local IANA timezone (e.g. "America/New_York"). */
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
  writeZonesSheet(wb, report, session);
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
// Segments sheet — the primary view
// ----------------------------------------------------------------------

const SEGMENT_HEADERS = [
  "Vehicle",
  "Date",
  "Day",
  "Type",
  "Start",
  "End",
  "Duration (hrs)",
  "Distance (mi)",
  "Origin",
  "Destination",
  "Entry EH (hrs)",
  "Exit EH (hrs)",
  "Δ EH (hrs)",
  "Map",
];

const SEGMENT_COL_WIDTHS = [
  22, // Vehicle
  12, // Date
  22, // Day
  10, // Type
  20, // Start
  20, // End
  14, // Duration
  14, // Distance
  50, // Origin
  50, // Destination
  14, // Entry EH
  14, // Exit EH
  12, // Δ EH
  14, // Map
];

const SEGMENTS_HEADER_ROW = 8;

async function writeSegmentsSheet(
  wb: ExcelJS.Workbook,
  report: MultiVehicleReport,
  logoImageId: number | null,
  session: GeotabSessionInfo | null
) {
  const seg = wb.addWorksheet("Segments", {
    views: [{ state: "frozen", ySplit: SEGMENTS_HEADER_ROW, xSplit: 1 }],
  });

  // Column widths
  SEGMENT_COL_WIDTHS.forEach((w, i) => {
    seg.getColumn(i + 1).width = w;
  });

  // Origin / Destination columns wrap their text so the full address stays
  // visible without manually resizing.
  seg.getColumn(9).alignment = { wrapText: true, vertical: "top" };
  seg.getColumn(10).alignment = { wrapText: true, vertical: "top" };

  // Reserve some height for the logo / metadata area
  for (let r = 1; r <= SEGMENTS_HEADER_ROW - 1; r++) {
    seg.getRow(r).height = 22;
  }

  // Logo
  if (logoImageId != null) {
    seg.addImage(logoImageId, {
      tl: { col: 0.2, row: 0.2 },
      ext: { width: 220, height: 116 },
    });
  }

  // Title in column C row 1
  seg.mergeCells("C1:F1");
  const titleCell = seg.getCell("C1");
  titleCell.value = "Engine Hours by Zone";
  titleCell.font = TITLE_FONT;
  titleCell.alignment = { vertical: "middle" };

  // Metadata block in rows 3-6, columns C:F
  const metaRows: Array<[string, string | number]> = [
    [
      "Period:",
      `${formatDateTime(report.fromDate)} — ${formatDateTime(report.toDate)}`,
    ],
    ["Time Zone:", userTimeZone()],
    [
      "Database:",
      session?.database
        ? `${session.database} (${session.server})`
        : "(unknown — standalone export)",
    ],
    ["Cluster radius:", `${(report.radiusMeters / 1609.34).toFixed(2)} miles`],
    ["Generated:", formatDateTime(new Date())],
  ];
  metaRows.forEach(([label, value], i) => {
    const r = i + 3;
    const labelCell = seg.getCell(`C${r}`);
    labelCell.value = label;
    labelCell.font = LABEL_FONT;
    seg.mergeCells(`D${r}:F${r}`);
    seg.getCell(`D${r}`).value = value;
  });

  // Roll-up totals on the right (cols H-J rows 3-6)
  const totalsRows: Array<[string, string | number]> = [
    [
      "Vehicles:",
      `${report.totals.successfulVehicleCount} of ${report.totals.vehicleCount}`,
    ],
    ["Total segments:", report.totals.totalSegments],
    [
      "Stop engine hours:",
      `${(report.totals.totalStopEngineSeconds / 3600).toFixed(2)} hrs`,
    ],
    [
      "Trip engine hours:",
      `${(report.totals.totalTripEngineSeconds / 3600).toFixed(2)} hrs`,
    ],
  ];
  totalsRows.forEach(([label, value], i) => {
    const r = i + 3;
    const labelCell = seg.getCell(`H${r}`);
    labelCell.value = label;
    labelCell.font = LABEL_FONT;
    seg.getCell(`I${r}`).value = value;
  });

  // Column header row
  const headerRow = seg.getRow(SEGMENTS_HEADER_ROW);
  SEGMENT_HEADERS.forEach((h, i) => {
    headerRow.getCell(i + 1).value = h;
  });
  styleHeaderRow(headerRow);
  headerRow.height = 18;

  // Data rows
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
          origin = formatLocation(s.fromZoneId, s.fromAddress);
          destination = formatLocation(s.toZoneId, s.toAddress);
          // Map link points to the destination so the user can see where
          // this trip ended up.
          url = mapUrlForPoint(session, s.toLat, s.toLng, s.toAddress);
        } else {
          // Stops sit at a single location — fill both columns the same
          // way so filters like "Origin = Z1" still surface the stop.
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
        r.getCell(7).value = msToHours(s.durationMs);
        r.getCell(8).value =
          s.type === "trip" ? kmToMi(s.distanceKm) : null;
        r.getCell(9).value = origin;
        r.getCell(10).value = destination;
        r.getCell(11).value = secToHours(s.entryEngineSeconds);
        r.getCell(12).value = secToHours(s.exitEngineSeconds);
        r.getCell(13).value = secToHours(s.accumulatedEngineSeconds);

        if (url) {
          const mapCell = r.getCell(14);
          mapCell.value = { text: "Open in Maps", hyperlink: url };
          mapCell.font = LINK_FONT;
        }

        // Subtle zebra striping for readability.
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
// Zones sheet
// ----------------------------------------------------------------------

function writeZonesSheet(
  wb: ExcelJS.Workbook,
  report: MultiVehicleReport,
  session: GeotabSessionInfo | null
) {
  const zones = wb.addWorksheet("Zones");
  zones.columns = [
    { header: "Vehicle", key: "vehicle", width: 22 },
    { header: "Zone", key: "zone", width: 8 },
    { header: "Address", key: "address", width: 40 },
    { header: "Latitude", key: "lat", width: 12 },
    { header: "Longitude", key: "lng", width: 12 },
    { header: "Visits", key: "visits", width: 8 },
    { header: "Total Stopped (hrs)", key: "stopped", width: 18 },
    { header: "Engine Hours Accumulated", key: "engineHours", width: 22 },
    { header: "Map", key: "map", width: 14 },
  ];
  styleHeaderRow(zones.getRow(1));
  zones.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 9 } };
  zones.views = [{ state: "frozen", ySplit: 1 }];

  let rowIdx = 2;
  for (const v of report.vehicles) {
    if (v.error) continue;
    const sorted = v.clusters
      .slice()
      .sort((a, b) => b.totalEngineSeconds - a.totalEngineSeconds);
    for (const c of sorted) {
      const row = zones.getRow(rowIdx++);
      row.getCell(1).value = v.deviceName;
      row.getCell(2).value = c.id;
      row.getCell(3).value = c.address ?? "";
      row.getCell(4).value = Number(c.centerLat.toFixed(6));
      row.getCell(5).value = Number(c.centerLng.toFixed(6));
      row.getCell(6).value = c.visits;
      row.getCell(7).value = msToHours(c.totalStoppedMs);
      row.getCell(8).value = secToHours(c.totalEngineSeconds);
      const url = mapUrlForPoint(session, c.centerLat, c.centerLng, c.address);
      if (url) {
        row.getCell(9).value = { text: "Open in Maps", hyperlink: url };
        row.getCell(9).font = LINK_FONT;
      }
    }
  }
}

// ----------------------------------------------------------------------
// Metadata sheet (full provenance) — kept as-is for archival
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
    { field: "Database", value: session?.database ?? "(unknown)" },
    { field: "Server", value: session?.server ?? "(unknown)" },
    { field: "Time Zone", value: userTimeZone() },
    { field: "From", value: formatDateTime(report.fromDate) },
    { field: "To", value: formatDateTime(report.toDate) },
    { field: "Cluster radius (m)", value: report.radiusMeters },
    { field: "Vehicles selected", value: report.totals.vehicleCount },
    {
      field: "Vehicles with data",
      value: report.totals.successfulVehicleCount,
    },
    { field: "Total segments", value: report.totals.totalSegments },
    {
      field: "Total stop engine hours",
      value: (report.totals.totalStopEngineSeconds / 3600).toFixed(2),
    },
    {
      field: "Total trip engine hours",
      value: (report.totals.totalTripEngineSeconds / 3600).toFixed(2),
    },
    { field: "Generated", value: formatDateTime(new Date()) },
  ]);
}

// ----------------------------------------------------------------------
// Failed Vehicles (only emitted when there's content for it)
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
