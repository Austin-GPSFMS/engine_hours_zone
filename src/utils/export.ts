/**
 * ExcelJS xlsx export — multi-vehicle, trip-level breakdown.
 *
 *   Sheet 1: "Segments" — primary data sheet, header includes:
 *              - GPSFMS logo (top-left)
 *              - Title, period, timezone, cluster radius, generation time
 *              - Then the flat segment table with auto-filter + frozen
 *                header & vehicle column
 *   Sheet 2: "Zones"    — per-vehicle zone roll-up
 *   Sheet 3: "Report Metadata" — full provenance block
 *   Sheet 4: "Failed Vehicles" — only added if any vehicles errored out
 *
 * Brand color: GPSFMS navy `#25477B`.
 */

import ExcelJS from "exceljs";
import logoUrl from "../assets/gpsfms-logo.png";
import type { MultiVehicleReport, VehicleBucket } from "../types";
import { formatDateTime, KM_PER_MILE } from "./format";

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

export async function exportToXlsx(report: MultiVehicleReport): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "GPSFMS Engine Hours by Zone";
  wb.created = new Date();

  // Load the logo once so we can drop it into multiple sheets if needed.
  let logoImageId: number | null = null;
  try {
    const bytes = await loadLogoBytes();
    logoImageId = wb.addImage({ buffer: bytes, extension: "png" });
  } catch (err) {
    console.warn("[EHZ] Logo embed failed (continuing without):", err);
  }

  await writeSegmentsSheet(wb, report, logoImageId);
  writeZonesSheet(wb, report);
  writeMetadataSheet(wb, report);
  writeFailuresSheet(wb, report);

  // ---------- Download ----------
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
  "From Zone",
  "From Address",
  "Zone",
  "Address",
  "Latitude",
  "Longitude",
  "Entry EH (hrs)",
  "Exit EH (hrs)",
  "Δ EH (hrs)",
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
  10, // From Zone
  40, // From Address
  10, // Zone
  40, // Address
  12, // Latitude
  12, // Longitude
  14, // Entry EH
  14, // Exit EH
  12, // Δ EH
];

/** Row at which the data header sits. Anything above is logo + metadata. */
const SEGMENTS_HEADER_ROW = 8;

async function writeSegmentsSheet(
  wb: ExcelJS.Workbook,
  report: MultiVehicleReport,
  logoImageId: number | null
) {
  const seg = wb.addWorksheet("Segments", {
    views: [{ state: "frozen", ySplit: SEGMENTS_HEADER_ROW, xSplit: 1 }],
  });

  // Column widths
  SEGMENT_COL_WIDTHS.forEach((w, i) => {
    seg.getColumn(i + 1).width = w;
  });

  // Reserve some height for the logo / metadata area
  for (let r = 1; r <= SEGMENTS_HEADER_ROW - 1; r++) {
    seg.getRow(r).height = 22;
  }

  // Logo (anchored to A1, fixed pixel extent so cell widths don't distort it).
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

  // Metadata block in rows 3-6, columns C:D
  const metaRows: Array<[string, string | number]> = [
    [
      "Period:",
      `${formatDateTime(report.fromDate)} — ${formatDateTime(report.toDate)}`,
    ],
    ["Time Zone:", userTimeZone()],
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

  // Roll-up totals on the right (cols H-I rows 3-6) so the user has the
  // headline numbers visible without scrolling to the Metadata sheet.
  const totalsRows: Array<[string, string | number]> = [
    ["Vehicles:", `${report.totals.successfulVehicleCount} of ${report.totals.vehicleCount}`],
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
  for (const v of report.vehicles) {
    if (v.error) continue;
    for (const day of v.days) {
      for (const s of day.segments) {
        const r = seg.getRow(rowIdx++);
        if (s.type === "trip") {
          r.values = [
            v.deviceName,
            day.day,
            day.dayLabel,
            "Trip",
            formatDateTime(s.start),
            formatDateTime(s.end),
            msToHours(s.durationMs),
            kmToMi(s.distanceKm),
            s.fromZoneId ?? "",
            s.fromAddress ?? "",
            // For trip rows, "Zone/Address" is where the trip ended (== next stop).
            s.toZoneId ?? "",
            s.toAddress ?? "",
            null,
            null,
            secToHours(s.entryEngineSeconds),
            secToHours(s.exitEngineSeconds),
            secToHours(s.accumulatedEngineSeconds),
          ];
        } else {
          r.values = [
            v.deviceName,
            day.day,
            day.dayLabel,
            "Stop",
            formatDateTime(s.start),
            formatDateTime(s.end),
            msToHours(s.durationMs),
            null,
            "",
            "",
            s.clusterId,
            s.address ?? "",
            Number(s.lat.toFixed(6)),
            Number(s.lng.toFixed(6)),
            secToHours(s.entryEngineSeconds),
            secToHours(s.exitEngineSeconds),
            secToHours(s.accumulatedEngineSeconds),
          ];
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

function writeZonesSheet(wb: ExcelJS.Workbook, report: MultiVehicleReport) {
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
  ];
  styleHeaderRow(zones.getRow(1));
  zones.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 8 } };
  zones.views = [{ state: "frozen", ySplit: 1 }];

  for (const v of report.vehicles) {
    if (v.error) continue;
    const sorted = v.clusters
      .slice()
      .sort((a, b) => b.totalEngineSeconds - a.totalEngineSeconds);
    for (const c of sorted) {
      zones.addRow({
        vehicle: v.deviceName,
        zone: c.id,
        address: c.address ?? "",
        lat: Number(c.centerLat.toFixed(6)),
        lng: Number(c.centerLng.toFixed(6)),
        visits: c.visits,
        stopped: msToHours(c.totalStoppedMs),
        engineHours: secToHours(c.totalEngineSeconds),
      });
    }
  }
}

// ----------------------------------------------------------------------
// Metadata sheet (full provenance) — kept as-is for archival
// ----------------------------------------------------------------------

function writeMetadataSheet(wb: ExcelJS.Workbook, report: MultiVehicleReport) {
  const meta = wb.addWorksheet("Report Metadata");
  meta.columns = [
    { header: "Field", key: "field", width: 28 },
    { header: "Value", key: "value", width: 60 },
  ];
  styleHeaderRow(meta.getRow(1));
  meta.addRows([
    { field: "Report", value: "Engine Hours by Zone" },
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
