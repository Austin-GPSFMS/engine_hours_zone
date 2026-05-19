/**
 * ExcelJS xlsx export — multi-vehicle, trip-level breakdown.
 *
 *   Sheet 1: "Report Metadata" — period, vehicle counts, top-line totals
 *   Sheet 2: "Segments"        — every trip + stop, flat rows, vehicle/day columns
 *   Sheet 3: "Zones"           — per-vehicle zone roll-up
 *
 * Brand color: GPSFMS navy `#25477B` (matches the Advanced Report Builder).
 */

import ExcelJS from "exceljs";
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

export async function exportToXlsx(report: MultiVehicleReport): Promise<void> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "GPSFMS Engine Hours by Zone";
  wb.created = new Date();

  // ---------- Sheet 1: Metadata ----------
  const meta = wb.addWorksheet("Report Metadata");
  meta.columns = [
    { header: "Field", key: "field", width: 28 },
    { header: "Value", key: "value", width: 60 },
  ];
  styleHeaderRow(meta.getRow(1));
  meta.addRows([
    { field: "From", value: formatDateTime(report.fromDate) },
    { field: "To", value: formatDateTime(report.toDate) },
    { field: "Cluster radius (m)", value: report.radiusMeters },
    {
      field: "Vehicles selected",
      value: report.totals.vehicleCount,
    },
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

  // ---------- Sheet 2: Segments ----------
  const seg = wb.addWorksheet("Segments");
  seg.columns = [
    { header: "Vehicle", key: "vehicle", width: 22 },
    { header: "Date", key: "date", width: 12 },
    { header: "Day", key: "dayLabel", width: 22 },
    { header: "Type", key: "type", width: 10 },
    { header: "Start", key: "start", width: 20 },
    { header: "End", key: "end", width: 20 },
    { header: "Duration (hrs)", key: "duration", width: 14 },
    { header: "Distance (mi)", key: "miles", width: 14 },
    { header: "Zone", key: "zone", width: 8 },
    { header: "Address", key: "address", width: 40 },
    { header: "Latitude", key: "lat", width: 12 },
    { header: "Longitude", key: "lng", width: 12 },
    { header: "Entry EH (hrs)", key: "entry", width: 14 },
    { header: "Exit EH (hrs)", key: "exit", width: 14 },
    { header: "Δ EH (hrs)", key: "delta", width: 12 },
  ];
  styleHeaderRow(seg.getRow(1));
  seg.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 15 } };
  seg.views = [{ state: "frozen", ySplit: 1, xSplit: 1 }];

  for (const v of report.vehicles) {
    if (v.error) continue;
    for (const day of v.days) {
      for (const s of day.segments) {
        if (s.type === "trip") {
          seg.addRow({
            vehicle: v.deviceName,
            date: day.day,
            dayLabel: day.dayLabel,
            type: "Trip",
            start: formatDateTime(s.start),
            end: formatDateTime(s.end),
            duration: msToHours(s.durationMs),
            miles: kmToMi(s.distanceKm),
            zone: null,
            address: null,
            lat: null,
            lng: null,
            entry: secToHours(s.entryEngineSeconds),
            exit: secToHours(s.exitEngineSeconds),
            delta: secToHours(s.accumulatedEngineSeconds),
          });
        } else {
          seg.addRow({
            vehicle: v.deviceName,
            date: day.day,
            dayLabel: day.dayLabel,
            type: "Stop",
            start: formatDateTime(s.start),
            end: formatDateTime(s.end),
            duration: msToHours(s.durationMs),
            miles: null,
            zone: s.clusterId,
            address: s.address ?? null,
            lat: Number(s.lat.toFixed(6)),
            lng: Number(s.lng.toFixed(6)),
            entry: secToHours(s.entryEngineSeconds),
            exit: secToHours(s.exitEngineSeconds),
            delta: secToHours(s.accumulatedEngineSeconds),
          });
        }
      }
    }
  }

  // ---------- Sheet 3: Zones (per-vehicle roll-up) ----------
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

  // ---------- Sheet 4: Failed vehicles (only if any) ----------
  const failed: VehicleBucket[] = report.vehicles.filter((v) => v.error);
  if (failed.length > 0) {
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
