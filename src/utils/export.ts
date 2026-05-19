/**
 * ExcelJS xlsx export — mirrors the styling approach used in the
 * Advanced Report Builder so workbooks look the same across our add-ins.
 *
 *   Sheet 1: "Report Metadata" — period, vehicle, totals
 *   Sheet 2: "Zones" — one row per cluster, sorted by engine hours
 *   Sheet 3: "Stops" — every individual visit (joinable back to a zone)
 *
 * Brand color: GPSFMS navy `#25477B`.
 */

import ExcelJS from "exceljs";
import type { Cluster, ZoneReport } from "../types";
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

export async function exportToXlsx(report: ZoneReport): Promise<void> {
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
    { field: "Vehicle", value: report.deviceName },
    { field: "From", value: formatDateTime(report.fromDate) },
    { field: "To", value: formatDateTime(report.toDate) },
    { field: "Distinct zones", value: report.clusters.length },
    { field: "Total visits", value: report.totals.totalVisits },
    {
      field: "Total engine hours in zones",
      value: (report.totals.totalZoneEngineSeconds / 3600).toFixed(2),
    },
    { field: "Generated", value: formatDateTime(new Date()) },
  ]);

  // ---------- Sheet 2: Zones ----------
  const zones = wb.addWorksheet("Zones");
  zones.columns = [
    { header: "Zone", key: "zone", width: 8 },
    { header: "Address", key: "address", width: 50 },
    { header: "Latitude", key: "lat", width: 12 },
    { header: "Longitude", key: "lng", width: 12 },
    { header: "Visits", key: "visits", width: 8 },
    { header: "Total Stopped (hrs)", key: "stopped", width: 18 },
    { header: "Engine Hours Accumulated", key: "engineHours", width: 22 },
  ];
  styleHeaderRow(zones.getRow(1));
  zones.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 7 } };
  zones.views = [{ state: "frozen", ySplit: 1 }];

  const sortedClusters: Cluster[] = report.clusters
    .slice()
    .sort((a, b) => b.totalEngineSeconds - a.totalEngineSeconds);
  for (const c of sortedClusters) {
    zones.addRow({
      zone: c.id,
      address: c.address ?? "",
      lat: Number(c.centerLat.toFixed(6)),
      lng: Number(c.centerLng.toFixed(6)),
      visits: c.visits,
      stopped: Number((c.totalStoppedMs / 3600000).toFixed(2)),
      engineHours: Number((c.totalEngineSeconds / 3600).toFixed(2)),
    });
  }

  // ---------- Sheet 3: Stops (individual visits) ----------
  const stopsSheet = wb.addWorksheet("Stops");
  stopsSheet.columns = [
    { header: "Zone", key: "zone", width: 8 },
    { header: "Address", key: "address", width: 50 },
    { header: "Arrive", key: "arrive", width: 22 },
    { header: "Depart", key: "depart", width: 22 },
    { header: "Stopped (hrs)", key: "stopped", width: 14 },
    { header: "Engine Hours Accumulated", key: "engineHours", width: 22 },
    { header: "Transit miles to next stop", key: "milesNext", width: 22 },
  ];
  styleHeaderRow(stopsSheet.getRow(1));
  stopsSheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: 7 } };
  stopsSheet.views = [{ state: "frozen", ySplit: 1 }];

  // Walk events in order to also pull transit distance to the next stop.
  let pendingTransitKm: number | null = null;
  for (const ev of report.events) {
    if (ev.type === "transit") {
      pendingTransitKm = ev.distanceKm;
      continue;
    }
    stopsSheet.addRow({
      zone: ev.cluster.id,
      address: ev.cluster.address ?? "",
      arrive: formatDateTime(ev.start),
      depart: formatDateTime(ev.end),
      stopped: Number((ev.durationMs / 3600000).toFixed(2)),
      engineHours:
        ev.engineSeconds != null
          ? Number((ev.engineSeconds / 3600).toFixed(2))
          : null,
      milesNext:
        pendingTransitKm != null
          ? Number((pendingTransitKm / KM_PER_MILE).toFixed(1))
          : null,
    });
    pendingTransitKm = null;
  }

  // ---------- Download ----------
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  const safeName = report.deviceName.replace(/[^a-z0-9]/gi, "_");
  a.download = `engine-hours-by-zone-${safeName}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
