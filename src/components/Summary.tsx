import type { ZoneReport } from "../types";
import { formatDateTime, secondsToHours } from "../utils/format";

/**
 * Compact KPI strip at the top of the results — period, distinct zones,
 * total visits, total engine hours accumulated in-zone.
 */
export function Summary({ report }: { report: ZoneReport }) {
  return (
    <div className="ehz-summary">
      <h3>{report.deviceName}</h3>
      <div className="ehz-summary-grid">
        <div>
          <span className="ehz-summary-label">Period</span>
          <span className="ehz-summary-value">
            {formatDateTime(report.fromDate)} – {formatDateTime(report.toDate)}
          </span>
        </div>
        <div>
          <span className="ehz-summary-label">Distinct zones</span>
          <span className="ehz-summary-value">{report.clusters.length}</span>
        </div>
        <div>
          <span className="ehz-summary-label">Total visits</span>
          <span className="ehz-summary-value">{report.totals.totalVisits}</span>
        </div>
        <div>
          <span className="ehz-summary-label">Engine hours in zones</span>
          <span className="ehz-summary-value">
            {secondsToHours(report.totals.totalZoneEngineSeconds)} hrs
          </span>
        </div>
      </div>
    </div>
  );
}
