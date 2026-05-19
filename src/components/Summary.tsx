import type { MultiVehicleReport } from "../types";
import { formatDateTime, secondsToHours } from "../utils/format";

/**
 * Top KPI strip — rolls up across all selected vehicles. Surfaces both the
 * stop engine hours (the headline "where did the hours accumulate" number)
 * and the trip engine hours (drive time) so they're easy to compare.
 */
export function Summary({ report }: { report: MultiVehicleReport }) {
  const failures = report.vehicles.filter((v) => v.error).length;
  return (
    <div className="ehz-summary">
      <h3>Summary</h3>
      <div className="ehz-summary-grid">
        <div>
          <span className="ehz-summary-label">Period</span>
          <span className="ehz-summary-value">
            {formatDateTime(report.fromDate)} – {formatDateTime(report.toDate)}
          </span>
        </div>
        <div>
          <span className="ehz-summary-label">Vehicles</span>
          <span className="ehz-summary-value">
            {report.totals.successfulVehicleCount}
            {failures > 0 && (
              <span style={{ color: "#b91c1c", fontWeight: 400, marginLeft: 6 }}>
                ({failures} failed)
              </span>
            )}
          </span>
        </div>
        <div>
          <span className="ehz-summary-label">Total segments</span>
          <span className="ehz-summary-value">{report.totals.totalSegments}</span>
        </div>
        <div>
          <span className="ehz-summary-label">Stop engine hours</span>
          <span className="ehz-summary-value">
            {secondsToHours(report.totals.totalStopEngineSeconds)} hrs
          </span>
        </div>
        <div>
          <span className="ehz-summary-label">Trip engine hours</span>
          <span className="ehz-summary-value">
            {secondsToHours(report.totals.totalTripEngineSeconds)} hrs
          </span>
        </div>
      </div>
    </div>
  );
}
