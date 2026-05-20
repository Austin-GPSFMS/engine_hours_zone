import type { MultiVehicleReport } from "../types";
import { METRIC_LABEL } from "../types";
import { formatDateTime, secondsToHours } from "../utils/format";

/**
 * Top KPI strip — rolls up across all selected vehicles. Surfaces stop hours
 * (the headline "where the time accumulated" number) and trip hours (drive
 * time), labeled with the active metric (Ignition vs Engine Hours).
 */
export function Summary({ report }: { report: MultiVehicleReport }) {
  const failures = report.vehicles.filter((v) => v.error).length;
  const metricLabel = METRIC_LABEL[report.metric];
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
          <span className="ehz-summary-label">Metric</span>
          <span className="ehz-summary-value">{metricLabel}</span>
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
          <span className="ehz-summary-label">Distinct zones</span>
          <span className="ehz-summary-value">{report.zones.length}</span>
        </div>
        <div>
          <span className="ehz-summary-label">Total segments</span>
          <span className="ehz-summary-value">{report.totals.totalSegments}</span>
        </div>
        <div>
          <span className="ehz-summary-label">Stop hours</span>
          <span className="ehz-summary-value">
            {secondsToHours(report.totals.totalStopSeconds)} hrs
          </span>
        </div>
        <div>
          <span className="ehz-summary-label">Trip hours</span>
          <span className="ehz-summary-value">
            {secondsToHours(report.totals.totalTripSeconds)} hrs
          </span>
        </div>
      </div>
    </div>
  );
}
