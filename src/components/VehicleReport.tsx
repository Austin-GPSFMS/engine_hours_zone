import { useState } from "react";
import type { DayBucket, Segment, VehicleBucket } from "../types";
import {
  formatDuration,
  formatTime,
  kmToMiles,
  secondsToHours,
} from "../utils/format";

/**
 * One collapsible section per vehicle. Header shows the vehicle name with
 * totals (trip hrs, stop hrs, segment count). Body lists each day as a
 * subheader followed by a flat segment table.
 */
export function VehicleReport({ vehicle }: { vehicle: VehicleBucket }) {
  const [open, setOpen] = useState(true);

  if (vehicle.error) {
    return (
      <div className="ehz-vehicle">
        <div className="ehz-vehicle-header">
          <strong>{vehicle.deviceName}</strong>
          <span style={{ color: "#b91c1c", fontSize: 13 }}>
            Failed: {vehicle.error}
          </span>
        </div>
      </div>
    );
  }

  if (vehicle.days.length === 0) {
    return (
      <div className="ehz-vehicle">
        <div className="ehz-vehicle-header">
          <strong>{vehicle.deviceName}</strong>
          <span style={{ color: "#6b7280", fontSize: 13 }}>
            No trips in range
          </span>
        </div>
      </div>
    );
  }

  return (
    <div className="ehz-vehicle">
      <button
        type="button"
        className="ehz-vehicle-header"
        onClick={() => setOpen((v) => !v)}
        aria-expanded={open}
      >
        <span className="ehz-vehicle-name">
          <span style={{ marginRight: 6 }}>{open ? "▾" : "▸"}</span>
          {vehicle.deviceName}
        </span>
        <span className="ehz-vehicle-stats">
          <span>
            {vehicle.totalTrips} trips · {vehicle.totalStops} stops ·{" "}
            {vehicle.clusters.length} zones
          </span>
          <span>
            Trip {secondsToHours(vehicle.totalTripEngineSeconds)} hrs · Stop{" "}
            <strong>
              {secondsToHours(vehicle.totalStopEngineSeconds)} hrs
            </strong>
          </span>
        </span>
      </button>
      {open && (
        <div className="ehz-vehicle-body">
          {vehicle.days.map((day) => (
            <DaySection key={day.day} day={day} />
          ))}
        </div>
      )}
    </div>
  );
}

function DaySection({ day }: { day: DayBucket }) {
  return (
    <div className="ehz-day">
      <div className="ehz-day-header">
        <span className="ehz-day-label">{day.dayLabel}</span>
        <span className="ehz-day-stats">
          Trip {secondsToHours(day.tripEngineSeconds)} hrs · Stop{" "}
          <strong>{secondsToHours(day.stopEngineSeconds)} hrs</strong>
        </span>
      </div>
      <div className="ehz-table-wrap">
        <table className="ehz-table ehz-segments">
          <thead>
            <tr>
              <th style={{ width: 130 }}>Time</th>
              <th style={{ width: 80 }}>Type</th>
              <th>Detail</th>
              <th style={{ width: 90 }}>Duration</th>
              <th style={{ width: 90 }} className="ehz-num">
                Entry EH
              </th>
              <th style={{ width: 90 }} className="ehz-num">
                Exit EH
              </th>
              <th style={{ width: 80 }} className="ehz-num">
                Δ EH
              </th>
            </tr>
          </thead>
          <tbody>
            {day.segments.map((s, i) => (
              <SegmentRow key={i} segment={s} />
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function SegmentRow({ segment }: { segment: Segment }) {
  const time = `${formatTime(segment.start)} – ${formatTime(segment.end)}`;
  const duration = formatDuration(segment.durationMs);
  const entry = secondsToHours(segment.entryEngineSeconds);
  const exit = secondsToHours(segment.exitEngineSeconds);
  const delta = secondsToHours(segment.accumulatedEngineSeconds);

  if (segment.type === "trip") {
    const from =
      segment.fromZoneId != null
        ? `${segment.fromZoneId} · ${segment.fromAddress ?? ""}`
        : null;
    const to =
      segment.toZoneId != null
        ? `${segment.toZoneId} · ${segment.toAddress ?? ""}`
        : null;
    return (
      <tr className="ehz-row-trip">
        <td>{time}</td>
        <td>
          <span className="ehz-pill ehz-pill-trip">Trip</span>
        </td>
        <td>
          <div>{kmToMiles(segment.distanceKm)} mi</div>
          {(from || to) && (
            <div className="ehz-tl-from-to">
              {from && (
                <>
                  <span className="ehz-tl-label">From:</span> {from}
                </>
              )}
              {from && to && <span style={{ margin: "0 6px" }}>→</span>}
              {to && (
                <>
                  <span className="ehz-tl-label">To:</span> {to}
                </>
              )}
            </div>
          )}
        </td>
        <td>{duration}</td>
        <td className="ehz-num">{entry}</td>
        <td className="ehz-num">{exit}</td>
        <td className="ehz-num">{delta}</td>
      </tr>
    );
  }
  return (
    <tr className="ehz-row-stop">
      <td>{time}</td>
      <td>
        <span className="ehz-pill ehz-pill-stop">Stop</span>
      </td>
      <td>
        <strong>{segment.clusterId}</strong> · {segment.address ?? "—"}
        <a
          className="ehz-maplink"
          href={`https://www.google.com/maps?q=${segment.lat},${segment.lng}`}
          target="_blank"
          rel="noopener"
        >
          map
        </a>
      </td>
      <td>{duration}</td>
      <td className="ehz-num">{entry}</td>
      <td className="ehz-num">{exit}</td>
      <td className="ehz-num">
        <strong>{delta}</strong>
      </td>
    </tr>
  );
}
