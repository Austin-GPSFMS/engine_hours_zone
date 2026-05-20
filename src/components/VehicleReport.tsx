import { useState } from "react";
import type {
  DayBucket,
  GeotabSessionInfo,
  Segment,
  VehicleBucket,
} from "../types";
import {
  formatDateTime,
  formatDuration,
  formatTime,
  kmToMiles,
  secondsToHours,
} from "../utils/format";
import { mapUrlForPoint } from "../utils/mapUrl";

interface VehicleReportProps {
  vehicle: VehicleBucket;
  /**
   * Current MyGeotab session (database + server). Used to build native map
   * URLs. When null we fall back to Google Maps.
   */
  session: GeotabSessionInfo | null;
}

/**
 * One collapsible section per vehicle. Header shows the vehicle name with
 * totals (trip hrs, stop hrs, segment count). Body lists each day as a
 * subheader followed by a flat segment table.
 */
export function VehicleReport({ vehicle, session }: VehicleReportProps) {
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
            {vehicle.zones.length} zones
          </span>
          <span>
            Trip {secondsToHours(vehicle.totalTripSeconds)} hrs · Stop{" "}
            <strong>{secondsToHours(vehicle.totalStopSeconds)} hrs</strong>
          </span>
        </span>
      </button>
      {open && (
        <div className="ehz-vehicle-body">
          {vehicle.anchor ? (
            <div className="ehz-anchor-note">
              <strong>Anchor:</strong>{" "}
              {secondsToHours(vehicle.anchor.value)} engine hrs at{" "}
              {formatDateTime(vehicle.anchor.dateTime)} — Entry/Exit values are
              absolute engine hours computed by subtracting ignition-on time
              backward from this reading.
            </div>
          ) : (
            <div className="ehz-anchor-note ehz-anchor-note-warn">
              <strong>No engine-hours anchor found.</strong> This device has
              never reported the engine-hours diagnostic (likely a 3-wire
              install). Entry/Exit values are relative cumulative ignition
              time within the report window — only Δ is a meaningful absolute.
            </div>
          )}
          {vehicle.days.map((day) => (
            <DaySection key={day.day} day={day} session={session} />
          ))}
        </div>
      )}
    </div>
  );
}

function DaySection({
  day,
  session,
}: {
  day: DayBucket;
  session: GeotabSessionInfo | null;
}) {
  return (
    <div className="ehz-day">
      <div className="ehz-day-header">
        <span className="ehz-day-label">{day.dayLabel}</span>
        <span className="ehz-day-stats">
          Trip {secondsToHours(day.tripSeconds)} hrs · Stop{" "}
          <strong>{secondsToHours(day.stopSeconds)} hrs</strong>
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
                Entry
              </th>
              <th style={{ width: 90 }} className="ehz-num">
                Exit
              </th>
              <th style={{ width: 80 }} className="ehz-num">
                Δ hrs
              </th>
            </tr>
          </thead>
          <tbody>
            {day.segments.map((s, i) => (
              <SegmentRow key={i} segment={s} session={session} />
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function SegmentRow({
  segment,
  session,
}: {
  segment: Segment;
  session: GeotabSessionInfo | null;
}) {
  const time = `${formatTime(segment.start)} – ${formatTime(segment.end)}`;
  const duration = formatDuration(segment.durationMs);
  const entry = secondsToHours(segment.entrySeconds);
  const exit = secondsToHours(segment.exitSeconds);
  const delta = secondsToHours(segment.accumulatedSeconds);

  if (segment.type === "trip") {
    const from =
      segment.fromZoneId != null
        ? `${segment.fromZoneId} · ${segment.fromAddress ?? ""}`
        : null;
    const to =
      segment.toZoneId != null
        ? `${segment.toZoneId} · ${segment.toAddress ?? ""}`
        : null;
    const mapHref = mapUrlForPoint(
      session,
      segment.toLat,
      segment.toLng,
      segment.toAddress
    );
    return (
      <tr className="ehz-row-trip">
        <td>{time}</td>
        <td>
          <span className="ehz-pill ehz-pill-trip">Trip</span>
        </td>
        <td>
          <div>
            {kmToMiles(segment.distanceKm)} mi
            {mapHref && (
              <a
                className="ehz-maplink"
                href={mapHref}
                target="_blank"
                rel="noopener"
              >
                map
              </a>
            )}
          </div>
          {(from || to) && (
            <div className="ehz-tl-from-to">
              {from && (
                <>
                  <span className="ehz-tl-label">Origin:</span> {from}
                </>
              )}
              {from && to && <span style={{ margin: "0 6px" }}>→</span>}
              {to && (
                <>
                  <span className="ehz-tl-label">Destination:</span> {to}
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

  const mapHref = mapUrlForPoint(
    session,
    segment.lat,
    segment.lng,
    segment.address
  );
  return (
    <tr className="ehz-row-stop">
      <td>{time}</td>
      <td>
        <span className="ehz-pill ehz-pill-stop">Stop</span>
      </td>
      <td>
        <strong>{segment.clusterId}</strong> · {segment.address ?? "—"}
        {mapHref && (
          <a
            className="ehz-maplink"
            href={mapHref}
            target="_blank"
            rel="noopener"
          >
            map
          </a>
        )}
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
