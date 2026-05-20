/**
 * Cumulative-metric math for ignition and engine-hours modes.
 *
 *   engineHours mode — DiagnosticEngineHoursAdjustmentId is a monotonically
 *                      increasing counter in seconds. Linearly interpolate
 *                      between the bracketing reported samples.
 *
 *   ignition mode with anchor —
 *                      Take a known engine-hours reading (the anchor) at a
 *                      recent timestamp, then walk ignition on/off events
 *                      between the anchor and the requested target to
 *                      compute the engine-hours value at that moment.
 *                      When the target is OLDER than the anchor (the usual
 *                      case for historical reports), the ignition-on time
 *                      between them is subtracted from the anchor value.
 *
 *   ignition mode no anchor —
 *                      Cumulative ignition-on seconds since the first event
 *                      in the records array. This is a RELATIVE value, only
 *                      meaningful as a delta. Used only when no engine-hours
 *                      reading exists for the device (true 3-wire installs).
 */

import type { EngineHoursAnchor, GeotabStatusData, Metric } from "../types";

/** Linearly interpolate cumulative engine-hour seconds at targetDate. */
export function interpolateEngineHours(
  records: GeotabStatusData[],
  targetDate: string | Date
): number | null {
  if (!records || records.length === 0) return null;
  const target = new Date(targetDate).getTime();
  let before: GeotabStatusData | null = null;
  let after: GeotabStatusData | null = null;
  for (const r of records) {
    const t = new Date(r.dateTime).getTime();
    if (t <= target) {
      before = r;
    } else {
      after = r;
      break;
    }
  }
  if (before && after) {
    const tB = new Date(before.dateTime).getTime();
    const tA = new Date(after.dateTime).getTime();
    if (tA === tB) return before.data;
    const ratio = (target - tB) / (tA - tB);
    return before.data + (after.data - before.data) * ratio;
  }
  if (before) return before.data;
  if (after) return after.data;
  return null;
}

/**
 * Total ignition-on seconds between two timestamps. Walks events in order,
 * tracking the state at startMs (last event before startMs gives initial
 * state) and accumulating ignition-on duration until endMs.
 */
export function integrateIgnitionBetween(
  events: GeotabStatusData[],
  startMs: number,
  endMs: number
): number {
  if (startMs >= endMs) return 0;
  if (!events || events.length === 0) return 0;

  // Determine the state at startMs from the last event at or before it.
  let prevState = 0;
  for (const ev of events) {
    const t = new Date(ev.dateTime).getTime();
    if (t <= startMs) {
      prevState = ev.data;
    } else {
      break;
    }
  }

  let total = 0;
  let prevTime = startMs;

  for (const ev of events) {
    const t = new Date(ev.dateTime).getTime();
    if (t <= startMs) continue;
    if (t >= endMs) {
      if (prevState === 1) total += (endMs - prevTime) / 1000;
      return total;
    }
    if (prevState === 1) total += (t - prevTime) / 1000;
    prevState = ev.data;
    prevTime = t;
  }
  if (prevState === 1) total += (endMs - prevTime) / 1000;
  return total;
}

/**
 * Cumulative ignition-on seconds from the first event in the records to
 * the target. Only meaningful for delta calculations. Used as the fallback
 * when no anchor is available.
 */
export function integrateIgnitionCumulative(
  events: GeotabStatusData[],
  targetDate: string | Date
): number | null {
  if (!events || events.length === 0) return null;
  const firstT = new Date(events[0].dateTime).getTime();
  const targetT = new Date(targetDate).getTime();
  if (targetT <= firstT) return 0;
  return integrateIgnitionBetween(events, firstT, targetT);
}

/**
 * Top-level helper — returns the cumulative metric value at `target`,
 * dispatching on metric and (for ignition) whether an anchor is available.
 */
export function metricValueAt(
  records: GeotabStatusData[],
  target: string | Date,
  metric: Metric,
  anchor?: EngineHoursAnchor | null
): number | null {
  if (metric === "engineHours") {
    return interpolateEngineHours(records, target);
  }
  // Ignition mode
  if (anchor) {
    const targetMs = new Date(target).getTime();
    const anchorMs = new Date(anchor.dateTime).getTime();
    if (targetMs === anchorMs) return anchor.value;
    if (targetMs < anchorMs) {
      const delta = integrateIgnitionBetween(records, targetMs, anchorMs);
      return Math.max(0, anchor.value - delta);
    }
    const delta = integrateIgnitionBetween(records, anchorMs, targetMs);
    return anchor.value + delta;
  }
  return integrateIgnitionCumulative(records, target);
}
