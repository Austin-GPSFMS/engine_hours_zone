/**
 * Cumulative-metric math for both ignition and engine-hours modes.
 *
 *   engineHours mode — DiagnosticEngineHoursAdjustmentId is a monotonically
 *                      increasing counter in seconds. Linearly interpolate
 *                      between the bracketing reported samples to estimate
 *                      the value at any timestamp.
 *
 *   ignition    mode — DiagnosticIgnitionId reports 1 (on) and 0 (off) at
 *                      every state change (plus periodic heartbeats). To
 *                      compute "cumulative ignition-on seconds up to time T"
 *                      we walk every event in order and add (event_time -
 *                      previous_event_time) whenever the previous state
 *                      was 1, stopping at T.
 *
 * Both functions return cumulative seconds since the start of the records
 * array. Callers compute deltas (exit − entry) for "seconds during this
 * segment".
 */

import type { GeotabStatusData, Metric } from "../types";

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
 * Cumulative ignition-on seconds from the first event up to targetDate.
 *
 * Walks events in chronological order. For each consecutive pair, if the
 * previous state was 1, the gap (in seconds) gets added to the running total.
 * When we cross targetDate, we account for the partial interval up to T and
 * stop.
 *
 * The first event establishes the "current state" — anything before the
 * first event is unknown and treated as if the metric clock starts at 0.
 * In practice we fetch with a padded window so the first event is well
 * before the first segment we care about, making the integration accurate
 * over the report window.
 */
export function integrateIgnition(
  records: GeotabStatusData[],
  targetDate: string | Date
): number | null {
  if (!records || records.length === 0) return null;
  const target = new Date(targetDate).getTime();
  let total = 0;
  let lastTime: number | null = null;
  let lastState: number | null = null;

  for (const r of records) {
    const t = new Date(r.dateTime).getTime();
    if (t >= target) {
      // Crossed the target — close out the partial interval.
      if (lastTime !== null && lastState === 1) {
        total += (target - lastTime) / 1000;
      }
      return total;
    }
    if (lastTime !== null && lastState === 1) {
      total += (t - lastTime) / 1000;
    }
    lastTime = t;
    lastState = r.data;
  }
  // Target is past the last event — extend the final state to the target.
  if (lastTime !== null && lastState === 1) {
    total += (target - lastTime) / 1000;
  }
  return total;
}

/**
 * Compute cumulative metric seconds at targetDate, dispatching on the
 * active metric. Returns null when there's no data to base the answer on.
 */
export function computeMetricAt(
  records: GeotabStatusData[],
  targetDate: string | Date,
  metric: Metric
): number | null {
  if (metric === "engineHours") {
    return interpolateEngineHours(records, targetDate);
  }
  return integrateIgnition(records, targetDate);
}
