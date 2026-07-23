/**
 * Cumulative-metric math for ignition and engine-hours modes.
 *
 *   engineHours mode (ignition-aware, preferred) —
 *                      Anchor to the latest DiagnosticEngineHoursAdjustmentId
 *                      reading and walk actual ignition events forward or
 *                      backward from that anchor. This produces per-segment
 *                      deltas that reflect real engine-on time rather than a
 *                      linear interpolation smeared across parked overnights,
 *                      while keeping the absolute Entry/Exit values pegged to
 *                      the calibrated diagnostic that MyGeotab's Vehicle page
 *                      displays. Used whenever we have both an anchor and
 *                      ignition events (always, since v3.6.0).
 *
 *   engineHours mode fallback —
 *                      When ignition events are somehow missing, fall back to
 *                      linear interpolation between engine-hours samples. This
 *                      preserves totals but can smear idle time into stops.
 *
 *   ignition mode with anchor —
 *                      Same math as ignition-aware engineHours: anchor to
 *                      the latest engine-hours reading, walk ignition events
 *                      to the target. Kept as an explicit mode so operators
 *                      can see the raw ignition-based number without any
 *                      engine-hours-diagnostic pegging (useful for
 *                      troubleshooting install-health issues).
 *
 *   ignition mode no anchor —
 *                      Cumulative ignition-on seconds since the first event
 *                      in the events array. This is a RELATIVE value, only
 *                      meaningful as a delta. Used only when no engine-hours
 *                      reading exists for the device (true 3-wire installs
 *                      that have never reported engine-hours data).
 */

import type { EngineHoursAnchor, GeotabStatusData, Metric } from "../types";

/**
 * The data required to compute a cumulative metric value at a moment in time.
 * `engineHoursRecords` is DiagnosticEngineHoursAdjustmentId samples;
 * `ignitionEvents` is DiagnosticIgnitionId events. Either can be empty, but
 * a fully populated bundle unlocks ignition-aware Engine Hours math.
 */
export interface MetricInputs {
  engineHoursRecords: GeotabStatusData[];
  ignitionEvents: GeotabStatusData[];
}

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
 * Absolute engine-hours seconds at `target`, computed by anchoring to the
 * latest engine-hours reading and adding/subtracting actual ignition-on time
 * between the anchor and the target. This is the core of ignition-aware
 * Engine Hours mode — it's how we avoid smearing sparse diagnostic samples
 * across long parked overnights.
 */
export function ignitionAnchoredValueAt(
  events: GeotabStatusData[],
  target: string | Date,
  anchor: EngineHoursAnchor
): number {
  const targetMs = new Date(target).getTime();
  const anchorMs = new Date(anchor.dateTime).getTime();
  if (targetMs === anchorMs) return anchor.value;
  if (targetMs < anchorMs) {
    const delta = integrateIgnitionBetween(events, targetMs, anchorMs);
    return Math.max(0, anchor.value - delta);
  }
  const delta = integrateIgnitionBetween(events, anchorMs, targetMs);
  return anchor.value + delta;
}

/**
 * Top-level helper — returns the cumulative metric value at `target`,
 * dispatching on metric, whether an anchor is available, and which datasets
 * are populated.
 *
 * For engineHours mode, we PREFER the ignition-aware path when both an
 * anchor and ignition events are on hand (this is the norm since v3.6.0),
 * falling back to sample interpolation only when ignition events are
 * unavailable. That fallback preserves totals but re-introduces the idle-
 * smear artifact on long parked stops — so we log a console warning to make
 * the degradation visible.
 */
export function metricValueAt(
  inputs: MetricInputs,
  target: string | Date,
  metric: Metric,
  anchor?: EngineHoursAnchor | null
): number | null {
  if (metric === "engineHours") {
    if (anchor && inputs.ignitionEvents.length > 0) {
      return ignitionAnchoredValueAt(inputs.ignitionEvents, target, anchor);
    }
    return interpolateEngineHours(inputs.engineHoursRecords, target);
  }
  // Ignition mode — anchored when we have engine-hours data on file, raw
  // cumulative otherwise.
  if (anchor) {
    return ignitionAnchoredValueAt(inputs.ignitionEvents, target, anchor);
  }
  return integrateIgnitionCumulative(inputs.ignitionEvents, target);
}
