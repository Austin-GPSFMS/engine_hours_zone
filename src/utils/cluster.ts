/**
 * Core algorithm — stop detection, engine hours interpolation, zone clustering.
 *
 * Pure functions, no React or API dependencies. Easy to unit test in isolation
 * (when we add tests later).
 */

import type {
  Cluster,
  GeotabStatusData,
  GeotabTrip,
  Stop,
  TimelineEvent,
} from "../types";

/** 1 mile expressed in meters — used as the default cluster radius. */
export const ONE_MILE_METERS = 1609.34;

/**
 * Haversine distance between two lat/lng pairs in meters.
 * Standard great-circle approximation, accurate to a fraction of a percent.
 */
export function haversineMeters(
  lat1: number,
  lon1: number,
  lat2: number,
  lon2: number
): number {
  const R = 6371000;
  const toRad = (d: number) => (d * Math.PI) / 180;
  const dLat = toRad(lat2 - lat1);
  const dLon = toRad(lon2 - lon1);
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(lat1)) *
      Math.cos(toRad(lat2)) *
      Math.sin(dLon / 2) *
      Math.sin(dLon / 2);
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

/**
 * Linearly interpolate the engine-hours cumulative seconds value at a
 * target timestamp from a time-sorted StatusData array. When the target
 * falls outside the data, we return the nearest available reading.
 *
 * Engine hours posts periodically (every few minutes while running, sparse
 * when the engine is off). So "engine hours at exactly 10:00:00" is really
 * the interpolated value between the bracketing reported samples.
 */
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
 * For each pair of consecutive trips, the stop period runs from
 * `trip.stop` → `trip.nextTripStart`. Build a list of stops with
 * engine-hour deltas across each window.
 *
 * Trips without a `stopPoint` or without `nextTripStart` (open-ended,
 * usually the last trip in the range) are skipped.
 */
export function buildStops(
  trips: GeotabTrip[],
  engineHours: GeotabStatusData[]
): Stop[] {
  const stops: Stop[] = [];
  for (let i = 0; i < trips.length; i++) {
    const trip = trips[i];
    const sp = trip.stopPoint;
    if (!sp || sp.x == null || sp.y == null) continue;
    if (!trip.nextTripStart) continue;
    const ehStart = interpolateEngineHours(engineHours, trip.stop);
    const ehEnd = interpolateEngineHours(engineHours, trip.nextTripStart);
    const accumulated =
      ehStart != null && ehEnd != null ? Math.max(0, ehEnd - ehStart) : null;
    stops.push({
      lat: sp.y,
      lng: sp.x,
      arrive: trip.stop,
      depart: trip.nextTripStart,
      durationMs:
        new Date(trip.nextTripStart).getTime() - new Date(trip.stop).getTime(),
      engineSecondsAccumulated: accumulated,
      tripIndex: i,
    });
  }
  return stops;
}

/**
 * Greedy clustering: each stop joins the first existing cluster whose
 * centroid is within the radius, else starts a new cluster. The centroid is
 * the running mean of all stops in the cluster.
 *
 * Intentionally simple — works well for the typical fleet pattern of yards,
 * jobsites, and customer addresses. Swap for DBSCAN later if dense urban
 * routes start producing merge artifacts.
 */
export function clusterStops(
  stops: Stop[],
  radiusMeters: number = ONE_MILE_METERS
): Cluster[] {
  const clusters: Cluster[] = [];
  for (const s of stops) {
    let joined = false;
    for (const c of clusters) {
      if (
        haversineMeters(s.lat, s.lng, c.centerLat, c.centerLng) <= radiusMeters
      ) {
        c.stops.push(s);
        c.centerLat =
          (c.centerLat * (c.stops.length - 1) + s.lat) / c.stops.length;
        c.centerLng =
          (c.centerLng * (c.stops.length - 1) + s.lng) / c.stops.length;
        joined = true;
        break;
      }
    }
    if (!joined) {
      clusters.push({
        id: `Z${clusters.length + 1}`,
        centerLat: s.lat,
        centerLng: s.lng,
        stops: [s],
        address: null,
        totalEngineSeconds: 0,
        totalStoppedMs: 0,
        visits: 0,
      });
    }
  }
  // Populate totals once clustering is settled.
  for (const c of clusters) {
    let totalEh = 0;
    let totalDur = 0;
    for (const s of c.stops) {
      if (s.engineSecondsAccumulated != null) totalEh += s.engineSecondsAccumulated;
      totalDur += s.durationMs;
    }
    c.totalEngineSeconds = totalEh;
    c.totalStoppedMs = totalDur;
    c.visits = c.stops.length;
  }
  return clusters;
}

/**
 * Weave trips (transit events) and stops (zone events) into a chronological
 * timeline keyed by tripIndex. Used by the Timeline component to render the
 * day-by-day "Zone A → Transit → Zone B" view.
 */
export function buildTimeline(
  trips: GeotabTrip[],
  clusters: Cluster[]
): TimelineEvent[] {
  const stopToCluster = new Map<number, Cluster>();
  for (const c of clusters) {
    for (const s of c.stops) stopToCluster.set(s.tripIndex, c);
  }
  const events: TimelineEvent[] = [];
  for (let i = 0; i < trips.length; i++) {
    const trip = trips[i];
    events.push({
      type: "transit",
      start: trip.start,
      end: trip.stop,
      distanceKm: trip.distance ?? 0,
      durationMs:
        new Date(trip.stop).getTime() - new Date(trip.start).getTime(),
    });
    const cluster = stopToCluster.get(i);
    if (cluster) {
      const stop = cluster.stops.find((s) => s.tripIndex === i);
      if (stop) {
        events.push({
          type: "zone",
          cluster,
          stop,
          start: stop.arrive,
          end: stop.depart,
          durationMs: stop.durationMs,
          engineSeconds: stop.engineSecondsAccumulated,
        });
      }
    }
  }
  return events;
}
