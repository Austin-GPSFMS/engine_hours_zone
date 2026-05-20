/**
 * Core algorithm — stop detection, metric computation, zone clustering,
 * segment construction, day bucketing.
 *
 * Pure functions, no React or API dependencies.
 *
 * Clustering is GLOBAL: callers pool stops across all vehicles before
 * calling clusterStops so the same physical location gets the same Zone ID
 * regardless of which vehicle reported it first.
 */

import type {
  Cluster,
  DayBucket,
  EngineHoursAnchor,
  GeotabStatusData,
  GeotabTrip,
  Metric,
  Segment,
  Stop,
} from "../types";
import { metricValueAt } from "./metric";

/** 1 mile expressed in meters — used as the default cluster radius. */
export const ONE_MILE_METERS = 1609.34;

/** Haversine distance between two lat/lng pairs in meters. */
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
 * Build the per-vehicle list of stops. Each stop sits between trip[i].stop
 * and trip[i].nextTripStart. Entry/exit metric seconds are computed against
 * the supplied StatusData (engine hours interpolation OR ignition integration,
 * dispatched on `metric`).
 *
 * The caller passes the owning deviceId so each Stop is tagged and can be
 * pooled across vehicles for global clustering.
 */
export function buildStops(
  deviceId: string,
  trips: GeotabTrip[],
  statusData: GeotabStatusData[],
  metric: Metric,
  anchor?: EngineHoursAnchor | null
): Stop[] {
  const stops: Stop[] = [];
  for (let i = 0; i < trips.length; i++) {
    const trip = trips[i];
    const sp = trip.stopPoint;
    if (!sp || sp.x == null || sp.y == null) continue;
    if (!trip.nextTripStart) continue;
    const entry = metricValueAt(statusData, trip.stop, metric, anchor);
    const exit = metricValueAt(statusData, trip.nextTripStart, metric, anchor);
    const accumulated =
      entry != null && exit != null ? Math.max(0, exit - entry) : null;
    stops.push({
      deviceId,
      lat: sp.y,
      lng: sp.x,
      arrive: trip.stop,
      depart: trip.nextTripStart,
      durationMs:
        new Date(trip.nextTripStart).getTime() - new Date(trip.stop).getTime(),
      entrySeconds: entry,
      exitSeconds: exit,
      accumulatedSeconds: accumulated,
      tripIndex: i,
    });
  }
  return stops;
}

/**
 * Greedy global clustering. Pool stops from every vehicle into a single
 * array and call this once — each cluster's stops can come from multiple
 * vehicles, and the cluster's id is stable across the whole report.
 *
 * Sort the input by arrive time first so cluster IDs roughly track which
 * zone got "discovered" earliest in the report window (cosmetic but nice).
 */
export function clusterStops(
  stops: Stop[],
  radiusMeters: number = ONE_MILE_METERS
): Cluster[] {
  const sorted = stops
    .slice()
    .sort((a, b) => new Date(a.arrive).getTime() - new Date(b.arrive).getTime());

  const clusters: Cluster[] = [];
  for (const s of sorted) {
    let joined = false;
    for (const c of clusters) {
      if (
        haversineMeters(s.lat, s.lng, c.centerLat, c.centerLng) <= radiusMeters
      ) {
        c.stops.push(s);
        c.vehicleIds.add(s.deviceId);
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
        totalSeconds: 0,
        totalStoppedMs: 0,
        visits: 0,
        vehicleIds: new Set<string>([s.deviceId]),
      });
    }
  }
  for (const c of clusters) {
    let totalSec = 0;
    let totalDur = 0;
    for (const s of c.stops) {
      if (s.accumulatedSeconds != null) totalSec += s.accumulatedSeconds;
      totalDur += s.durationMs;
    }
    c.totalSeconds = totalSec;
    c.totalStoppedMs = totalDur;
    c.visits = c.stops.length;
  }
  return clusters;
}

/**
 * Interleave a single vehicle's trips and its stops (which now reference
 * global clusters) into a chronological segment list. The cluster lookup
 * is supplied from outside so multiple vehicles can share global cluster
 * IDs / addresses.
 */
export function buildSegments(
  trips: GeotabTrip[],
  statusData: GeotabStatusData[],
  stops: Stop[],
  clusterByTripIndex: Map<number, Cluster>,
  metric: Metric,
  anchor?: EngineHoursAnchor | null
): Segment[] {
  // We still need quick access to per-stop entry/exit values keyed by trip.
  const stopByTripIndex = new Map<number, Stop>();
  for (const s of stops) stopByTripIndex.set(s.tripIndex, s);

  const segments: Segment[] = [];
  for (let i = 0; i < trips.length; i++) {
    const trip = trips[i];
    const entry = metricValueAt(statusData, trip.start, metric, anchor);
    const exit = metricValueAt(statusData, trip.stop, metric, anchor);
    const acc =
      entry != null && exit != null ? Math.max(0, exit - entry) : null;

    // Origin = stop immediately before this trip; destination = stop after.
    const fromCluster = clusterByTripIndex.get(i - 1);
    const toCluster = clusterByTripIndex.get(i);

    segments.push({
      type: "trip",
      start: trip.start,
      end: trip.stop,
      durationMs:
        new Date(trip.stop).getTime() - new Date(trip.start).getTime(),
      distanceKm: trip.distance ?? 0,
      fromZoneId: fromCluster?.id ?? null,
      fromAddress: fromCluster?.address ?? null,
      fromLat: fromCluster?.centerLat ?? null,
      fromLng: fromCluster?.centerLng ?? null,
      toZoneId: toCluster?.id ?? null,
      toAddress: toCluster?.address ?? null,
      toLat: toCluster?.centerLat ?? null,
      toLng: toCluster?.centerLng ?? null,
      entrySeconds: entry,
      exitSeconds: exit,
      accumulatedSeconds: acc,
    });

    const cluster = clusterByTripIndex.get(i);
    if (cluster) {
      const stop = stopByTripIndex.get(i);
      if (stop) {
        segments.push({
          type: "stop",
          start: stop.arrive,
          end: stop.depart,
          durationMs: stop.durationMs,
          clusterId: cluster.id,
          address: cluster.address,
          lat: stop.lat,
          lng: stop.lng,
          entrySeconds: stop.entrySeconds,
          exitSeconds: stop.exitSeconds,
          accumulatedSeconds: stop.accumulatedSeconds,
        });
      }
    }
  }
  return segments;
}

/** Group segments by local-time day. Returns days sorted ascending. */
export function bucketByDay(segments: Segment[]): DayBucket[] {
  const map = new Map<string, DayBucket>();
  for (const seg of segments) {
    const d = new Date(seg.start);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const key = `${y}-${m}-${day}`;
    let bucket = map.get(key);
    if (!bucket) {
      bucket = {
        day: key,
        dayLabel: d.toLocaleDateString(undefined, {
          weekday: "long",
          year: "numeric",
          month: "short",
          day: "numeric",
        }),
        segments: [],
        stopSeconds: 0,
        tripSeconds: 0,
      };
      map.set(key, bucket);
    }
    bucket.segments.push(seg);
    if (seg.accumulatedSeconds != null) {
      if (seg.type === "stop") bucket.stopSeconds += seg.accumulatedSeconds;
      else bucket.tripSeconds += seg.accumulatedSeconds;
    }
  }
  return Array.from(map.values()).sort((a, b) => a.day.localeCompare(b.day));
}
