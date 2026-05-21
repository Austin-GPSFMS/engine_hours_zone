/**
 * Multi-vehicle orchestration with GLOBAL zone clustering and an optional
 * per-vehicle engine-hours anchor for ignition mode.
 *
 * Pipeline:
 *   1. For each device (concurrent batches):
 *        a. In ignition mode, fetch the latest engine-hours reading as the
 *           anchor — runs alongside the rest of the fetch.
 *        b. Fetch Trip + ignition-or-engine-hours StatusData. The status
 *           fetch's upper bound is extended through the anchor when present.
 *        c. Build the per-vehicle stop list using metricValueAt (which
 *           subtracts ignition time backward from the anchor for absolute
 *           engine-hours-equivalent values).
 *   2. Pool every stop from every vehicle and cluster ONCE so Zone IDs are
 *      shared across the fleet.
 *   3. Reverse-geocode every global cluster center.
 *   4. For each vehicle: build segments (passing the same anchor), bucket
 *      by day, derive per-vehicle zone stats.
 */

import type {
  Cluster,
  EngineHoursAnchor,
  GeotabApi,
  GeotabDevice,
  GeotabStatusData,
  GeotabTrip,
  LatestIgnition,
  Metric,
  MultiVehicleReport,
  Stop,
  VehicleBucket,
  VehicleZoneStats,
} from "../types";
import {
  fetchAddresses,
  fetchEngineHoursAnchor,
  fetchLatestIgnition,
  fetchTripsAndStatus,
  friendlyError,
} from "../api/geotab";
import {
  bucketByDay,
  buildSegments,
  buildStops,
  clusterStops,
} from "./cluster";

const CONCURRENCY = 3;

interface PartialBucket {
  deviceId: string;
  deviceName: string;
  trips: GeotabTrip[];
  statusData: GeotabStatusData[];
  stops: Stop[];
  anchor: EngineHoursAnchor | null;
  latestIgnition: LatestIgnition | null;
  error?: string;
}

export interface BuildArgs {
  api: GeotabApi;
  deviceIds: string[];
  devicesById: Map<string, GeotabDevice>;
  fromDate: string;
  toDate: string;
  radiusMeters: number;
  metric: Metric;
  onProgress?: (done: number, total: number, currentName: string) => void;
}

/**
 * Geotab occasionally emits Trip records with distance == 0 AND start ==
 * stop — usually triggered by an ignition-on event without movement, or
 * other internal trip-generator edge cases. These add nothing but noise
 * to the report (zero-duration, zero-distance, no real Origin). Filter
 * them out before they hit the segment builder. Tiny-but-real trips
 * (e.g. 0.01 mi over 15 s) are kept since at least one condition fails.
 */
function isMeaningfulTrip(t: GeotabTrip): boolean {
  const dist = t.distance ?? 0;
  if (dist > 0) return true;
  const startMs = new Date(t.start).getTime();
  const stopMs = new Date(t.stop).getTime();
  return stopMs - startMs >= 1000;
}

async function processDevice(
  api: GeotabApi,
  deviceId: string,
  deviceName: string,
  fromDate: string,
  toDate: string,
  metric: Metric
): Promise<PartialBucket> {
  try {
    // Run both lookups in parallel — they're independent, so we save a
    // round-trip per device. The anchor (engine-hours adjustment) backs
    // the spot-check column; the latest ignition event backs the install-
    // health column that surfaces 3-wire installs stuck "on".
    const [anchor, latestIgnition] = await Promise.all([
      fetchEngineHoursAnchor(api, deviceId),
      fetchLatestIgnition(api, deviceId),
    ]);

    const { trips, statusData } = await fetchTripsAndStatus(
      api,
      deviceId,
      fromDate,
      toDate,
      metric,
      anchor?.dateTime
    );
    const meaningfulTrips = trips.filter(isMeaningfulTrip);
    const stops = buildStops(
      deviceId,
      meaningfulTrips,
      statusData,
      metric,
      anchor
    );

    return {
      deviceId,
      deviceName,
      trips: meaningfulTrips,
      statusData,
      stops,
      anchor,
      latestIgnition,
    };
  } catch (err) {
    return {
      deviceId,
      deviceName,
      trips: [],
      statusData: [],
      stops: [],
      anchor: null,
      latestIgnition: null,
      error: friendlyError(err),
    };
  }
}

export async function buildMultiVehicleReport({
  api,
  deviceIds,
  devicesById,
  fromDate,
  toDate,
  radiusMeters,
  metric,
  onProgress,
}: BuildArgs): Promise<MultiVehicleReport> {
  // --- Phase 1: per-vehicle fetch + buildStops ---
  const partials: PartialBucket[] = [];
  let done = 0;

  for (let i = 0; i < deviceIds.length; i += CONCURRENCY) {
    const batch = deviceIds.slice(i, i + CONCURRENCY);
    const batchResults = await Promise.all(
      batch.map(async (deviceId) => {
        const device = devicesById.get(deviceId);
        const deviceName = device?.name ?? deviceId;
        return processDevice(
          api,
          deviceId,
          deviceName,
          fromDate,
          toDate,
          metric
        );
      })
    );
    for (const p of batchResults) {
      partials.push(p);
      done++;
      onProgress?.(done, deviceIds.length, p.deviceName);
    }
  }

  // --- Phase 2: pool stops, cluster globally ---
  const allStops: Stop[] = partials.flatMap((p) => p.stops);
  const globalClusters = clusterStops(allStops, radiusMeters);

  // --- Phase 3: batched reverse-geocode across global clusters ---
  if (globalClusters.length > 0) {
    const addrs = await fetchAddresses(
      api,
      globalClusters.map((c) => ({ lat: c.centerLat, lng: c.centerLng }))
    );
    globalClusters.forEach((c, i) => {
      c.address =
        addrs[i] ?? `${c.centerLat.toFixed(5)}, ${c.centerLng.toFixed(5)}`;
    });
  }

  // --- Phase 4: per-vehicle segments + day buckets + zone stats ---
  const vehicles: VehicleBucket[] = partials.map((p) => {
    if (p.error) {
      return {
        deviceId: p.deviceId,
        deviceName: p.deviceName,
        days: [],
        zones: [],
        totalStopSeconds: 0,
        totalTripSeconds: 0,
        totalStops: 0,
        totalTrips: 0,
        anchor: p.anchor,
        latestIgnition: p.latestIgnition,
        error: p.error,
      };
    }

    const clusterByTripIndex = new Map<number, Cluster>();
    for (const c of globalClusters) {
      for (const s of c.stops) {
        if (s.deviceId === p.deviceId) {
          clusterByTripIndex.set(s.tripIndex, c);
        }
      }
    }

    const segments = buildSegments(
      p.trips,
      p.statusData,
      p.stops,
      clusterByTripIndex,
      metric,
      p.anchor
    );
    const days = bucketByDay(segments);

    const zoneStats: VehicleZoneStats[] = [];
    for (const c of globalClusters) {
      const mine = c.stops.filter((s) => s.deviceId === p.deviceId);
      if (mine.length === 0) continue;
      let sec = 0;
      let dur = 0;
      for (const s of mine) {
        if (s.accumulatedSeconds != null) sec += s.accumulatedSeconds;
        dur += s.durationMs;
      }
      zoneStats.push({
        zone: c,
        visits: mine.length,
        totalSeconds: sec,
        totalStoppedMs: dur,
      });
    }

    return {
      deviceId: p.deviceId,
      deviceName: p.deviceName,
      days,
      zones: zoneStats,
      totalStopSeconds: days.reduce((s, d) => s + d.stopSeconds, 0),
      totalTripSeconds: days.reduce((s, d) => s + d.tripSeconds, 0),
      totalStops: p.stops.length,
      totalTrips: p.trips.length,
      anchor: p.anchor,
      latestIgnition: p.latestIgnition,
    };
  });

  // --- Phase 5: top-line rollups ---
  const successful = vehicles.filter((v) => !v.error);
  const totals = {
    vehicleCount: vehicles.length,
    successfulVehicleCount: successful.length,
    totalSegments: vehicles.reduce(
      (s, v) => s + v.days.reduce((d, day) => d + day.segments.length, 0),
      0
    ),
    totalStopSeconds: vehicles.reduce((s, v) => s + v.totalStopSeconds, 0),
    totalTripSeconds: vehicles.reduce((s, v) => s + v.totalTripSeconds, 0),
  };

  vehicles.sort(
    (a, b) =>
      b.totalStopSeconds +
      b.totalTripSeconds -
      (a.totalStopSeconds + a.totalTripSeconds)
  );

  return {
    fromDate,
    toDate,
    radiusMeters,
    metric,
    zones: globalClusters,
    vehicles,
    totals,
  };
}
