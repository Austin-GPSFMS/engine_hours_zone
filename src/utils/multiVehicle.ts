/**
 * Multi-vehicle orchestration with GLOBAL zone clustering.
 *
 * Pipeline:
 *   1. For each device (concurrent batches), fetch Trip + StatusData and
 *      build the per-vehicle stop list. Stops are tagged with deviceId.
 *   2. Pool every stop from every vehicle into a single array and cluster
 *      ONCE — so 315 US-70 gets the same Zone ID regardless of which
 *      vehicle visited it first. The result is `globalClusters`.
 *   3. Reverse-geocode every global cluster center in one GetAddresses call.
 *   4. For each vehicle: build segments using a per-vehicle lookup of
 *      tripIndex → cluster (so each segment row references the canonical
 *      global zone), then bucket segments by day, and compute per-vehicle
 *      zone stats over the cluster subset this vehicle actually visited.
 */

import type {
  Cluster,
  GeotabApi,
  GeotabDevice,
  GeotabStatusData,
  GeotabTrip,
  Metric,
  MultiVehicleReport,
  Stop,
  VehicleBucket,
  VehicleZoneStats,
} from "../types";
import {
  fetchAddresses,
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
      batch.map(async (deviceId): Promise<PartialBucket> => {
        const device = devicesById.get(deviceId);
        const deviceName = device?.name ?? deviceId;
        try {
          const { trips, statusData } = await fetchTripsAndStatus(
            api,
            deviceId,
            fromDate,
            toDate,
            metric
          );
          const stops = buildStops(deviceId, trips, statusData, metric);
          return { deviceId, deviceName, trips, statusData, stops };
        } catch (err) {
          return {
            deviceId,
            deviceName,
            trips: [],
            statusData: [],
            stops: [],
            error: friendlyError(err),
          };
        }
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
        error: p.error,
      };
    }

    // Build a per-vehicle map: tripIndex → global cluster. We walk the
    // global clusters once and pick out the stops whose deviceId matches.
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
      metric
    );
    const days = bucketByDay(segments);

    // Per-vehicle zone stats: iterate the clusters this vehicle visited
    // and aggregate over THIS vehicle's stops only.
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

  // Sort vehicles by total operating time (descending) so the busiest
  // vehicle is at the top of the in-app and Excel views.
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
