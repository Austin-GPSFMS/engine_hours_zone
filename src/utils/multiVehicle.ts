/**
 * Multi-vehicle orchestration.
 *
 *   buildMultiVehicleReport(api, deviceIds, from, to, radius, onProgress)
 *     1. For each device (in concurrent batches), fetch Trip + StatusData.
 *     2. Build stops and cluster them per-device (each vehicle's zones are
 *        independent — Z1 on Vehicle A is unrelated to Z1 on Vehicle B).
 *     3. After all per-vehicle work is done, reverse-geocode every cluster
 *        center across all vehicles in ONE GetAddresses call.
 *     4. Build trip+stop segments per vehicle (now that addresses exist).
 *     5. Group segments by local-time day.
 *
 * Concurrency is capped to keep us polite to the MyGeotab rate limiter
 * even when running across hundreds of vehicles.
 */

import type {
  GeotabApi,
  GeotabDevice,
  MultiVehicleReport,
  VehicleBucket,
} from "../types";
import {
  fetchAddresses,
  fetchTripsAndEngineHours,
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
  bucket: VehicleBucket;
  /** Trips kept around so we can build segments after geocoding addresses. */
  trips: import("../types").GeotabTrip[];
  /** Engine-hours samples kept for the same reason. */
  engineHours: import("../types").GeotabStatusData[];
}

export interface BuildArgs {
  api: GeotabApi;
  deviceIds: string[];
  devicesById: Map<string, GeotabDevice>;
  fromDate: string;
  toDate: string;
  radiusMeters: number;
  /** Optional progress callback — fired after each vehicle's fetch completes. */
  onProgress?: (done: number, total: number, currentName: string) => void;
}

export async function buildMultiVehicleReport({
  api,
  deviceIds,
  devicesById,
  fromDate,
  toDate,
  radiusMeters,
  onProgress,
}: BuildArgs): Promise<MultiVehicleReport> {
  const partials: PartialBucket[] = [];
  let done = 0;

  // --- Phase 1: per-vehicle fetch + cluster (concurrent batches) ---
  for (let i = 0; i < deviceIds.length; i += CONCURRENCY) {
    const batch = deviceIds.slice(i, i + CONCURRENCY);
    const batchResults = await Promise.all(
      batch.map(async (deviceId): Promise<PartialBucket> => {
        const device = devicesById.get(deviceId);
        const deviceName = device?.name ?? deviceId;
        try {
          const { trips, engineHours } = await fetchTripsAndEngineHours(
            api,
            deviceId,
            fromDate,
            toDate
          );
          const stops = buildStops(trips, engineHours);
          const clusters = clusterStops(stops, radiusMeters);
          return {
            bucket: {
              deviceId,
              deviceName,
              days: [],
              clusters,
              totalStopEngineSeconds: 0,
              totalTripEngineSeconds: 0,
              totalStops: stops.length,
              totalTrips: trips.length,
            },
            trips,
            engineHours,
          };
        } catch (err) {
          return {
            bucket: {
              deviceId,
              deviceName,
              days: [],
              clusters: [],
              totalStopEngineSeconds: 0,
              totalTripEngineSeconds: 0,
              totalStops: 0,
              totalTrips: 0,
              error: friendlyError(err),
            },
            trips: [],
            engineHours: [],
          };
        }
      })
    );
    for (const p of batchResults) {
      partials.push(p);
      done++;
      onProgress?.(done, deviceIds.length, p.bucket.deviceName);
    }
  }

  // --- Phase 2: batched reverse-geocode across ALL clusters ---
  const allClusters = partials.flatMap((p) => p.bucket.clusters);
  if (allClusters.length > 0) {
    const addrs = await fetchAddresses(
      api,
      allClusters.map((c) => ({ lat: c.centerLat, lng: c.centerLng }))
    );
    allClusters.forEach((c, idx) => {
      c.address =
        addrs[idx] ?? `${c.centerLat.toFixed(5)}, ${c.centerLng.toFixed(5)}`;
    });
  }

  // --- Phase 3: build segments + day buckets per vehicle ---
  for (const p of partials) {
    if (p.bucket.error) continue;
    const segments = buildSegments(
      p.trips,
      p.engineHours,
      p.bucket.clusters
    );
    const days = bucketByDay(segments);
    p.bucket.days = days;
    p.bucket.totalStopEngineSeconds = days.reduce(
      (s, d) => s + d.stopEngineSeconds,
      0
    );
    p.bucket.totalTripEngineSeconds = days.reduce(
      (s, d) => s + d.tripEngineSeconds,
      0
    );
  }

  // --- Phase 4: roll up totals ---
  const vehicles = partials.map((p) => p.bucket);
  const successful = vehicles.filter((v) => !v.error);
  const totals = {
    vehicleCount: vehicles.length,
    successfulVehicleCount: successful.length,
    totalSegments: vehicles.reduce(
      (s, v) => s + v.days.reduce((d, day) => d + day.segments.length, 0),
      0
    ),
    totalStopEngineSeconds: vehicles.reduce(
      (s, v) => s + v.totalStopEngineSeconds,
      0
    ),
    totalTripEngineSeconds: vehicles.reduce(
      (s, v) => s + v.totalTripEngineSeconds,
      0
    ),
  };

  // Sort vehicles: most engine hours first (most relevant to the user).
  vehicles.sort(
    (a, b) =>
      b.totalStopEngineSeconds +
      b.totalTripEngineSeconds -
      (a.totalStopEngineSeconds + a.totalTripEngineSeconds)
  );

  return {
    fromDate,
    toDate,
    radiusMeters,
    vehicles,
    totals,
  };
}
