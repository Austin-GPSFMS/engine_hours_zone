/**
 * Promise-based wrappers around the MyGeotab Add-In `api.call` interface.
 *
 * Same shape as the Advanced Report Builder's api/geotab.ts: every call goes
 * through `apiCall<T>` so we have a single place to pace requests and surface
 * typed errors.
 */

import type {
  GeotabApi,
  GeotabAddress,
  GeotabDevice,
  GeotabStatusData,
  GeotabTrip,
} from "../types";

/** Small spacing between back-to-back calls (rate-limit politeness). */
const INTER_CALL_DELAY_MS = 50;

/** The calibrated cumulative engine hours diagnostic — matches MyGeotab's UI. */
export const ENGINE_HOURS_DIAGNOSTIC_ID = "DiagnosticEngineHoursAdjustmentId";

/** Generic typed call wrapper. Resolves with the API result, rejects on failure. */
export function apiCall<T = unknown>(
  api: GeotabApi,
  method: string,
  params: unknown
): Promise<T> {
  return new Promise((resolve, reject) => {
    try {
      api.call(
        method,
        params,
        (result) => {
          setTimeout(() => resolve(result as T), INTER_CALL_DELAY_MS);
        },
        (err) => reject(err)
      );
    } catch (err) {
      reject(err);
    }
  });
}

/** Make a fetch error nice to display to humans. */
export function friendlyError(err: unknown): string {
  if (err == null) return "Unknown error.";
  if (typeof err === "string") return err;
  if (typeof err === "object") {
    const e = err as { name?: string; message?: string };
    if (e.name === "InvalidUserException" || /InvalidUser/i.test(String(err))) {
      return "Session expired — refresh the MyGeotab page to re-authenticate.";
    }
    if (e.message) return e.message;
  }
  return String(err);
}

/** Fetch active devices, sorted by name, with the synthetic NoDeviceId removed. */
export async function fetchDevices(api: GeotabApi): Promise<GeotabDevice[]> {
  const devices = await apiCall<GeotabDevice[]>(api, "Get", {
    typeName: "Device",
    resultsLimit: 5000,
  });
  return devices
    .filter((d) => d.id !== "NoDeviceId")
    .sort((a, b) => (a.name ?? "").localeCompare(b.name ?? ""));
}

/**
 * Fetch trips and engine-hours StatusData for a device over a date range.
 *
 * The StatusData window is padded by ±1 hour so we can interpolate cleanly
 * at the edges of the report window. Both arrays come back time-sorted.
 */
export async function fetchTripsAndEngineHours(
  api: GeotabApi,
  deviceId: string,
  fromDate: string,
  toDate: string
): Promise<{ trips: GeotabTrip[]; engineHours: GeotabStatusData[] }> {
  const pad = 60 * 60 * 1000;
  const statusFrom = new Date(new Date(fromDate).getTime() - pad).toISOString();
  const statusTo = new Date(new Date(toDate).getTime() + pad).toISOString();

  const [trips, engineHours] = await Promise.all([
    apiCall<GeotabTrip[]>(api, "Get", {
      typeName: "Trip",
      search: {
        deviceSearch: { id: deviceId },
        fromDate,
        toDate,
      },
      resultsLimit: 5000,
    }),
    apiCall<GeotabStatusData[]>(api, "Get", {
      typeName: "StatusData",
      search: {
        deviceSearch: { id: deviceId },
        diagnosticSearch: { id: ENGINE_HOURS_DIAGNOSTIC_ID },
        fromDate: statusFrom,
        toDate: statusTo,
      },
      resultsLimit: 50000,
    }),
  ]);

  return {
    trips: trips
      .slice()
      .sort((a, b) => new Date(a.start).getTime() - new Date(b.start).getTime()),
    engineHours: engineHours
      .slice()
      .sort(
        (a, b) => new Date(a.dateTime).getTime() - new Date(b.dateTime).getTime()
      ),
  };
}

/**
 * Reverse-geocode an array of lat/lng pairs using the MyGeotab built-in.
 * Returns the formatted-address string for each coordinate, or null if the
 * geocode failed for that index. Never throws — falls back to null per row.
 */
export async function fetchAddresses(
  api: GeotabApi,
  coords: Array<{ lat: number; lng: number }>
): Promise<Array<string | null>> {
  if (coords.length === 0) return [];
  try {
    const result = await apiCall<GeotabAddress[]>(api, "GetAddresses", {
      coordinates: coords.map((c) => ({ x: c.lng, y: c.lat })),
      movingAddresses: false,
    });
    return coords.map((_, i) => result?.[i]?.formattedAddress ?? null);
  } catch {
    return coords.map(() => null);
  }
}
