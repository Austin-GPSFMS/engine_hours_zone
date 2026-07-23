/**
 * Promise-based wrappers around the MyGeotab Add-In `api.call` interface.
 *
 * Mirrors the shape of advanced_report_builder/src/api/geotab.ts: every call
 * goes through `apiCall<T>` so we have a single place to pace requests and
 * surface typed errors.
 */

import type {
  EngineHoursAnchor,
  GeotabApi,
  GeotabAddress,
  GeotabDevice,
  GeotabGroup,
  GeotabSessionInfo,
  GeotabStatusData,
  GeotabTrip,
  LatestIgnition,
  Metric,
} from "../types";

/** Small spacing between back-to-back calls (rate-limit politeness). */
const INTER_CALL_DELAY_MS = 50;

/** The calibrated cumulative engine hours diagnostic — matches MyGeotab's UI. */
export const ENGINE_HOURS_DIAGNOSTIC_ID = "DiagnosticEngineHoursAdjustmentId";
/** Ignition state diagnostic — 1 when ignition is on, 0 when off. Available
 *  on every device including 3-wire installs that lack engine-bus data. */
export const IGNITION_DIAGNOSTIC_ID = "DiagnosticIgnitionId";

// (diagnosticIdFor removed in v3.6.0 — the pipeline always dual-fetches
//  both diagnostics so Engine Hours mode can be ignition-aware.)

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

/**
 * Resolve the current session — database slug + active server hostname.
 *
 * Uses the Add-In API's `getSession(cb)` (callback-style, no Promise) and
 * wraps it. If the API doesn't expose getSession or it times out (rare),
 * we resolve with empty strings so callers can fall back gracefully.
 */
export function fetchSession(api: GeotabApi): Promise<GeotabSessionInfo> {
  return new Promise((resolve) => {
    if (typeof api.getSession !== "function") {
      resolve({ database: "", server: "" });
      return;
    }
    let settled = false;
    const safeResolve = (s: GeotabSessionInfo) => {
      if (settled) return;
      settled = true;
      resolve(s);
    };
    // Hard timeout so we never block app render on a stuck getSession.
    setTimeout(() => safeResolve({ database: "", server: "" }), 3000);
    try {
      api.getSession((credentials, server) => {
        safeResolve({
          database: credentials?.database ?? "",
          server: server || "my.geotab.com",
        });
      });
    } catch {
      safeResolve({ database: "", server: "" });
    }
  });
}

/** Fetch the full Group list keyed by id. */
export async function fetchGroups(
  api: GeotabApi
): Promise<Map<string, GeotabGroup>> {
  const groups = await apiCall<GeotabGroup[]>(api, "Get", {
    typeName: "Group",
    resultsLimit: 5000,
  });
  return new Map(groups.map((g) => [g.id, g]));
}

/**
 * Fetch devices scoped to a set of groups. When `groupIds` is empty we
 * default to GroupCompanyId, which returns the full fleet.
 *
 * Archived devices (activeTo in the past) are excluded unless explicitly
 * requested. Sort order: name ascending.
 */
export async function fetchDevices(
  api: GeotabApi,
  groupIds: string[] = [],
  includeArchived: boolean = false
): Promise<GeotabDevice[]> {
  const search =
    groupIds.length > 0
      ? { groups: groupIds.map((id) => ({ id })) }
      : { groups: [{ id: "GroupCompanyId" }] };
  const devices = await apiCall<GeotabDevice[]>(api, "Get", {
    typeName: "Device",
    search,
    resultsLimit: 50000,
  });
  const filtered = includeArchived
    ? devices
    : devices.filter((d) => {
        if (d.id === "NoDeviceId") return false;
        if (!d.activeTo) return true;
        const t = new Date(d.activeTo).getTime();
        return isNaN(t) ? true : t > Date.now();
      });
  return filtered.sort((a, b) =>
    (a.name ?? "").localeCompare(b.name ?? "")
  );
}

/**
 * Find the most recent DiagnosticEngineHoursAdjustmentId reading for a
 * device, regardless of when it was reported. Used as the anchor for
 * ignition-mode reports so the Entry/Exit values come out as absolute
 * engine-hours readings rather than relative cumulative ignition time.
 *
 * Returns null when the device has never reported engine-hours data
 * (true 3-wire installs without engine-bus wiring). The caller falls back
 * to relative cumulative in that case.
 *
 * Strategy: try a 30-day lookback first (covers virtually every active
 * vehicle), then widen to 365 days, then 5 years. Takes the latest
 * sample by dateTime from whichever window succeeds.
 */
export async function fetchEngineHoursAnchor(
  api: GeotabApi,
  deviceId: string
): Promise<EngineHoursAnchor | null> {
  const now = Date.now();
  const windows = [
    30 * 24 * 60 * 60 * 1000,
    365 * 24 * 60 * 60 * 1000,
    5 * 365 * 24 * 60 * 60 * 1000,
  ];
  for (const lookback of windows) {
    const from = new Date(now - lookback).toISOString();
    const to = new Date(now).toISOString();
    const records = await apiCall<GeotabStatusData[]>(api, "Get", {
      typeName: "StatusData",
      search: {
        deviceSearch: { id: deviceId },
        diagnosticSearch: { id: ENGINE_HOURS_DIAGNOSTIC_ID },
        fromDate: from,
        toDate: to,
      },
      resultsLimit: 500,
    });
    if (records && records.length > 0) {
      const latest = records.reduce((acc, r) =>
        !acc || new Date(r.dateTime).getTime() > new Date(acc.dateTime).getTime()
          ? r
          : acc
      );
      return { dateTime: latest.dateTime, value: latest.data };
    }
  }
  return null;
}

/**
 * Find the most recent DiagnosticIgnitionId record for a device, regardless
 * of when it was reported. Returns the latest state (on/off) and timestamp.
 *
 * Used by the Vehicles spot-check sheet to flag install-health issues: a
 * device whose last ignition event was "on" more than 24 hours ago without
 * a corresponding "off" is almost certainly a wiring problem (loose terminal,
 * splice failure, etc.) and will inflate ignition-mode engine-hour values.
 *
 * Same staircase lookback as fetchEngineHoursAnchor — 30d → 365d → 5y — so
 * we never spend bandwidth on a 5-year window for a device that reported
 * yesterday.
 */
export async function fetchLatestIgnition(
  api: GeotabApi,
  deviceId: string
): Promise<LatestIgnition | null> {
  const now = Date.now();
  const windows = [
    30 * 24 * 60 * 60 * 1000,
    365 * 24 * 60 * 60 * 1000,
    5 * 365 * 24 * 60 * 60 * 1000,
  ];
  for (const lookback of windows) {
    const from = new Date(now - lookback).toISOString();
    const to = new Date(now).toISOString();
    const records = await apiCall<GeotabStatusData[]>(api, "Get", {
      typeName: "StatusData",
      search: {
        deviceSearch: { id: deviceId },
        diagnosticSearch: { id: IGNITION_DIAGNOSTIC_ID },
        fromDate: from,
        toDate: to,
      },
      resultsLimit: 500,
    });
    if (records && records.length > 0) {
      const latest = records.reduce((acc, r) =>
        !acc || new Date(r.dateTime).getTime() > new Date(acc.dateTime).getTime()
          ? r
          : acc
      );
      return { dateTime: latest.dateTime, on: latest.data === 1 };
    }
  }
  return null;
}

/**
 * Fetch trips + BOTH the engine-hours diagnostic and the ignition diagnostic
 * for a single device. Since v3.6.0 we always pull both so that Engine Hours
 * mode can use ignition-aware math (per-segment Δ from real ignition-on time)
 * regardless of which mode the user picked.
 *
 * Padding rules (both diagnostics use the same window):
 *  - Front pad: 24h before the report window, so we always have the last
 *    event before the window to establish initial ignition state and the
 *    last engine-hours sample bracketing the first trip.
 *  - Back pad: 1h past the report window's end. If an anchor timestamp is
 *    supplied and it falls later than that, extend through the anchor + 1h
 *    so the integration between the anchor and each segment has continuous
 *    event coverage.
 *
 * `diagnosticIdFor(metric)` is retained for callers that want the primary
 * metric only, but the main pipeline uses this dual-fetch.
 */
export async function fetchTripsAndStatus(
  api: GeotabApi,
  deviceId: string,
  fromDate: string,
  toDate: string,
  _metric: Metric,
  anchorEnd?: string | null
): Promise<{
  trips: GeotabTrip[];
  engineHoursRecords: GeotabStatusData[];
  ignitionEvents: GeotabStatusData[];
}> {
  const fromPadMs = 24 * 60 * 60 * 1000;
  const statusFrom = new Date(
    new Date(fromDate).getTime() - fromPadMs
  ).toISOString();

  // Determine the upper bound of the StatusData fetch. If an anchor is
  // supplied and it sits past the report window, extend through it so
  // ignition-anchored back-projection has continuous event coverage.
  let statusToMs = new Date(toDate).getTime() + 60 * 60 * 1000;
  if (anchorEnd) {
    const anchorMs = new Date(anchorEnd).getTime() + 60 * 60 * 1000;
    if (anchorMs > statusToMs) statusToMs = anchorMs;
  }
  const statusTo = new Date(statusToMs).toISOString();

  const [trips, engineHoursRecords, ignitionEvents] = await Promise.all([
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
    apiCall<GeotabStatusData[]>(api, "Get", {
      typeName: "StatusData",
      search: {
        deviceSearch: { id: deviceId },
        diagnosticSearch: { id: IGNITION_DIAGNOSTIC_ID },
        fromDate: statusFrom,
        toDate: statusTo,
      },
      resultsLimit: 50000,
    }),
  ]);

  const byDate = (
    a: GeotabStatusData,
    b: GeotabStatusData
  ): number =>
    new Date(a.dateTime).getTime() - new Date(b.dateTime).getTime();

  return {
    trips: trips
      .slice()
      .sort((a, b) => new Date(a.start).getTime() - new Date(b.start).getTime()),
    engineHoursRecords: engineHoursRecords.slice().sort(byDate),
    ignitionEvents: ignitionEvents.slice().sort(byDate),
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
