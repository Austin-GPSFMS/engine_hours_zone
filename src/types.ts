/**
 * Domain types for Engine Hours by Zone.
 *
 * The MyGeotab Add-In API is loosely typed (no official SDK types for the
 * callback surface). We narrow at the boundary in api/geotab.ts.
 */

export interface GeotabApi {
  call: (
    method: string,
    params: unknown,
    success: (result: unknown) => void,
    failure: (err: unknown) => void
  ) => void;
  /**
   * Returns the current credentials and the active server hostname.
   * The callback gets `(credentials, server)` — credentials.database is
   * the URL slug (e.g. "gpsfms_pilot") and server is the host (e.g.
   * "my.geotab.com" or a partner-specific server).
   */
  getSession?: (
    callback: (
      credentials: { database: string; userName: string; sessionId: string },
      server: string
    ) => void
  ) => void;
}

export interface GeotabSessionInfo {
  database: string;
  server: string;
}

export interface GeotabPageState {
  getGroupFilter?: (cb: (groups: unknown) => void) => void;
  setGroupFilter?: (groups: unknown) => void;
  getState?: (cb: (state: unknown) => void) => void;
  setState?: (state: unknown) => void;
  translate?: (key: string) => string;
}

export interface GeotabDevice {
  id: string;
  name?: string;
  serialNumber?: string;
  licensePlate?: string;
  activeFrom?: string;
  activeTo?: string;
  groups?: Array<{ id: string }>;
}

export interface GeotabGroup {
  id: string;
  name?: string;
  children?: Array<{ id: string }>;
}

/** A Geotab Coordinate object — note `x` is longitude, `y` is latitude. */
export interface GeotabCoordinate {
  x: number;
  y: number;
}

/** Subset of Trip fields we use. */
export interface GeotabTrip {
  id?: string;
  start: string;
  stop: string;
  nextTripStart?: string;
  stopPoint?: GeotabCoordinate;
  distance?: number;
}

/**
 * Subset of StatusData fields we use.
 *
 * `id` is null on synthetic interpolated samples that Geotab returns when
 * the query asks for interpolation. We use it to distinguish reported vs.
 * synthesized values in the engine-hours mode if we ever opt into that.
 */
export interface GeotabStatusData {
  id?: string | null;
  dateTime: string;
  /**
   * For DiagnosticEngineHoursAdjustmentId this is cumulative engine seconds.
   * For DiagnosticIgnitionId this is the ignition STATE — 1 (on) or 0 (off).
   */
  data: number;
}

/** A reverse-geocoded address from the GetAddresses call. */
export interface GeotabAddress {
  formattedAddress?: string;
}

/**
 * Which underlying signal we're using to measure operating time.
 *
 * - `ignition` — DiagnosticIgnitionId, integrated across on/off events.
 *                Works on every device including 3-wire GO Rugged installs
 *                that don't read the engine bus. This is the default.
 * - `engineHours` — DiagnosticEngineHoursAdjustmentId, the calibrated
 *                   cumulative value MyGeotab's UI uses. Requires the engine
 *                   bus to be wired (J1939 / J1708).
 */
export type Metric = "ignition" | "engineHours";

export const METRIC_LABEL: Record<Metric, string> = {
  ignition: "Ignition Time",
  engineHours: "Engine Hours",
};

/**
 * A known engine-hours reading used to anchor ignition-mode values to an
 * absolute engine-hours scale. `value` is cumulative engine seconds (the
 * raw API unit for DiagnosticEngineHoursAdjustmentId).
 */
export interface EngineHoursAnchor {
  dateTime: string;
  value: number;
}

/**
 * A single stop period (gap between two consecutive trips) on a device.
 * Carries both the entry and exit cumulative metric readings so we can
 * show "200.45 → 203.55" alongside the delta.
 */
export interface Stop {
  /** Owner device — tags the stop so we can pool stops globally for
   *  cross-vehicle clustering while still computing per-vehicle stats. */
  deviceId: string;
  lat: number;
  lng: number;
  arrive: string;
  depart: string;
  durationMs: number;
  /** Cumulative metric seconds at arrive time. */
  entrySeconds: number | null;
  /** Cumulative metric seconds at depart time. */
  exitSeconds: number | null;
  /** exit - entry, clamped at 0. */
  accumulatedSeconds: number | null;
  tripIndex: number;
}

/** A cluster of stops within the configured radius — what we call a "zone". */
export interface Cluster {
  id: string;
  centerLat: number;
  centerLng: number;
  stops: Stop[];
  address: string | null;
  /** Global totals — sum across all vehicles that visited this zone. */
  totalSeconds: number;
  totalStoppedMs: number;
  visits: number;
  /** Unique device IDs of vehicles that have a stop in this cluster. */
  vehicleIds: Set<string>;
}

// -----------------------------------------------------------------------
// Segments — trip + stop chronology built off the global clusters
// -----------------------------------------------------------------------

export interface TripSegment {
  type: "trip";
  start: string;
  end: string;
  durationMs: number;
  distanceKm: number;
  /** Zone the vehicle departed from. */
  fromZoneId: string | null;
  fromAddress: string | null;
  fromLat: number | null;
  fromLng: number | null;
  /** Zone the vehicle arrived at. */
  toZoneId: string | null;
  toAddress: string | null;
  toLat: number | null;
  toLng: number | null;
  entrySeconds: number | null;
  exitSeconds: number | null;
  accumulatedSeconds: number | null;
}

export interface StopSegment {
  type: "stop";
  start: string;
  end: string;
  durationMs: number;
  clusterId: string;
  address: string | null;
  lat: number;
  lng: number;
  entrySeconds: number | null;
  exitSeconds: number | null;
  accumulatedSeconds: number | null;
}

export type Segment = TripSegment | StopSegment;

export interface DayBucket {
  day: string;
  dayLabel: string;
  segments: Segment[];
  stopSeconds: number;
  tripSeconds: number;
}

/** Per-vehicle stats AT a specific global zone. */
export interface VehicleZoneStats {
  /** Reference to the global cluster. */
  zone: Cluster;
  visits: number;
  totalSeconds: number;
  totalStoppedMs: number;
}

export interface VehicleBucket {
  deviceId: string;
  deviceName: string;
  days: DayBucket[];
  /** Subset of report.zones this vehicle visited, with per-vehicle stats. */
  zones: VehicleZoneStats[];
  totalStopSeconds: number;
  totalTripSeconds: number;
  totalStops: number;
  totalTrips: number;
  /**
   * For ignition mode: the engine-hours reading used to anchor this
   * vehicle's absolute hour values. Null when no engine data is available
   * (true 3-wire installs) — in that case Entry/Exit are relative cumulative.
   */
  anchor?: EngineHoursAnchor | null;
  error?: string;
}

export interface MultiVehicleReport {
  fromDate: string;
  toDate: string;
  radiusMeters: number;
  metric: Metric;
  /** Global zones (shared zone IDs across all vehicles). */
  zones: Cluster[];
  vehicles: VehicleBucket[];
  totals: {
    vehicleCount: number;
    successfulVehicleCount: number;
    totalSegments: number;
    totalStopSeconds: number;
    totalTripSeconds: number;
  };
}
