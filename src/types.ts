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

/** Subset of StatusData fields we use. */
export interface GeotabStatusData {
  id?: string;
  dateTime: string;
  /** For DiagnosticEngineHoursAdjustmentId this is cumulative engine seconds. */
  data: number;
}

/** A reverse-geocoded address from the GetAddresses call. */
export interface GeotabAddress {
  formattedAddress?: string;
}

/**
 * A single stop period (gap between two consecutive trips) on a device.
 * Carries both the entry and exit cumulative engine-hours readings so we
 * can show "200.45 → 203.55" alongside the delta.
 */
export interface Stop {
  lat: number;
  lng: number;
  arrive: string;
  depart: string;
  durationMs: number;
  /** Cumulative engine seconds at arrive time (interpolated). */
  entryEngineSeconds: number | null;
  /** Cumulative engine seconds at depart time (interpolated). */
  exitEngineSeconds: number | null;
  /** exit - entry, clamped at 0. */
  engineSecondsAccumulated: number | null;
  tripIndex: number;
}

/** A cluster of stops within the configured radius — what we call a "zone". */
export interface Cluster {
  id: string;
  centerLat: number;
  centerLng: number;
  stops: Stop[];
  address: string | null;
  totalEngineSeconds: number;
  totalStoppedMs: number;
  visits: number;
}

// -----------------------------------------------------------------------
// Trip-level segment model — the primary view post-v1.1
// -----------------------------------------------------------------------

/** A driving segment between two trip endpoints. */
export interface TripSegment {
  type: "trip";
  start: string;
  end: string;
  durationMs: number;
  distanceKm: number;
  /** Zone the vehicle departed from — the previous stop's cluster. */
  fromZoneId: string | null;
  fromAddress: string | null;
  /** Zone the vehicle arrived at — the next stop's cluster. */
  toZoneId: string | null;
  toAddress: string | null;
  entryEngineSeconds: number | null;
  exitEngineSeconds: number | null;
  accumulatedEngineSeconds: number | null;
}

/** A stop segment — the vehicle is parked at a zone for a period of time. */
export interface StopSegment {
  type: "stop";
  start: string;
  end: string;
  durationMs: number;
  /** Cluster ID this stop falls into (Z1, Z2, …). */
  clusterId: string;
  /** Reverse-geocoded address, or null if geocode failed. */
  address: string | null;
  lat: number;
  lng: number;
  entryEngineSeconds: number | null;
  exitEngineSeconds: number | null;
  accumulatedEngineSeconds: number | null;
}

export type Segment = TripSegment | StopSegment;

export interface DayBucket {
  /** ISO date "2026-05-17" used as map key + sort key. */
  day: string;
  /** Human-readable label, e.g. "Sunday, May 17, 2026". */
  dayLabel: string;
  segments: Segment[];
  /** Sum of accumulated engine seconds across stop segments on this day. */
  stopEngineSeconds: number;
  /** Sum of accumulated engine seconds across trip segments on this day. */
  tripEngineSeconds: number;
}

export interface VehicleBucket {
  deviceId: string;
  deviceName: string;
  days: DayBucket[];
  clusters: Cluster[];
  /** Total stop engine seconds across all days (= sum of cluster totals). */
  totalStopEngineSeconds: number;
  /** Total trip engine seconds across all days. */
  totalTripEngineSeconds: number;
  totalStops: number;
  totalTrips: number;
  /** If this vehicle failed entirely, the error message; otherwise undefined. */
  error?: string;
}

export interface MultiVehicleReport {
  fromDate: string;
  toDate: string;
  /** Cluster radius (m) used to generate this report. */
  radiusMeters: number;
  vehicles: VehicleBucket[];
  totals: {
    vehicleCount: number;
    successfulVehicleCount: number;
    totalSegments: number;
    totalStopEngineSeconds: number;
    totalTripEngineSeconds: number;
  };
}
