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

/** A single stop period (gap between two consecutive trips) on a device. */
export interface Stop {
  lat: number;
  lng: number;
  arrive: string;
  depart: string;
  durationMs: number;
  /** Cumulative engine seconds delta across this stop, or null if data missing. */
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

/** A chronological event in the per-day timeline. */
export type TimelineEvent =
  | {
      type: "zone";
      cluster: Cluster;
      stop: Stop;
      start: string;
      end: string;
      durationMs: number;
      engineSeconds: number | null;
    }
  | {
      type: "transit";
      start: string;
      end: string;
      distanceKm: number;
      durationMs: number;
    };

/** Output of buildReport — what the UI renders. */
export interface ZoneReport {
  deviceName: string;
  fromDate: string;
  toDate: string;
  clusters: Cluster[];
  events: TimelineEvent[];
  totals: {
    totalVisits: number;
    totalZoneEngineSeconds: number;
  };
}
