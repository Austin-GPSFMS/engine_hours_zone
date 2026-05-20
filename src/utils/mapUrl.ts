/**
 * Map URL builders.
 *
 * Default behavior: open the MyGeotab native map zoomed to a small bounding
 * box around the requested point. URL shape (matches what MyGeotab generates
 * when you click a location pin):
 *
 *   https://{server}/{database}/#map,drivers:all,liveVehicleIds:all,mapBounds:!(north,east,south,west)
 *
 * The `mapBounds:!(...)` segment is a YAML-flavored flow-style array used by
 * MyGeotab's URL hash router. The four numbers are the NE-then-SW corners of
 * the visible rectangle.
 *
 * When we don't know the database/server (e.g. someone opens the Add-In's
 * built HTML directly outside MyGeotab), we fall back to Google Maps.
 */

import type { GeotabSessionInfo } from "../types";

/** Half-width of the bounding box in degrees. ~0.003 deg ≈ 333 m. */
const BOUNDS_OFFSET_DEG = 0.003;

/** Round to 5 decimal places (~1 m precision) to keep URLs short. */
function fmt(n: number): string {
  return n.toFixed(5);
}

export function geotabMapUrl(
  session: GeotabSessionInfo | null,
  lat: number | null | undefined,
  lng: number | null | undefined
): string | null {
  if (lat == null || lng == null) return null;
  if (!session || !session.database || !session.server) return null;
  const north = fmt(lat + BOUNDS_OFFSET_DEG);
  const east = fmt(lng + BOUNDS_OFFSET_DEG);
  const south = fmt(lat - BOUNDS_OFFSET_DEG);
  const west = fmt(lng - BOUNDS_OFFSET_DEG);
  return `https://${session.server}/${session.database}/#map,drivers:all,liveVehicleIds:all,mapBounds:!(${north},${east},${south},${west})`;
}

export function googleMapsUrl(
  lat: number | null | undefined,
  lng: number | null | undefined,
  address?: string | null
): string | null {
  if (lat != null && lng != null) {
    return `https://www.google.com/maps?q=${lat},${lng}`;
  }
  if (address) {
    return `https://www.google.com/maps?q=${encodeURIComponent(address)}`;
  }
  return null;
}

/**
 * Top-level helper: prefer the MyGeotab map, fall back to Google Maps when
 * no session is available. Returns null only when we have nothing to point at.
 */
export function mapUrlForPoint(
  session: GeotabSessionInfo | null,
  lat: number | null | undefined,
  lng: number | null | undefined,
  address?: string | null
): string | null {
  return (
    geotabMapUrl(session, lat, lng) ?? googleMapsUrl(lat, lng, address)
  );
}
