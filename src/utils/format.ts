/**
 * Display helpers — time/duration/HTML escape.
 */

/** "1h 23m" or "—" when null/NaN. */
export function formatDuration(ms: number | null | undefined): string {
  if (ms == null || isNaN(ms)) return "—";
  const totalMinutes = Math.max(0, Math.round(ms / 60000));
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${h}h ${m < 10 ? "0" : ""}${m}m`;
}

/** "08:23 AM" — locale-aware. */
export function formatTime(d: string | Date): string {
  return new Date(d).toLocaleTimeString([], {
    hour: "2-digit",
    minute: "2-digit",
  });
}

/** Full date+time, locale-aware. */
export function formatDateTime(d: string | Date): string {
  return new Date(d).toLocaleString();
}

/** Just the weekday + date — e.g. "Mon May 19 2026". */
export function formatDay(d: string | Date): string {
  return new Date(d).toDateString();
}

/** Cumulative engine seconds → fractional hours, 2 decimals. */
export function secondsToHours(seconds: number | null | undefined): string {
  if (seconds == null) return "—";
  return (seconds / 3600).toFixed(2);
}

/** Kilometers → miles, 1 decimal. */
export const KM_PER_MILE = 1.609344;
export function kmToMiles(km: number | null | undefined): string {
  if (km == null) return "—";
  return (km / KM_PER_MILE).toFixed(1);
}
