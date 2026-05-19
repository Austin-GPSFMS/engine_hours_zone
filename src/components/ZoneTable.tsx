import type { Cluster } from "../types";
import { formatDuration, secondsToHours } from "../utils/format";

/**
 * Zones table — sorted by engine hours accumulated (descending).
 * Each row links to a Google Maps view of the zone center.
 */
export function ZoneTable({ clusters }: { clusters: Cluster[] }) {
  const sorted = clusters
    .slice()
    .sort((a, b) => b.totalEngineSeconds - a.totalEngineSeconds);

  if (sorted.length === 0) {
    return <div className="ehz-empty">No zones to display.</div>;
  }

  return (
    <div className="ehz-table-wrap">
      <table className="ehz-table">
        <thead>
          <tr>
            <th>Zone</th>
            <th>Location</th>
            <th>Visits</th>
            <th>Total stopped</th>
            <th>Engine hrs accumulated</th>
          </tr>
        </thead>
        <tbody>
          {sorted.map((c) => (
            <tr key={c.id}>
              <td>
                <strong>{c.id}</strong>
              </td>
              <td>
                {c.address ?? "—"}{" "}
                <a
                  className="ehz-maplink"
                  href={`https://www.google.com/maps?q=${c.centerLat},${c.centerLng}`}
                  target="_blank"
                  rel="noopener"
                >
                  map
                </a>
              </td>
              <td>{c.visits}</td>
              <td>{formatDuration(c.totalStoppedMs)}</td>
              <td>
                <strong>{secondsToHours(c.totalEngineSeconds)}</strong>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
