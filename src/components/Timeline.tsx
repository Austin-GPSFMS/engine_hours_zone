import type { TimelineEvent } from "../types";
import {
  formatDay,
  formatDuration,
  formatTime,
  kmToMiles,
  secondsToHours,
} from "../utils/format";

/**
 * Chronological timeline view — for each day, walks the alternating
 * Zone / Transit / Zone / Transit sequence.
 */
export function Timeline({ events }: { events: TimelineEvent[] }) {
  if (events.length === 0) {
    return <div className="ehz-empty">No timeline events.</div>;
  }

  // Group events into day buckets so we render a day header before each block.
  const days: Array<{ label: string; events: TimelineEvent[] }> = [];
  for (const ev of events) {
    const label = formatDay(ev.start);
    const last = days[days.length - 1];
    if (last && last.label === label) {
      last.events.push(ev);
    } else {
      days.push({ label, events: [ev] });
    }
  }

  return (
    <div className="ehz-timeline">
      {days.map((day, di) => (
        <div key={di}>
          <div className="ehz-timeline-day">{day.label}</div>
          {day.events.map((ev, ei) => {
            if (ev.type === "zone") {
              return (
                <div key={ei} className="ehz-tl-row ehz-tl-zone">
                  <div className="ehz-tl-time">
                    {formatTime(ev.start)} – {formatTime(ev.end)}
                  </div>
                  <div className="ehz-tl-body">
                    <strong>Zone {ev.cluster.id}</strong> · {ev.cluster.address ?? "—"}
                    <div className="ehz-tl-stats">
                      Stopped {formatDuration(ev.durationMs)} ·{" "}
                      {ev.engineSeconds != null ? (
                        <strong>
                          {secondsToHours(ev.engineSeconds)} engine hrs accumulated
                        </strong>
                      ) : (
                        <em>engine hrs unavailable</em>
                      )}
                    </div>
                  </div>
                </div>
              );
            }
            return (
              <div key={ei} className="ehz-tl-row ehz-tl-transit">
                <div className="ehz-tl-time">
                  {formatTime(ev.start)} – {formatTime(ev.end)}
                </div>
                <div className="ehz-tl-body">
                  → Transit {kmToMiles(ev.distanceKm)} mi · {formatDuration(ev.durationMs)}
                </div>
              </div>
            );
          })}
        </div>
      ))}
    </div>
  );
}
