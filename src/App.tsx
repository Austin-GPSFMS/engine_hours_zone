/**
 * Engine Hours by Zone — main app.
 *
 *   Toolbar: Vehicle dropdown · Date range · Run / Export buttons
 *   Banners: standalone preview, errors
 *   Results: Summary KPIs · Zones table · Chronological timeline
 *
 * Same Zenith vocabulary as the Advanced Report Builder so the page reads
 * native inside MyGeotab.
 */

import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Banner,
  Button,
  Card,
  Content,
  DateRange,
  Dropdown,
  GET_LAST_SEVEN_DAYS_OPTION,
  GET_LAST_THIRTY_DAYS_OPTION,
  GET_LAST_MONTH_OPTION,
  GET_LAST_WEEK_OPTION,
  GET_THIS_MONTH_OPTION,
  GET_THIS_WEEK_OPTION,
  GET_TODAY_OPTION,
  GET_YESTERDAY_OPTION,
  type IDateRangeValue,
  type ISelectionItem,
} from "@geotab/zenith";
import type {
  Cluster,
  GeotabApi,
  GeotabDevice,
  GeotabPageState,
  ZoneReport,
} from "./types";
import {
  fetchAddresses,
  fetchDevices,
  fetchTripsAndEngineHours,
  friendlyError,
} from "./api/geotab";
import {
  buildStops,
  buildTimeline,
  clusterStops,
  ONE_MILE_METERS,
} from "./utils/cluster";
import { exportToXlsx } from "./utils/export";
import { Summary } from "./components/Summary";
import { ZoneTable } from "./components/ZoneTable";
import { Timeline } from "./components/Timeline";

interface AppProps {
  api: GeotabApi | null;
  pageState: GeotabPageState | null;
}

const dateRangeOptions = [
  GET_TODAY_OPTION(),
  GET_YESTERDAY_OPTION(),
  GET_THIS_WEEK_OPTION(),
  GET_LAST_WEEK_OPTION(),
  GET_THIS_MONTH_OPTION(),
  GET_LAST_MONTH_OPTION(),
  GET_LAST_SEVEN_DAYS_OPTION(),
  GET_LAST_THIRTY_DAYS_OPTION(),
];

const radiusItems: ISelectionItem[] = [
  { id: "1609.34", name: "Radius: 1 mile" },
  { id: "402.336", name: "Radius: 0.25 mile" },
  { id: "804.672", name: "Radius: 0.5 mile" },
  { id: "3218.69", name: "Radius: 2 miles" },
];

function defaultDateRange(): IDateRangeValue {
  const last7 = GET_LAST_SEVEN_DAYS_OPTION();
  const r = last7.getRange();
  return { from: r.from, to: r.to, label: last7.label };
}

export default function App({ api, pageState: _pageState }: AppProps) {
  const insideMyGeotab = api != null;

  // ---- State ----
  const [devices, setDevices] = useState<GeotabDevice[]>([]);
  const [devicesLoaded, setDevicesLoaded] = useState(false);
  const [devicesErr, setDevicesErr] = useState<string | null>(null);

  const [selectedDeviceId, setSelectedDeviceId] = useState<string | null>(null);
  const [dateRange, setDateRange] = useState<IDateRangeValue>(() =>
    defaultDateRange()
  );
  const [radiusMeters, setRadiusMeters] = useState<number>(ONE_MILE_METERS);

  const [isBuilding, setIsBuilding] = useState(false);
  const [buildErr, setBuildErr] = useState<string | null>(null);
  const [report, setReport] = useState<ZoneReport | null>(null);

  const [isExporting, setIsExporting] = useState(false);

  // ---- Device load ----
  useEffect(() => {
    if (!api) return;
    let cancelled = false;
    fetchDevices(api)
      .then((ds) => {
        if (cancelled) return;
        setDevices(ds);
        setDevicesLoaded(true);
      })
      .catch((e) => {
        if (cancelled) return;
        setDevicesErr(friendlyError(e));
        setDevicesLoaded(true);
      });
    return () => {
      cancelled = true;
    };
  }, [api]);

  const deviceItems = useMemo<ISelectionItem[]>(
    () =>
      devices.map((d) => ({
        id: d.id,
        name: d.name ?? d.id,
      })),
    [devices]
  );

  // ---- Handlers ----
  const onBuild = useCallback(async () => {
    if (!api) return;
    if (!selectedDeviceId) {
      setBuildErr("Pick a vehicle.");
      return;
    }
    if (!dateRange.from || !dateRange.to) {
      setBuildErr("Pick a date range.");
      return;
    }
    setBuildErr(null);
    setReport(null);
    setIsBuilding(true);
    try {
      const fromISO = new Date(dateRange.from).toISOString();
      const toISO = new Date(dateRange.to).toISOString();
      const device = devices.find((d) => d.id === selectedDeviceId);
      const deviceName = device?.name ?? selectedDeviceId;

      const { trips, engineHours } = await fetchTripsAndEngineHours(
        api,
        selectedDeviceId,
        fromISO,
        toISO
      );
      if (!trips || trips.length < 2) {
        throw new Error(
          `Need at least two trips in the date range to identify stops between them. Found ${trips?.length ?? 0}.`
        );
      }
      const stops = buildStops(trips, engineHours);
      if (stops.length === 0) {
        throw new Error("No usable stops found in this date range.");
      }
      const clusters: Cluster[] = clusterStops(stops, radiusMeters);

      // Reverse-geocode all cluster centers in one call.
      const addresses = await fetchAddresses(
        api,
        clusters.map((c) => ({ lat: c.centerLat, lng: c.centerLng }))
      );
      clusters.forEach((c, i) => {
        c.address =
          addresses[i] ?? `${c.centerLat.toFixed(5)}, ${c.centerLng.toFixed(5)}`;
      });

      const events = buildTimeline(trips, clusters);
      const totalZoneEngineSeconds = clusters.reduce(
        (acc, c) => acc + c.totalEngineSeconds,
        0
      );

      setReport({
        deviceName,
        fromDate: fromISO,
        toDate: toISO,
        clusters,
        events,
        totals: {
          totalVisits: stops.length,
          totalZoneEngineSeconds,
        },
      });
    } catch (e) {
      setBuildErr(friendlyError(e));
    } finally {
      setIsBuilding(false);
    }
  }, [api, selectedDeviceId, dateRange, devices, radiusMeters]);

  const onExport = useCallback(async () => {
    if (!report) return;
    setIsExporting(true);
    try {
      await exportToXlsx(report);
    } catch (e) {
      setBuildErr(`Export failed: ${friendlyError(e)}`);
    } finally {
      setIsExporting(false);
    }
  }, [report]);

  const onDeviceChange = (items: ISelectionItem[]) => {
    const id = items[0]?.id;
    setSelectedDeviceId(id != null ? String(id) : null);
  };

  const onRadiusChange = (items: ISelectionItem[]) => {
    const id = items[0]?.id;
    if (id != null) setRadiusMeters(parseFloat(String(id)));
  };

  // ---- Render ----
  return (
    <div className="ehz-page">
      <header className="ehz-page-header">
        <div>
          <h2>Engine Hours by Zone</h2>
          <p>
            Where did this vehicle accumulate engine hours? Stops within the
            chosen radius of each other are grouped into a single zone.
          </p>
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <Button
            type="secondary"
            onClick={onExport}
            disabled={!report || isExporting || isBuilding}
          >
            {isExporting ? "Exporting…" : "Export to Excel"}
          </Button>
          <Button
            type="primary"
            onClick={onBuild}
            disabled={!insideMyGeotab || isBuilding || !devicesLoaded}
          >
            {isBuilding ? "Building…" : "Run report"}
          </Button>
        </div>
      </header>

      <div className="ehz-toolbar">
        <Dropdown
          value={selectedDeviceId ? [selectedDeviceId] : []}
          dataItems={deviceItems}
          onChange={onDeviceChange}
          errorHandler={(e) => console.error("[EHZ] Vehicle:", e)}
          forceSelection={false}
          multiselect={false}
          showSelection
          showCounterPill={false}
          placeholder={devicesLoaded ? "Select vehicle" : "Loading vehicles…"}
        />
        <DateRange
          options={dateRangeOptions}
          value={dateRange}
          defaultValue={dateRange}
          onChange={(v: IDateRangeValue) => setDateRange(v)}
          withCalendar
        />
        <Dropdown
          value={[String(radiusMeters)]}
          dataItems={radiusItems}
          onChange={onRadiusChange}
          errorHandler={(e) => console.error("[EHZ] Radius:", e)}
          forceSelection
          multiselect={false}
          showSelection
          showCounterPill={false}
          placeholder="Cluster radius"
        />
      </div>

      <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
        {!insideMyGeotab && (
          <Banner type="info" header="Standalone preview">
            Open this page from inside MyGeotab to query live data.
          </Banner>
        )}
        {devicesErr && (
          <Banner type="error" header="Couldn't load vehicles">
            {devicesErr}
          </Banner>
        )}
        {buildErr && (
          <Banner type="error" header="Report failed">
            {buildErr}
          </Banner>
        )}
      </div>

      {report ? (
        <>
          <Summary report={report} />

          <Card title="Zones" fullWidth>
            <Content>
              <ZoneTable clusters={report.clusters} />
            </Content>
          </Card>

          <Card title="Timeline" fullWidth>
            <Content>
              <Timeline events={report.events} />
            </Content>
          </Card>
        </>
      ) : (
        !isBuilding &&
        insideMyGeotab && (
          <div className="ehz-empty">
            Pick a vehicle and date range, then click <strong>Run report</strong>.
          </div>
        )
      )}

      <footer className="ehz-footer">
        <small>
          Engine hours sourced from{" "}
          <code>DiagnosticEngineHoursAdjustmentId</code>. Values interpolated
          between bracketing StatusData samples.
        </small>
      </footer>
    </div>
  );
}
