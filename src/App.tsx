/**
 * Engine Hours by Zone — main app.
 *
 *   Toolbar: Group filter · Vehicle multi-select · Date range · Radius
 *            Buttons: Run · Export
 *   Banners: standalone, errors, progress
 *   Results: Summary KPIs · One VehicleReport per selected vehicle
 *
 * Same Zenith vocabulary as the Advanced Report Builder so the page reads
 * native inside MyGeotab.
 */

import { useCallback, useEffect, useMemo, useState } from "react";
import {
  Banner,
  Button,
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
  GeotabApi,
  GeotabDevice,
  GeotabGroup,
  GeotabPageState,
  MultiVehicleReport,
} from "./types";
import {
  fetchDevices,
  fetchGroups,
  friendlyError,
} from "./api/geotab";
import { ONE_MILE_METERS } from "./utils/cluster";
import { buildMultiVehicleReport } from "./utils/multiVehicle";
import { exportToXlsx } from "./utils/export";
import { Summary } from "./components/Summary";
import { VehicleReport } from "./components/VehicleReport";
import { GroupFilterPicker } from "./components/GroupFilterPicker";

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

const ALL_VEHICLES_ID = "__ALL__";

function defaultDateRange(): IDateRangeValue {
  const last7 = GET_LAST_SEVEN_DAYS_OPTION();
  const r = last7.getRange();
  return { from: r.from, to: r.to, label: last7.label };
}

export default function App({ api, pageState: _pageState }: AppProps) {
  const insideMyGeotab = api != null;

  // ---- Groups & devices ----
  const [groupsById, setGroupsById] = useState<Map<string, GeotabGroup>>(
    () => new Map()
  );
  const [groupsLoaded, setGroupsLoaded] = useState(false);
  const [groupsErr, setGroupsErr] = useState<string | null>(null);

  const [selectedGroupIds, setSelectedGroupIds] = useState<string[]>([
    "GroupCompanyId",
  ]);

  const [devices, setDevices] = useState<GeotabDevice[]>([]);
  const [devicesLoading, setDevicesLoading] = useState(false);
  const [devicesErr, setDevicesErr] = useState<string | null>(null);

  const [selectedDeviceIds, setSelectedDeviceIds] = useState<string[]>([]);

  // ---- Form ----
  const [dateRange, setDateRange] = useState<IDateRangeValue>(() =>
    defaultDateRange()
  );
  const [radiusMeters, setRadiusMeters] = useState<number>(ONE_MILE_METERS);

  // ---- Build state ----
  const [isBuilding, setIsBuilding] = useState(false);
  const [buildErr, setBuildErr] = useState<string | null>(null);
  const [report, setReport] = useState<MultiVehicleReport | null>(null);
  const [progress, setProgress] = useState<{
    done: number;
    total: number;
    currentName: string;
  } | null>(null);
  const [isExporting, setIsExporting] = useState(false);

  // ---- Initial group load ----
  useEffect(() => {
    if (!api) return;
    let cancelled = false;
    fetchGroups(api)
      .then((m) => {
        if (cancelled) return;
        setGroupsById(m);
        setGroupsLoaded(true);
      })
      .catch((e) => {
        if (cancelled) return;
        setGroupsErr(friendlyError(e));
        setGroupsLoaded(true);
      });
    return () => {
      cancelled = true;
    };
  }, [api]);

  // ---- Re-fetch devices when group filter changes ----
  useEffect(() => {
    if (!api) return;
    let cancelled = false;
    setDevicesLoading(true);
    setDevicesErr(null);
    fetchDevices(api, selectedGroupIds, false)
      .then((ds) => {
        if (cancelled) return;
        setDevices(ds);
        // Drop any selected device IDs that fell out of the new group scope.
        setSelectedDeviceIds((prev) => {
          const inScope = new Set(ds.map((d) => d.id));
          return prev.filter((id) => id === ALL_VEHICLES_ID || inScope.has(id));
        });
      })
      .catch((e) => {
        if (cancelled) return;
        setDevicesErr(friendlyError(e));
      })
      .finally(() => {
        if (!cancelled) setDevicesLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [api, selectedGroupIds]);

  // ---- Build vehicle picker items ----
  const deviceItems = useMemo<ISelectionItem[]>(() => {
    const all: ISelectionItem = {
      id: ALL_VEHICLES_ID,
      name: `All vehicles in group (${devices.length})`,
    };
    return [
      all,
      ...devices.map((d) => ({ id: d.id, name: d.name ?? d.id })),
    ];
  }, [devices]);

  const effectiveDeviceIds = useMemo<string[]>(() => {
    if (selectedDeviceIds.includes(ALL_VEHICLES_ID)) {
      return devices.map((d) => d.id);
    }
    return selectedDeviceIds;
  }, [selectedDeviceIds, devices]);

  const devicesById = useMemo(
    () => new Map(devices.map((d) => [d.id, d])),
    [devices]
  );

  // ---- Handlers ----
  const onRun = useCallback(async () => {
    if (!api) return;
    if (effectiveDeviceIds.length === 0) {
      setBuildErr("Pick at least one vehicle.");
      return;
    }
    if (!dateRange.from || !dateRange.to) {
      setBuildErr("Pick a date range.");
      return;
    }
    setBuildErr(null);
    setReport(null);
    setIsBuilding(true);
    setProgress({ done: 0, total: effectiveDeviceIds.length, currentName: "" });
    try {
      const fromISO = new Date(dateRange.from).toISOString();
      const toISO = new Date(dateRange.to).toISOString();
      const r = await buildMultiVehicleReport({
        api,
        deviceIds: effectiveDeviceIds,
        devicesById,
        fromDate: fromISO,
        toDate: toISO,
        radiusMeters,
        onProgress: (done, total, currentName) =>
          setProgress({ done, total, currentName }),
      });
      setReport(r);
    } catch (e) {
      setBuildErr(friendlyError(e));
    } finally {
      setIsBuilding(false);
      setProgress(null);
    }
  }, [api, effectiveDeviceIds, dateRange, devicesById, radiusMeters]);

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

  const onDevicesChange = (items: ISelectionItem[]) => {
    // If "All" got picked, collapse the selection to just All so the chip
    // count stays small.
    const ids = items.map((i) => String(i.id));
    if (ids.includes(ALL_VEHICLES_ID)) {
      setSelectedDeviceIds([ALL_VEHICLES_ID]);
    } else {
      setSelectedDeviceIds(ids);
    }
  };

  const onRadiusChange = (items: ISelectionItem[]) => {
    const id = items[0]?.id;
    if (id != null) setRadiusMeters(parseFloat(String(id)));
  };

  const onGroupsChange = useCallback((ids: string[]) => {
    setSelectedGroupIds(ids.length > 0 ? ids : ["GroupCompanyId"]);
  }, []);

  // ---- Render ----
  return (
    <div className="ehz-page">
      <header className="ehz-page-header">
        <div>
          <h2>Engine Hours by Zone</h2>
          <p>
            Multi-vehicle, trip-level breakdown of engine hours by location.
            Stops within the chosen radius are clustered into a zone.
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
            onClick={onRun}
            disabled={
              !insideMyGeotab ||
              isBuilding ||
              devicesLoading ||
              effectiveDeviceIds.length === 0
            }
          >
            {isBuilding ? "Running…" : "Run report"}
          </Button>
        </div>
      </header>

      <div className="ehz-toolbar">
        {groupsLoaded ? (
          <GroupFilterPicker
            groupsById={groupsById}
            initialGroupIds={selectedGroupIds}
            onChange={onGroupsChange}
            onError={(e) => setGroupsErr(e.message)}
          />
        ) : (
          <div style={{ color: "#6b7280", fontSize: 13, fontStyle: "italic" }}>
            Loading groups…
          </div>
        )}
        <Dropdown
          value={selectedDeviceIds}
          dataItems={deviceItems}
          onChange={onDevicesChange}
          errorHandler={(e) => console.error("[EHZ] Vehicles:", e)}
          forceSelection={false}
          multiselect
          showSelection
          showCounterPill
          placeholder={
            devicesLoading ? "Loading vehicles…" : "Select vehicles"
          }
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
        {groupsErr && (
          <Banner type="error" header="Couldn't load groups">
            {groupsErr}
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
        {progress && (
          <div className="ehz-progress">
            <span>
              Processing vehicle {progress.done} of {progress.total}
              {progress.currentName ? ` — ${progress.currentName}` : ""}
            </span>
            <div className="ehz-progress-bar">
              <div
                style={{
                  width: `${(progress.done / Math.max(1, progress.total)) * 100}%`,
                }}
              />
            </div>
          </div>
        )}
      </div>

      {report ? (
        <>
          <Summary report={report} />
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              gap: 12,
            }}
          >
            {report.vehicles.map((v) => (
              <VehicleReport key={v.deviceId} vehicle={v} />
            ))}
          </div>
        </>
      ) : (
        !isBuilding &&
        insideMyGeotab && (
          <div className="ehz-empty">
            Pick a group + vehicle(s) + date range, then click{" "}
            <strong>Run report</strong>.
          </div>
        )
      )}

      <footer className="ehz-footer">
        <small>
          Engine hours sourced from{" "}
          <code>DiagnosticEngineHoursAdjustmentId</code>. Values interpolated
          between bracketing StatusData samples. Trip and stop boundaries from
          the Trip object's <code>start</code> / <code>stop</code> /{" "}
          <code>nextTripStart</code>.
        </small>
      </footer>
    </div>
  );
}
