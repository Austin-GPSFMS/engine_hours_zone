/**
 * Entry point.
 *
 * The MyGeotab add-in lifecycle calls a function registered at
 *   window.geotab.addin.engineHoursByZone
 * which returns { initialize, focus, blur }.
 *
 *   initialize(api, state, callback) — MyGeotab calls this when the page mounts.
 *                                      We create the React root here and pass
 *                                      api+state into App.
 *   focus(api, state)                — Called when the page gains focus.
 *                                      We re-render App with the (possibly
 *                                      refreshed) api/state.
 *   blur()                           — Called when the page loses focus. No-op.
 *
 * Mirrors advanced_report_builder/src/main.tsx exactly — including the
 * standalone-preview fallback so that opening dist/index.html directly in
 * a browser (outside MyGeotab) still paints the page.
 */

import { StrictMode } from "react";
import { createRoot, type Root } from "react-dom/client";
import App from "./App";
import type { GeotabApi, GeotabPageState } from "./types";
import "./styles.css";

declare global {
  interface Window {
    geotab?: {
      addin?: Record<
        string,
        () => {
          initialize: (
            api: GeotabApi,
            state: GeotabPageState,
            callback: () => void
          ) => void;
          focus: (api: GeotabApi, state: GeotabPageState) => void;
          blur: () => void;
        }
      >;
    };
  }
}

let root: Root | null = null;
let currentApi: GeotabApi | null = null;
let currentState: GeotabPageState | null = null;

function mount() {
  const container = document.getElementById("root");
  if (!container) {
    console.error("[EHZ] #root element not found");
    return;
  }
  if (!root) {
    root = createRoot(container);
  }
  root.render(
    <StrictMode>
      <App api={currentApi} pageState={currentState} />
    </StrictMode>
  );
}

window.geotab = window.geotab || {};
window.geotab.addin = window.geotab.addin || {};
window.geotab.addin.engineHoursByZone = function () {
  return {
    initialize(api, state, callback) {
      currentApi = api;
      currentState = state;
      try {
        mount();
      } catch (e) {
        console.error("[EHZ] initialize failed:", e);
      }
      callback();
    },
    focus(api, state) {
      currentApi = api;
      currentState = state;
      mount();
    },
    blur() {
      // No-op. Hook here later if we need to cancel inflight requests.
    },
  };
};

// Standalone load: paint the App with no api if MyGeotab never calls
// initialize within 500ms (i.e. someone opened dist/index.html directly).
if (typeof window !== "undefined") {
  const standaloneTimer = window.setTimeout(() => {
    if (!root) {
      console.warn(
        "[EHZ] No MyGeotab initialize() detected — rendering standalone preview."
      );
      mount();
    }
  }, 500);

  const origInit = window.geotab!.addin!.engineHoursByZone;
  window.geotab!.addin!.engineHoursByZone = function () {
    window.clearTimeout(standaloneTimer);
    return origInit();
  };
}
