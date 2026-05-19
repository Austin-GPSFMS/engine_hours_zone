import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

/**
 * Vite config for the GPSFMS Engine Hours by Zone MyGeotab add-in.
 *
 * Hosted at:
 *   https://austin-gpsfms.github.io/engine_hours_zone/dist/index.html
 *
 * `base: ""` produces CLEAN relative asset paths in the built index.html
 * (`<script src="assets/index-XXX.js">` — no leading `/`, no `./`). This
 * is the safest choice for MyGeotab add-ins because:
 *   - Absolute paths starting with `/` get duplicated by MyGeotab's iframe
 *     loader, producing `/engine_hours_zone/dist//engine_hours_zone/dist/...`
 *   - `./` paths can survive into the request URL as literal `/./` segments
 *     that some servers (including GitHub Pages) treat as separate path
 *     segments and 404 on.
 * Bare relative paths resolve naturally against the document URL with no
 * artifacts in the request URL.
 */
export default defineConfig({
  base: "",
  plugins: [react()],
  build: {
    outDir: "dist",
    assetsDir: "assets",
    sourcemap: true,
    // Inline assets up to ~20KB so we ship fewer files.
    assetsInlineLimit: 20480,
    rollupOptions: {
      output: {
        // Pin filenames so MyGeotab's iframe cache doesn't go stale silently.
        entryFileNames: "assets/[name]-[hash].js",
        chunkFileNames: "assets/[name]-[hash].js",
        assetFileNames: "assets/[name]-[hash].[ext]",
      },
    },
  },
  server: {
    port: 5173,
    strictPort: false,
  },
});
