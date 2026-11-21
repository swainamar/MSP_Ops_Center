import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// The build lands in ../static/app and is served by msp_runbook.py at /app/ (no Node.js on the server).
// For development, `npm run dev` proxies the API to a running MSP Ops Center on port 8080.
export default defineConfig({
  plugins: [react()],
  base: "/app/",
  build: { outDir: "../static/app", emptyOutDir: true, sourcemap: false, assetsDir: "assets", chunkSizeWarningLimit: 600 },
  server: { port: 5173, proxy: { "/api": "http://127.0.0.1:8080", "/static": "http://127.0.0.1:8080" } },
});
