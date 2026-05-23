import { defineConfig } from "vite"

export default defineConfig({
  server: {
    port: 5173,
    host: true, // bind to all interfaces so the dev server is LAN-reachable
  },
  build: {
    target: "es2022",
    outDir: "dist",
  },
})
