import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  server: {
    host: true,
    allowedHosts: true,
    proxy: {
      "/ocr": "http://127.0.0.1:5000",
      "/preview": "http://127.0.0.1:5000",
      "/download": "http://127.0.0.1:5000",
      "/upload": "http://127.0.0.1:5000",
      "/api": "http://127.0.0.1:5000",
    },
  },
});