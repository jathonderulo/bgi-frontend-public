import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import {bgiurl} from "./src/url.ts";

export default defineConfig({
  plugins: [react()],
  server: {
    proxy: {
      "/api": `${bgiurl}`,
    },
  },
});
