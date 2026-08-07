import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// This must match your GitHub repo name so GitHub Pages loads assets correctly.
export default defineConfig({
  plugins: [react()],
  base: "/LL-APP/",
});
