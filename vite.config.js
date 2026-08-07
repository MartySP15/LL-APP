import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// IMPORTANT: change "base" to match your GitHub repo name, e.g. "/orders-invoicing-app/"
// This is required for GitHub Pages to load assets from the right path.
// If you're deploying to a custom domain, set base to "/".
export default defineConfig({
  plugins: [react()],
  base: "/orders-invoicing-app/",
});
