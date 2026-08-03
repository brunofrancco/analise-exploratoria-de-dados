import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// base: "./" makes the build use relative asset paths, so it works
// both on GitHub Pages project sites (username.github.io/repo-name/)
// and if served from any other subpath.
export default defineConfig({
  plugins: [react()],
  base: "./",
});
