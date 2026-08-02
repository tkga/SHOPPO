import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  base: "./", // relative paths so it works at username.github.io/<repo>/ or any subpath
  plugins: [react()],
  build: {
    outDir: "dist",
  },
});
