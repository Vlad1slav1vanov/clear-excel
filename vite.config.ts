import { resolve } from "node:path";

import { defineConfig } from "vite";

export default defineConfig({
  build: {
    emptyOutDir: true,
    lib: {
      entry: resolve(import.meta.dirname, "src/index.ts"),
      fileName: () => "index.js",
      formats: ["es"],
    },
    minify: false,
    rolldownOptions: {
      external: [/^node:/, "jszip"],
    },
    sourcemap: true,
    target: "es2022",
  },
});
