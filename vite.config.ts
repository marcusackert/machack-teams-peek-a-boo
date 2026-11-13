import { defineConfig } from "vite";
import basicSsl from "@vitejs/plugin-basic-ssl";
import { resolve } from "node:path";

const page = (name: string) => resolve(import.meta.dirname, name);

export default defineConfig({
  // Relative base so the build works at a domain root or a sub-path (e.g. GitHub Pages).
  base: "./",
  plugins: [basicSsl()],
  server: { port: 5173, strictPort: true },
  build: {
    rollupOptions: {
      input: {
        index: page("index.html"),
        config: page("config.html"),
        authStart: page("auth-start.html"),
        authEnd: page("auth-end.html"),
      },
    },
  },
});
