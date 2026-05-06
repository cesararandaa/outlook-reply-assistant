import { defineConfig } from "vite";
import mkcert from "vite-plugin-mkcert";
import { resolve } from "path";

export default defineConfig({
  plugins: [mkcert()],
  server: {
    port: 3000,
    https: true,
    proxy: {
      "/api": {
        target: "http://localhost:3001",
        changeOrigin: true,
      },
    },
  },
  build: {
    outDir: "dist",
    rollupOptions: {
      input: {
        taskpane: resolve(__dirname, "public/taskpane.html"),
      },
    },
  },
  publicDir: "public",
});
