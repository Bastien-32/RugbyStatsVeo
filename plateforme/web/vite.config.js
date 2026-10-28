import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// En developpement, Vite fait lui-meme la passerelle vers
// l'API : le front appelle /api comme en production, et
// aucune origine ne change entre les deux.
export default defineConfig({
  plugins: [react()],
  server: {
    host: true,
    port: 5173,
    proxy: {
      "/api": {
        target: "http://localhost:8080",
        changeOrigin: true,
        rewrite: (chemin) => chemin.replace(/^\/api/, ""),
      },
    },
  },
});
