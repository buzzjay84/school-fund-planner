import { resolve } from "node:path";
import { defineConfig, externalizeDepsPlugin } from "electron-vite";

export default defineConfig({
  main: { plugins: [externalizeDepsPlugin()] },
  preload: { plugins: [externalizeDepsPlugin()] },
  // src/image is served/copied as-is, so logos load from ./school_logo/<file>.
  renderer: { publicDir: resolve("src/image") },
});
