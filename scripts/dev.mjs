import { spawn } from "node:child_process";
import { createRequire } from "node:module";
import path from "node:path";

const require = createRequire(import.meta.url);
const cli = path.join(path.dirname(require.resolve("electron-vite/package.json")), "bin/electron-vite.js");
const env = { ...process.env };
delete env.ELECTRON_RUN_AS_NODE;

// --watch: restart Electron when main/preload code changes, not only reload the page.
const child = spawn(process.execPath, [cli, "dev", "--watch"], {
  cwd: process.cwd(),
  env,
  stdio: "inherit",
});

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => child.kill(signal));
}
child.on("exit", (code) => process.exit(code ?? 1));
