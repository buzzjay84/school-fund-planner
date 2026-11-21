import { app, BrowserWindow, dialog, ipcMain, Menu, shell } from "electron";
import updater from "electron-updater"; // CommonJS package: take autoUpdater off the default export
import { appWorkbook, sobWorkbook } from "./reports.js";
import path from "node:path";
import { DatabaseSync } from "node:sqlite";
import appIcon from "../image/app_logo.png?asset";

// Keep the data folder the app has always used: with a productName set, Electron would otherwise
// switch to "%APPDATA%\School Fund Planner" and the saved budget would look lost.
app.setPath("userData", path.join(app.getPath("appData"), "school-fund-planner"));

// The budget is stored in a SQLite database in the app's data folder
// (e.g. %APPDATA%\school-fund-planner\fund-planner.db), one row holding the whole budget.
const databaseFile = () => path.join(app.getPath("userData"), "fund-planner.db");
let db = null;
function database() {
  if (!db) {
    db = new DatabaseSync(databaseFile());
    db.exec(`PRAGMA journal_mode = WAL;
      CREATE TABLE IF NOT EXISTS budget (id INTEGER PRIMARY KEY CHECK (id = 1), data TEXT NOT NULL, updated_at TEXT NOT NULL);`);
  }
  return db;
}

function loadBudgetRow() {
  return database().prepare("SELECT data FROM budget WHERE id = 1").get()?.data ?? null;
}

// Each save is one SQLite statement, so it is written completely or not at all.
function saveBudgetRow(json) {
  JSON.parse(json); // refuse anything that is not valid JSON
  database().prepare(`INSERT INTO budget (id, data, updated_at) VALUES (1, ?, ?)
    ON CONFLICT (id) DO UPDATE SET data = excluded.data, updated_at = excluded.updated_at`).run(json, new Date().toISOString());
  return databaseFile();
}

async function saveWorkbook(workbook, title, name, window) {
  const { canceled, filePath } = await dialog.showSaveDialog(window, {
    title,
    defaultPath: path.join(app.getPath("documents"), name),
    filters: [{ name: "Excel workbook", extensions: ["xlsx"] }],
  });
  // Electron on Windows can leave the page without keyboard focus after a native dialog: hand it back.
  window?.focus(); window?.webContents.focus();
  if (canceled || !filePath) return null;
  try {
    await workbook.xlsx.writeFile(filePath);
  } catch (error) {
    // Excel locks a workbook while it is open (EBUSY); say so instead of the raw error.
    if (error.code === "EBUSY") throw new Error(`FILE_OPEN:${path.basename(filePath)}`);
    throw error;
  }
  return filePath;
}

const fileName = (school, label) => `${String(school || "School").replace(/[\\/:*?"<>|]/g, "-")} - ${label}.xlsx`;

// ── Updates from GitHub Releases (as in ieces-portal): the renderer's update dialog drives
// check → download → restart & install; status and progress are pushed to it.
const { autoUpdater } = updater;
const RELEASES = "https://github.com/buzzjay84/school-fund-planner/releases/latest";
autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = true;
const sendAll = (channel, data) => BrowserWindow.getAllWindows().forEach((window) => window.webContents.send(channel, data));
autoUpdater.on("checking-for-update", () => sendAll("updater:status", { status: "checking" }));
autoUpdater.on("update-available", (info) => sendAll("updater:status", { status: "available", version: info.version, releaseDate: info.releaseDate }));
autoUpdater.on("update-not-available", () => sendAll("updater:status", { status: "up-to-date" }));
autoUpdater.on("download-progress", (progress) => sendAll("updater:progress", { percent: Math.round(progress.percent) }));
autoUpdater.on("update-downloaded", (info) => sendAll("updater:status", { status: "downloaded", version: info.version }));
autoUpdater.on("error", (error) => sendAll("updater:status", { status: "error", message: error.message }));

function buildMenu() {
  const isMac = process.platform === "darwin";
  Menu.setApplicationMenu(Menu.buildFromTemplate([
    ...(isMac ? [{ role: "appMenu" }] : []),
    { role: "fileMenu" },
    { role: "editMenu" },
    { role: "viewMenu" },
    { role: "windowMenu" },
    {
      role: "help",
      submenu: [
        { label: "Check for Updates…", click: () => sendAll("menu:checkForUpdates") },
        { label: "Release Notes", click: () => shell.openExternal(RELEASES) },
        { type: "separator" },
        { label: `Version ${app.getVersion()}`, enabled: false },
      ],
    },
  ]));
}

function createWindow() {
  const window = new BrowserWindow({
    width: 1440,
    height: 940,
    minWidth: 1120,
    minHeight: 720,
    backgroundColor: "#f5f6f3",
    title: "School Fund Planner",
    icon: appIcon,
    webPreferences: {
      preload: path.join(import.meta.dirname, "../preload/index.mjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });
  if (process.env.ELECTRON_RENDERER_URL) window.loadURL(process.env.ELECTRON_RENDERER_URL);
  else window.loadFile(path.join(import.meta.dirname, "../renderer/index.html"));
}

app.whenReady().then(() => {
  ipcMain.handle("sob:export", (event, data) => saveWorkbook(sobWorkbook(data), "Save School Operating Budget", fileName(data.school, `SOB ${data.fiscalYear}`), BrowserWindow.fromWebContents(event.sender)));
  ipcMain.handle("app:export", (event, data) => saveWorkbook(appWorkbook(data), "Save Annual Procurement Plan", fileName(data.school, `APP ${data.fiscalYear}`), BrowserWindow.fromWebContents(event.sender)));
  ipcMain.on("budget:load", (event) => {
    try { event.returnValue = loadBudgetRow(); } catch (error) { console.error(error); event.returnValue = null; }
  });
  ipcMain.on("budget:path", (event) => { event.returnValue = databaseFile(); });
  ipcMain.handle("budget:save", (_event, json) => saveBudgetRow(json));
  ipcMain.handle("app:version", () => app.getVersion());
  ipcMain.handle("updater:check", async () => {
    // electron-updater skips the check silently in an unpackaged (npm run dev) app, so say so.
    if (!app.isPackaged) return { ok: false, message: "Updates work only in the installed app, not in development mode." };
    try { await autoUpdater.checkForUpdates(); return { ok: true }; } catch (error) { return { ok: false, message: error.message }; }
  });
  ipcMain.handle("updater:download", async () => {
    // macOS only installs updates of a code-signed app; these builds are unsigned, so download the new version by hand.
    if (process.platform === "darwin") { await shell.openExternal(RELEASES); return { ok: true, external: true }; }
    try { await autoUpdater.downloadUpdate(); return { ok: true }; } catch (error) { return { ok: false, message: error.message }; }
  });
  ipcMain.handle("updater:install", () => autoUpdater.quitAndInstall());
  buildMenu();
  createWindow();
  // Quiet check a few seconds after launch; the renderer only speaks up if a new version is found.
  if (app.isPackaged) setTimeout(() => autoUpdater.checkForUpdates().catch(() => {}), 3000);
  app.on("activate", () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});

app.on("window-all-closed", () => { if (process.platform !== "darwin") app.quit(); });
