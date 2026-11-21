import { contextBridge, ipcRenderer } from "electron";

contextBridge.exposeInMainWorld("sob", {
  exportApp: (payload) => ipcRenderer.invoke("app:export", payload),
  exportSob: (payload) => ipcRenderer.invoke("sob:export", payload),
  // Loading is synchronous so the first screen renders with the saved budget.
  loadBudget: () => ipcRenderer.sendSync("budget:load"),
  budgetPath: () => ipcRenderer.sendSync("budget:path"),
  saveBudget: (json) => ipcRenderer.invoke("budget:save", json),
  appVersion: () => ipcRenderer.invoke("app:version"),
  checkForUpdates: () => ipcRenderer.invoke("updater:check"),
  downloadUpdate: () => ipcRenderer.invoke("updater:download"),
  installUpdate: () => ipcRenderer.invoke("updater:install"),
  // Push events from the main process; each returns a function that removes the listener.
  onUpdaterStatus: (callback) => { const listener = (_event, data) => callback(data); ipcRenderer.on("updater:status", listener); return () => ipcRenderer.removeListener("updater:status", listener); },
  onUpdaterProgress: (callback) => { const listener = (_event, data) => callback(data); ipcRenderer.on("updater:progress", listener); return () => ipcRenderer.removeListener("updater:progress", listener); },
  onMenuCheckForUpdates: (callback) => { const listener = () => callback(); ipcRenderer.on("menu:checkForUpdates", listener); return () => ipcRenderer.removeListener("menu:checkForUpdates", listener); },
});
