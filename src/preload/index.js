import { contextBridge, ipcRenderer } from "electron";

contextBridge.exposeInMainWorld("sob", {
  saveAllocation: (payload) => ipcRenderer.invoke("allocation:save", payload),
  exportApp: (payload) => ipcRenderer.invoke("app:export", payload),
  // Loading is synchronous so the first screen renders with the saved budget.
  loadBudget: () => ipcRenderer.sendSync("budget:load"),
  budgetPath: () => ipcRenderer.sendSync("budget:path"),
  saveBudget: (json) => ipcRenderer.invoke("budget:save", json),
});
