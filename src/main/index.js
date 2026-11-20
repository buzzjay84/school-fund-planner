import { app, BrowserWindow, dialog, ipcMain, shell } from "electron";
import ExcelJS from "exceljs";
import path from "node:path";
import { DatabaseSync } from "node:sqlite";
import appIcon from "../image/app_logo.png?asset";

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

async function saveWorkbook(workbook, title, name) {
  const { canceled, filePath } = await dialog.showSaveDialog({
    title,
    defaultPath: path.join(app.getPath("documents"), name),
    filters: [{ name: "Excel workbook", extensions: ["xlsx"] }],
  });
  if (canceled || !filePath) return null;
  await workbook.xlsx.writeFile(filePath);
  return filePath;
}

async function exportAllocation({ source, school, schoolId, fiscalYear, annual, aligned }) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("MOOE Allocation");
  sheet.addRow(["SCHOOL MOOE ALLOCATION"]);
  sheet.addRow([school || "School"]);
  sheet.addRow([`School ID: ${schoolId || "—"} | Fiscal year: ${fiscalYear || "—"}`]);
  sheet.addRow([`Prepared in ${source}`]);
  sheet.addRow(["Annual NEP allocation", annual]);
  sheet.addRow([]);
  sheet.addRow(["MOOE Category", "Default NEP %", "Default allocation", "Realignment", "Aligned allocation", "Aligned %"]);
  aligned.forEach(({ name, rate, amount }) => sheet.addRow([name, rate, annual * rate, amount - annual * rate, amount, annual ? amount / annual : 0]));
  const sum = (pick) => aligned.reduce((total, row) => total + pick(row), 0);
  sheet.addRow(["TOTAL", sum((row) => row.rate), sum((row) => annual * row.rate), sum((row) => row.amount - annual * row.rate), sum((row) => row.amount), annual ? sum((row) => row.amount) / annual : 0]);
  sheet.columns = [{ width: 42 }, { width: 18 }, { width: 22 }, { width: 20 }, { width: 22 }, { width: 18 }];
  sheet.getRow(1).font = { bold: true, size: 16, color: { argb: "FFFFFFFF" } };
  sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF173A4A" } };
  sheet.getRow(7).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(7).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF356C66" } };
  for (let row = 8; row <= sheet.rowCount; row += 1) {
    sheet.getCell(row, 2).numFmt = "0.00%";
    [3, 4, 5].forEach((col) => { sheet.getCell(row, col).numFmt = '"₱"#,##0.00;-"₱"#,##0.00'; });
    sheet.getCell(row, 6).numFmt = "0.00%";
  }
  return saveWorkbook(workbook, "Save MOOE allocation", "MOOE allocation.xlsx");
}

async function exportApp({ source, school, schoolId, fiscalYear, items }) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Quarterly APP", { views: [{ state: "frozen", ySplit: 5, xSplit: 6 }] });
  sheet.mergeCells("A1:N1");
  sheet.getCell("A1").value = "ANNUAL PROCUREMENT PLAN — QUARTERLY CONSOLIDATION";
  sheet.getCell("A1").font = { bold: true, size: 15, color: { argb: "FFFFFFFF" } };
  sheet.getCell("A1").fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF173A4A" } };
  sheet.getCell("A1").alignment = { horizontal: "center" };
  sheet.mergeCells("A2:N2");
  sheet.getCell("A2").value = `${school} | School ID: ${schoolId || "—"} | FY ${fiscalYear || "—"}`;
  sheet.getCell("A2").font = { bold: true, size: 12 };
  sheet.mergeCells("A3:N3");
  sheet.getCell("A3").value = `Consolidated from monthly SOB workbook: ${source}`;
  sheet.getCell("A3").font = { italic: true, color: { argb: "FF687983" } };
  for (let quarter = 0; quarter < 4; quarter += 1) {
    const start = 7 + quarter * 2;
    sheet.mergeCells(4, start, 4, start + 1);
    sheet.getCell(4, start).value = `${quarter + 1}${["st", "nd", "rd", "th"][quarter]} Quarter`;
  }
  const headings = ["MOOE Category", "Item Description", "Total Qty", "Unit", "Unit Cost", "Total Amount", "Qty", "Amount", "Qty", "Amount", "Qty", "Amount", "Qty", "Amount"];
  sheet.addRow([]);
  sheet.addRow(headings);
  sheet.getRow(5).font = { bold: true, color: { argb: "FFFFFFFF" } };
  sheet.getRow(5).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF356C66" } };
  const sorted = [...items].sort((a, b) => a.category.localeCompare(b.category) || a.description.localeCompare(b.description));
  for (const item of sorted) {
    const totalQty = item.quantities.reduce((sum, quantity) => sum + quantity, 0);
    const totalAmount = item.amounts.reduce((sum, amount) => sum + amount, 0);
    sheet.addRow([item.category, item.description, totalQty, item.unit, totalQty ? totalAmount / totalQty : 0, totalAmount,
      item.quantities[0], item.amounts[0], item.quantities[1], item.amounts[1], item.quantities[2], item.amounts[2], item.quantities[3], item.amounts[3]]);
  }
  sheet.columns = [30, 34, 12, 12, 16, 18, 10, 16, 10, 16, 10, 16, 10, 16].map((width) => ({ width }));
  for (let row = 6; row <= sheet.rowCount; row += 1) {
    for (const col of [5, 6, 8, 10, 12, 14]) sheet.getCell(row, col).numFmt = '"₱"#,##0.00';
    for (const col of [3, 7, 9, 11, 13]) sheet.getCell(row, col).numFmt = "#,##0.##";
  }
  return saveWorkbook(workbook, "Create quarterly APP workbook", `${String(school).replace(/[\\/:*?"<>|]/g, "-")} - Quarterly APP.xlsx`);
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
  ipcMain.handle("allocation:save", (_event, payload) => exportAllocation(payload));
  ipcMain.handle("app:export", (_event, payload) => exportApp(payload));
  ipcMain.on("budget:load", (event) => {
    try { event.returnValue = loadBudgetRow(); } catch (error) { console.error(error); event.returnValue = null; }
  });
  ipcMain.on("budget:path", (event) => { event.returnValue = databaseFile(); });
  ipcMain.handle("budget:save", (_event, json) => saveBudgetRow(json));
  ipcMain.handle("price:search", (_event, item) => {
    const query = String(item || "").trim().slice(0, 180);
    if (!query) throw new Error("Enter an item description before checking local prices.");
    return shell.openExternal(`https://www.google.com/search?hl=en-PH&gl=ph&q=${encodeURIComponent(`${query} price Philippines`)}`);
  });
  createWindow();
  app.on("activate", () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});

app.on("window-all-closed", () => { if (process.platform !== "darwin") app.quit(); });
