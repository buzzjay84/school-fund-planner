// The Excel reports' formulas: month sheets chain through the BREAKDOWN sheet, and the APP adds quarters.
import assert from "node:assert/strict";
import test from "node:test";
import { appWorkbook, colName, sobWorkbook } from "../src/main/reports.js";

const data = {
  school: "Test Elementary School", schoolId: "1", fiscalYear: 2026, nep: 1200000, abbr: "TES", level: "(Elementary School)", district: "East District I",
  signatories: { preparedBy: "ANA CRUZ", preparedPosition: "Principal I", sobApprovedBy: "JUAN DELA CRUZ, Ed.D" },
  categories: [{ name: "Office Supplies Expense", rate: 0.5, aligned: 500000 }, { name: "Water Expense", rate: 0.5, aligned: 700000 }],
  realignments: [{ values: { 0: -100000, 1: 100000 } }],
  months: Array.from({ length: 12 }, (_, month) => [
    { index: 0, items: [{ quantity: 10, unit: "ream", description: "Bond Paper A4", unitCost: 285 }] },
    ...(month === 0 ? [] : [{ index: 1, items: [{ quantity: 1, unit: "month", description: "Water Bill", unitCost: 5000 }] }]),
  ]),
  lines: [
    { category: "Office Supplies Expense", description: "Bond Paper A4", unit: "ream", unitCost: 285, quantities: [30, 30, 30, 30] },
    { category: "Water Expense", description: "Water Bill", unit: "month", unitCost: 5000, quantities: [2, 3, 3, 3] },
  ],
};
const formula = (sheet, address) => sheet.getCell(address).value?.formula;

test("column names", () => {
  assert.equal(colName(1), "A");
  assert.equal(colName(26), "Z");
  assert.equal(colName(27), "AA");
});

test("SOB: BREAKDOWN plus JAN to DEC, linked by formulas", () => {
  const workbook = sobWorkbook(data);
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), ["BREAKDOWN 2026", "JAN", "FEB", "MAR", "APR", "MAY", "JUNE", "JULY", "AUG", "SEPT", "OCT", "NOV", "DEC"]);
  const breakdown = workbook.getWorksheet("BREAKDOWN 2026");
  // One realignment row (7), then TOTAL 8, PROPOSE 9, Difference 10, PERCENTAGE 11, January 12/13.
  assert.equal(breakdown.getCell("C7").value, -100000);
  assert.equal(breakdown.getCell("C7").numFmt.startsWith("[Color10]"), true); // + green, − red
  assert.equal(formula(breakdown, "C8"), "SUM(C6:C7)");
  assert.equal(breakdown.getCell("C9").value, 500000);
  assert.equal(formula(breakdown, "C10"), "C8-C9");
  assert.equal(breakdown.getCell("B12").value, "JANUARY");
  assert.equal(formula(breakdown, "C13"), "C9-C12");
  assert.equal(formula(breakdown, "C15"), "C13-C14");

  const jan = workbook.getWorksheet("JAN");
  assert.equal(formula(jan, "D12"), "'BREAKDOWN 2026'!C9"); // January starts from the proposed amount
  assert.equal(formula(jan, "J13"), "F13*I13");
  assert.equal(formula(jan, "J15"), "SUM(J13:J14)"); // item plus one blank row for adding in Excel
  assert.equal(formula(breakdown, "C12"), "'JAN'!J15");
  assert.equal(breakdown.getCell("D12").value, 0); // Water not on the January sheet

  const feb = workbook.getWorksheet("FEB");
  assert.equal(formula(feb, "D12"), "'BREAKDOWN 2026'!C13"); // carries January's balance
  assert.equal(formula(feb, "D11"), "'BREAKDOWN 2026'!E13");
  const names = [];
  feb.eachRow((row) => row.eachCell((cell) => { if (cell.font?.bold && /^[A-Z]/.test(String(cell.value))) names.push(cell.value); }));
  assert.ok(names.includes("ANA CRUZ") && names.includes("JUAN DELA CRUZ, Ed.D"));
});

test("APP: quarter quantities add up and amounts multiply", () => {
  const sheet = appWorkbook(data).worksheets[0];
  assert.equal(sheet.getCell("A14").value.startsWith("Supplies and Materials"), true);
  assert.equal(formula(sheet, "C15"), "G15+I15+K15+M15");
  assert.equal(formula(sheet, "F15"), "C15*E15");
  assert.equal(formula(sheet, "H15"), "E15*G15");
  assert.equal(sheet.getCell("A18").value, "Utility Expense");
  assert.equal(formula(sheet, "F21"), "F16+F20"); // Total of the two sub-totals
});
