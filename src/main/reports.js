// Excel reports in the same layout as the school workbooks (SOB IECES 2026.xlsx, APP 2023 JHS.xlsx),
// with live formulas: change a quantity or unit cost and the subtotals, BREAKDOWN balances and next months follow.
import ExcelJS from "exceljs";

const ACCOUNTING = '_-* #,##0.00_-;-* #,##0.00_-;_-* "-"??_-;_-@_-';
// Realignment amounts: additions green, deductions red. In the number format, so the color follows the sign when edited in Excel.
const SIGNED = '[Color10]_-* #,##0.00_-;[Red]-* #,##0.00_-;_-* "-"??_-;_-@_-';
const THIN = { style: "thin" };
const BOX = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const YELLOW = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
const GREEN = { type: "pattern", pattern: "solid", fgColor: { argb: "FF92D050" } };
const BLUE = { type: "pattern", pattern: "solid", fgColor: { argb: "FF00B0F0" } };
const MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const SHEETS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUNE", "JULY", "AUG", "SEPT", "OCT", "NOV", "DEC"];
const REGION = "Region IX, Zamboanga Peninsula";

export const colName = (index) => { // 1 → A, 27 → AA
  let name = "";
  for (let n = index; n > 0; n = Math.floor((n - 1) / 26)) name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  return name;
};
const ref = (sheet, address) => `'${sheet.replace(/'/g, "''")}'!${address}`;

function put(sheet, address, value, style = {}) {
  const cell = sheet.getCell(address);
  cell.value = value;
  Object.assign(cell, style);
  return cell;
}

function title(sheet, row, lastCol, text, font, height) {
  sheet.mergeCells(row, 1, row, lastCol);
  put(sheet, `A${row}`, text, { font, alignment: { horizontal: "center", vertical: "middle" } });
  if (height) sheet.getRow(row).height = height;
}

// BREAKDOWN sheet: D.O. breakdown of the NEP, realignment, proposed (aligned) amounts and percentages.
// With months, it also carries each month's spending per category (linked to the month sheets) and the running balance.
// Returns the rows other sheets link to.
// Row layout of the BREAKDOWN sheet: one realignment row per recorded move (−/+ pairs like the school workbook),
// so everything below them shifts with the number of moves.
function breakdownLayout(data) {
  const realign = 7; const moves = Math.max(1, data.realignments?.length || 0);
  const total = realign + moves; const proposed = total + 1; const diff = total + 2; const pct = total + 3;
  const balanceRows = MONTHS.map((_, month) => pct + 2 + month * 2);
  return { name: `BREAKDOWN ${data.fiscalYear}`, first: 3, T: colName(3 + data.categories.length), realign, moves, total, proposed, diff, pct, balanceRows };
}

function breakdownSheet(workbook, data, monthLinks) {
  const layout = breakdownLayout(data);
  const { name, realign, total, proposed, diff, pct } = layout;
  const sheet = workbook.addWorksheet(name, { views: [{ state: "frozen", xSplit: 2 }] });
  const count = data.categories.length;
  const first = 3; const last = first + count - 1; const totalCol = last + 1;
  const C = colName; const T = colName(totalCol); const L = colName(last);
  const base = { font: { name: "Calibri", size: 11 }, alignment: { horizontal: "center" }, border: BOX, numFmt: ACCOUNTING };
  const label = (row, text, style) => put(sheet, `B${row}`, text, { ...base, ...style });

  sheet.mergeCells(1, 2, 1, totalCol); sheet.mergeCells(2, 2, 2, totalCol);
  put(sheet, "B1", `${data.abbr} MOOE ${data.fiscalYear}`, { font: { name: "Calibri", size: 48, bold: true }, alignment: { horizontal: "center", vertical: "middle" } });
  put(sheet, "B2", "REALLIGNMENT", { font: { name: "Calibri", size: 48, bold: true }, alignment: { horizontal: "center", vertical: "middle" } });
  sheet.getRow(1).height = 61.5; sheet.getRow(2).height = 61.5; sheet.getRow(3).height = 51;
  put(sheet, "B3", "MOOE", { font: { name: "Calibri", size: 22, bold: true }, alignment: { horizontal: "center", vertical: "middle" }, border: BOX });
  sheet.mergeCells("B4:B5");
  put(sheet, "B4", data.nep, { ...base, font: { name: "Calibri", size: 18 } });
  label(6, "D.O. BREAKDOWN", { fill: GREEN });
  const moveRows = data.realignments?.length ? data.realignments : [{ values: {} }];
  moveRows.forEach((move, at) => label(realign + at, move.label || null));
  label(total, "TOTAL", { fill: BLUE });
  label(proposed, "PROPOSE REALLIGNMENT", { font: { name: "Calibri", size: 11, bold: true }, fill: YELLOW, alignment: { horizontal: "center", wrapText: true } });
  label(diff, "Difference", { font: { name: "Calibri", size: 11, bold: true }, fill: YELLOW });
  label(pct, "PERCENTAGE", { fill: YELLOW });

  data.categories.forEach((category, index) => {
    const col = C(first + index);
    put(sheet, `${col}3`, category.name.toUpperCase(), { font: { name: "Calibri", size: 10, bold: true }, alignment: { horizontal: "center", vertical: "middle", wrapText: true }, border: BOX });
    sheet.mergeCells(`${col}4:${col}5`);
    put(sheet, `${col}4`, category.rate, { ...base, font: { name: "Calibri", size: 11, bold: true }, fill: YELLOW, alignment: { horizontal: "center", vertical: "middle" }, numFmt: "0.00%" });
    put(sheet, `${col}6`, { formula: `$B$4*${col}4` }, { ...base, fill: GREEN });
    moveRows.forEach((move, at) => put(sheet, `${col}${realign + at}`, move.values[index] || null, { ...base, font: { name: "Calibri", size: 11, bold: true }, numFmt: SIGNED }));
    put(sheet, `${col}${total}`, { formula: `SUM(${col}6:${col}${total - 1})` }, { ...base, fill: BLUE });
    // Typed like the school workbook: the month sheets follow this row, and Difference flags a mismatch with TOTAL.
    put(sheet, `${col}${proposed}`, category.aligned, { ...base, font: { name: "Calibri", size: 11, bold: true }, fill: YELLOW });
    put(sheet, `${col}${diff}`, { formula: `${col}${total}-${col}${proposed}` }, { ...base, font: { name: "Calibri", size: 11, bold: true }, fill: YELLOW });
    put(sheet, `${col}${pct}`, { formula: `IF($B$4=0,0,${col}${proposed}/$B$4)` }, { ...base, fill: YELLOW, numFmt: "0.00%" });
  });
  put(sheet, `${T}3`, "TOTAL", { font: { name: "Calibri", size: 10, bold: true }, alignment: { horizontal: "center", vertical: "middle" }, border: BOX });
  sheet.mergeCells(`${T}4:${T}5`);
  put(sheet, `${T}4`, { formula: `SUM(${C(first)}4:${L}5)` }, { ...base, numFmt: "0.00%" });
  for (let row = 6; row <= diff; row += 1) {
    const moveRow = row >= realign && row < total;
    put(sheet, `${T}${row}`, { formula: `SUM(${C(first)}${row}:${L}${row})` }, { ...base, font: { name: "Calibri", size: 11, bold: !moveRow }, numFmt: moveRow ? SIGNED : ACCOUNTING, fill: row === 6 ? GREEN : row === total ? BLUE : row >= proposed ? YELLOW : undefined });
  }
  put(sheet, `${T}${pct}`, { formula: `SUM(${C(first)}${pct}:${L}${pct})` }, { ...base, fill: YELLOW, numFmt: "0.00%" });
  // Monthly release (NEP ÷ 12) and the same less 5% tax, as in the school workbook.
  const V = C(totalCol + 1); const W = C(totalCol + 2);
  put(sheet, `${V}5`, "monthly"); put(sheet, `${W}5`, "with less tax");
  put(sheet, `${V}6`, { formula: "B4/12" }, { numFmt: ACCOUNTING });
  put(sheet, `${W}6`, { formula: `${V}6-(${V}6*0.05)` }, { numFmt: ACCOUNTING });

  monthLinks.forEach((links, month) => {
    const bal = layout.balanceRows[month]; const row = bal - 1;
    label(row, MONTHS[month].toUpperCase());
    label(bal, "BALANCE", { font: { name: "Calibri", size: 11, bold: true, color: { argb: "FFFF0000" } }, fill: YELLOW });
    data.categories.forEach((_, index) => {
      const col = C(first + index);
      put(sheet, `${col}${row}`, links[index] ? { formula: ref(SHEETS[month], links[index]) } : 0, base);
      put(sheet, `${col}${bal}`, { formula: `${col}${month ? bal - 2 : proposed}-${col}${row}` }, { ...base, font: { name: "Calibri", size: 11, bold: true, color: { argb: "FFFF0000" } }, fill: YELLOW });
    });
    put(sheet, `${T}${row}`, { formula: `SUM(${C(first)}${row}:${L}${row})` }, base);
    put(sheet, `${T}${bal}`, { formula: `SUM(${C(first)}${bal}:${L}${bal})` }, { ...base, font: { name: "Calibri", size: 11, bold: true, color: { argb: "FFFF0000" } }, fill: YELLOW });
    put(sheet, `${V}${row}`, { formula: `${T}${row}-${V}6` }, { numFmt: ACCOUNTING }); // above 0 = over the monthly release
  });
  sheet.getColumn(1).width = 0.7; sheet.getColumn(2).width = 25.6;
  for (let col = first; col <= totalCol + 2; col += 1) sheet.getColumn(col).width = 15;
  sheet.pageSetup = { paperSize: 14, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 1, horizontalCentered: true, printArea: `A1:${T}${layout.balanceRows[11]}` };
  return layout;
}

// Signature block: [label, column] pairs on one row, then names (bold) and lines below them three rows down.
function signatures(sheet, row, blocks, font) {
  blocks.forEach(([text, from, to, person]) => {
    put(sheet, `${colName(from)}${row}`, text, { font: { ...font }, alignment: { horizontal: "left" } });
    (person || []).forEach((line, at) => {
      const r = row + 3 + at;
      sheet.mergeCells(r, from, r, to);
      put(sheet, `${colName(from)}${r}`, line || "", { font: { ...font, bold: at === 0 }, alignment: { horizontal: "center" } });
    });
  });
}

function monthSheet(workbook, data, breakdown, month, groups) {
  const sheet = workbook.addWorksheet(SHEETS[month]);
  const f = (size, extra) => ({ name: "Calibri", size, ...extra });
  title(sheet, 1, 11, "Republic of the Philippines", f(12), 14.1);
  title(sheet, 2, 11, "Department of Education", f(12), 14.1);
  title(sheet, 3, 11, REGION, f(11), 14.1);
  title(sheet, 4, 11, "Isabela City, Division", f(11), 14.1);
  title(sheet, 5, 11, data.district || "", f(11), 14.1);
  title(sheet, 6, 11, data.school.toUpperCase(), f(16, { bold: true }), 16.5);
  title(sheet, 7, 11, "SCHOOL OPERATING BUDGET", f(11, { bold: true }), 14.1);
  title(sheet, 8, 11, `for the Month of ${MONTHS[month]} ${data.fiscalYear}`, f(10, { bold: true, color: { argb: "FF00B0F0" } }), 14.1);
  const heads = ["AIP Item No.", "Page", "Annual Budget", "Balance for the month", "Stock No.", "QTY.", "Unit of Issue", "Item Desciption", "Estimated Unit Cost", "Estimated Cost", "Running Balance"];
  heads.forEach((text, index) => put(sheet, `${colName(index + 1)}10`, text, { font: f(10, { bold: true }), alignment: { horizontal: "center", vertical: "middle", wrapText: true }, border: BOX }));
  sheet.getRow(10).height = 36.75;
  const cell = { font: f(10), alignment: { horizontal: "center" }, border: BOX };
  const money = { ...cell, numFmt: ACCOUNTING };
  const B = (address) => ref(breakdown.name, address);
  const prevBalance = (col) => B(`${col}${month ? breakdown.balanceRows[month - 1] : breakdown.proposed}`);
  put(sheet, "C11", { formula: B("B4") }, { ...money, font: f(10, { bold: true }) });
  put(sheet, "D11", { formula: month ? B(`${breakdown.T}${breakdown.balanceRows[month - 1]}`) : "C11" }, money);
  ["A", "B", "E", "F", "G", "H", "I", "J", "K"].forEach((col) => put(sheet, `${col}11`, null, cell));

  const links = [];
  let row = 12;
  for (const { index, items } of groups) {
    const category = data.categories[index];
    const col = colName(breakdown.first + index);
    const head = row;
    ["A", "B", "C"].forEach((c) => put(sheet, `${c}${row}`, null, cell));
    put(sheet, `D${row}`, { formula: prevBalance(col) }, money);
    sheet.mergeCells(`E${row}:K${row}`);
    put(sheet, `E${row}`, category.name.toUpperCase(), { font: f(10, { bold: true }), alignment: { horizontal: "left" }, border: BOX });
    // Every item, plus one blank line inside the SUM range for adding an item in Excel.
    [...items, null].forEach((item, at) => {
      row += 1;
      ["A", "B", "C", "D", "E", "F", "G", "H", "I", "K"].forEach((c) => put(sheet, `${c}${row}`, null, cell));
      if (item) {
        put(sheet, `A${row}`, item.aipNo || null, cell); put(sheet, `B${row}`, item.page || null, cell);
        put(sheet, `E${row}`, at + 1, cell); put(sheet, `F${row}`, item.quantity, cell);
        put(sheet, `G${row}`, item.unit, cell); put(sheet, `H${row}`, item.description, { ...cell, alignment: { horizontal: "left" } });
        put(sheet, `I${row}`, item.unitCost, money);
      }
      put(sheet, `J${row}`, { formula: `F${row}*I${row}` }, money);
    });
    row += 1;
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((c) => put(sheet, `${c}${row}`, null, cell));
    put(sheet, `I${row}`, "Subtotal", { ...cell, alignment: { horizontal: "right" } });
    put(sheet, `J${row}`, { formula: `SUM(J${head + 1}:J${row - 1})` }, { ...money, font: f(10, { bold: true }) });
    put(sheet, `K${row}`, { formula: `D${head}-J${row}` }, { ...money, font: f(10, { bold: true }) });
    links[index] = `J${row}`;
    row += 1;
  }
  put(sheet, `L${row}`, { formula: B(`${colName(breakdown.first + data.categories.length + 1)}6`) }, { font: f(11, { color: { argb: "FFFFC000" } }), fill: YELLOW, numFmt: ACCOUNTING });
  put(sheet, `M${row}`, "monthly allocation");
  row += 1;
  const subtotals = groups.map(({ index }) => links[index]);
  const yellowBold = { font: f(10, { bold: true }), fill: YELLOW, border: BOX, alignment: { horizontal: "center" }, numFmt: ACCOUNTING };
  put(sheet, `I${row}`, "GRAND TOTAL", { ...yellowBold, alignment: {} });
  put(sheet, `J${row}`, subtotals.length ? { formula: subtotals.join("+") } : 0, yellowBold);
  // Balance of every category, including ones not listed this month, from the BREAKDOWN sheet.
  put(sheet, `K${row}`, { formula: B(`${breakdown.T}${breakdown.balanceRows[month]}`) }, yellowBold);
  put(sheet, `L${row}`, { formula: `J${row}-L${row - 1}` }, { numFmt: ACCOUNTING });

  const s = data.signatories;
  signatures(sheet, row + 3, [
    ["Prepared by:", 1, 6, [s.preparedBy, s.preparedPosition]],
    ["Recommending Approval:", 8, 11, [s.sobRecommendedBy, s.sobRecommendedPosition, s.sobRecommendedOffice]],
  ], f(10));
  signatures(sheet, row + 9, [["Approved by:", 1, 6, [s.sobApprovedBy, s.sobApprovedPosition, s.sobApprovedOffice]]], f(10));
  [5.9, 5, 12.9, 18.7, 5.7, 6.1, 9.1, 22.6, 11.6, 12.3, 15.4, 22.6].forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
  sheet.pageSetup = { paperSize: 14, orientation: "portrait", fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true, printArea: `A1:K${row + 14}`, margins: { left: 0.12, right: 0.12, top: 0.39, bottom: 0.2, header: 0.31, footer: 0.31 } };
  return links;
}

// data.months[m] = [{ index (into data.categories), items: [{ aipNo, page, quantity, unit, description, unitCost }] }]
export function sobWorkbook(data) {
  const workbook = new ExcelJS.Workbook();
  // The BREAKDOWN sheet links to the month sheets' subtotals, which are only known once those are laid out:
  // lay them out first in a scratch workbook to learn each subtotal's cell, then build the real one in order.
  const scratch = new ExcelJS.Workbook();
  const draft = breakdownLayout(data);
  const links = data.months.map((groups, month) => monthSheet(scratch, data, draft, month, groups));
  const breakdown = breakdownSheet(workbook, data, links);
  data.months.forEach((groups, month) => monthSheet(workbook, data, breakdown, month, groups));
  return workbook;
}

// APP sections as in the school APP workbook; an expense that fits none of them goes under Other MOOE.
const APP_GROUPS = [
  ["Supplies and Materials Expense (Office Supplies Expense,Other Supplies,Fuel,Oil & Lubricants Expense)", /suppl|fuel|semi-expendable|forms|textbook|food|drugs|medical|chalk|zoological/i],
  ["Utility Expense", /water|electric/i],
  ["Communication Expense", /mobile|internet|landline|postage|communication|telephone/i],
  ["Repair and Maintenance Expense (Office Equipment, Fidelity Bond & Representation/Security/Maintenance)", /r&m|repair|fidelity|representation|security|janitorial|transportation and delivery|general services/i],
  ["Travelling Expense", /travel/i],
  ["Training Expense", /training/i],
  ["Other Maintenance and Operating Expenses", /./],
];

// data.lines = [{ category, description, unit, unitCost, quantities: [q1, q2, q3, q4] }]
export function appWorkbook(data) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(data.level ? data.level.replace(/[()]/g, "").slice(0, 31) : "APP");
  const f = (size, extra) => ({ name: "Arial", size, ...extra });
  title(sheet, 1, 14, "Republic of the Philippines", f(11), 14.25);
  title(sheet, 2, 14, "DEPARTMENT OF EDUCATION", f(11, { bold: true }), 15.75);
  title(sheet, 3, 14, REGION, f(11), 14.25);
  title(sheet, 4, 14, "ISABELA CITY SCHOOLS DIVISION", f(11), 14.25);
  title(sheet, 5, 14, "Isabela City, Basilan", f(11), 14.25);
  title(sheet, 6, 14, data.school.toUpperCase(), f(12, { bold: true }), 15.75);
  title(sheet, 7, 14, `School ID: ${data.schoolId || ""}`, f(11), 14.25);
  title(sheet, 8, 14, `ANNUAL PROCUREMENT PLAN F.Y. ${data.fiscalYear}`, f(12, { bold: true }), 15.75);
  title(sheet, 9, 14, data.level || "", f(11, { bold: true, color: { argb: "FF00B0F0" } }), 17.25);
  const cell = { font: f(11), alignment: { horizontal: "center" }, border: BOX };
  const money = { ...cell, numFmt: ACCOUNTING };
  ["ITEM NO.", "ITEM DECRIPTION", "QTY", "UNIT", "UNIT COST", "TOTAL AMOUNT"].forEach((text, index) => {
    for (const row of [11, 12, 13]) put(sheet, `${colName(index + 1)}${row}`, row === 11 ? text : null, { ...cell, font: f(11, { bold: true }), alignment: { horizontal: "center", vertical: "middle", wrapText: true } });
  });
  sheet.mergeCells("G11:N11");
  put(sheet, "G11", "DISTRIBUTION", { ...cell, font: f(18, { bold: true }), alignment: { horizontal: "center", vertical: "middle" } });
  ["1st Qtr", "2nd Qtr", "3rd Qtr", "4th Qtr"].forEach((text, quarter) => {
    const col = 7 + quarter * 2;
    sheet.mergeCells(12, col, 12, col + 1);
    put(sheet, `${colName(col)}12`, text, { ...cell, font: f(11, { bold: true }) });
    put(sheet, `${colName(col)}13`, "QTY", cell);
    put(sheet, `${colName(col + 1)}13`, "Amount", cell);
  });
  sheet.getRow(11).height = 30;

  const sections = APP_GROUPS.map(([name]) => ({ name, lines: [] }));
  data.lines.forEach((line) => sections[APP_GROUPS.findIndex(([, pattern]) => pattern.test(line.category))].lines.push(line));
  const subtotalRows = [];
  let row = 14;
  for (const section of sections.filter((entry) => entry.lines.length)) {
    sheet.mergeCells(row, 1, row, 14);
    put(sheet, `A${row}`, section.name, { ...cell, font: f(11, { bold: true, italic: true }), alignment: { horizontal: "left" } });
    const start = row + 1;
    section.lines.forEach((line, at) => {
      row += 1;
      put(sheet, `A${row}`, at + 1, cell);
      put(sheet, `B${row}`, line.description, { ...cell, alignment: { horizontal: "left" } });
      put(sheet, `C${row}`, { formula: `G${row}+I${row}+K${row}+M${row}` }, cell);
      put(sheet, `D${row}`, line.unit, cell);
      put(sheet, `E${row}`, line.unitCost, { ...money, alignment: {} });
      put(sheet, `F${row}`, { formula: `C${row}*E${row}` }, { ...money, alignment: {} });
      line.quantities.forEach((quantity, quarter) => {
        const q = colName(7 + quarter * 2);
        put(sheet, `${q}${row}`, quantity || null, cell);
        put(sheet, `${colName(8 + quarter * 2)}${row}`, { formula: `E${row}*${q}${row}` }, money);
      });
    });
    row += 1;
    sheet.mergeCells(row, 1, row, 5);
    put(sheet, `A${row}`, "Sub-Total", { ...cell, font: f(11, { bold: true }) });
    put(sheet, `F${row}`, { formula: `SUM(F${start}:F${row - 1})` }, { ...money, font: f(11, { bold: true }), alignment: {} });
    ["G", "I", "K", "M"].forEach((c) => put(sheet, `${c}${row}`, null, cell));
    ["H", "J", "L", "N"].forEach((c) => put(sheet, `${c}${row}`, { formula: `SUM(${c}${start}:${c}${row - 1})` }, { ...money, font: f(11, { bold: true }) }));
    subtotalRows.push(row);
    row += 2;
  }
  row -= 1;
  sheet.mergeCells(row, 1, row, 5);
  put(sheet, `A${row}`, "Total", { ...cell, font: f(11, { bold: true }), alignment: { horizontal: "right" } });
  const sum = (c) => (subtotalRows.length ? { formula: subtotalRows.map((r) => `${c}${r}`).join("+") } : 0);
  put(sheet, `F${row}`, sum("F"), { ...money, font: f(11, { bold: true }), alignment: {} });
  ["G", "I", "K", "M"].forEach((c) => put(sheet, `${c}${row}`, null, cell));
  ["H", "J", "L", "N"].forEach((c) => put(sheet, `${c}${row}`, sum(c), { ...money, font: f(11, { bold: true }) }));

  const s = data.signatories;
  signatures(sheet, row + 4, [
    ["Prepared by:", 1, 3, [s.preparedBy, s.preparedPosition]],
    ["Recommending Approval:", 5, 8, [s.appRecommendedBy, s.appRecommendedPosition]],
    ["Approved:", 10, 14, [s.appApprovedBy, s.appApprovedPosition, s.appApprovedOffice]],
  ], f(11));
  [6.3, 34.1, 9, 10.3, 11.9, 16.4, 9.4, 12.9, 9.4, 12.9, 9.4, 12.9, 9.4, 12.9].forEach((width, index) => { sheet.getColumn(index + 1).width = width; });
  sheet.pageSetup = { paperSize: 14, orientation: "landscape", fitToPage: true, fitToWidth: 1, fitToHeight: 0, horizontalCentered: true, printArea: `A1:N${row + 9}`, margins: { left: 0.5, right: 0.5, top: 0.39, bottom: 0.12, header: 0.31, footer: 0.31 } };
  return workbook;
}
