import "./style.css";
import { SCHOOLS, isSecondary } from "./schools.js";

const app = document.querySelector("#app");
const STORAGE_KEY = "fund-planner-budget-v1";
const currency = new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP", minimumFractionDigits: 2 });
const money = (value) => currency.format(number(value));
const wholeCurrency = new Intl.NumberFormat("en-PH", { style: "currency", currency: "PHP", maximumFractionDigits: 0 });
// Round before formatting and turn -0 into 0, so a rounded amount never shows as "-₱0".
const wholeMoney = { format: (value) => wholeCurrency.format(Math.round(Number(value) || 0) || 0) };
const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const quarterForMonth = (month) => Math.floor(months.indexOf(month) / 3);
const newId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`;
const number = (value) => Math.max(0, Number(String(value ?? "").replace(/,/g, "")) || 0);
const amount = (value) => number(value).toLocaleString("en-PH", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const total = (values) => values.reduce((sum, value) => sum + number(value), 0);
// Division Office default breakdown of the NEP (SOB IECES 2026.xlsx, BREAKDOWN sheet), in percent.
const DEFAULT_MOOE = [
  ["Travelling Expense - Local", 13.99], ["Training Expense", 14.94], ["Office Supplies Expense", 17.92],
  ["Fuel, Oil and Lubricant Expense", 0.18], ["Other Supplies and Materials Expense", 23.02], ["Water Expense", 1.53],
  ["Electricity Expense", 14.39], ["Postage and Courier Expense", 0.75], ["Mobile Expense", 0.07],
  ["Internet Expense", 10.58], ["Janitorial Expense", 0.47], ["Security Expense", 0.51],
  ["R&M - School Building", 1.15], ["R&M - Transportation Equipment", 0.11], ["R&M - Plant, Property & Equipment", 0.04],
  ["Fidelity Bond Premium", 0.08], ["Representation Expense", 0.08], ["Transportation and Delivery Expense", 0.19],
];
// Units seen in the school SOB/APP workbooks, plus common supply units.
const UNITS = [
  ["Supplies", ["pcs", "ream", "box", "pack", "set", "bottle", "bag", "can", "roll", "pad", "dozen", "pair", "unit", "block"]],
  ["Measures", ["kg", "liter", "gallon", "meter", "bd ft"]],
  ["Services and time", ["month", "year", "day", "trip", "pax", "per head", "lot", "job"]],
];
// Other MOOE object codes (UACS) a school may add, as listed in the IECES admin portal (MooePage.jsx).
// Accounts already in the default breakdown are left out so nothing is listed twice.
const OTHER_MOOE = [
  ["Supplies & Materials", ["Accountable Forms Expenses", "Non-Accountable Forms Expenses", "Animal/Zoological Supplies Expenses", "Food Supplies Expenses", "Drugs and Medicines Expenses", "Medical, Dental & Lab Supplies", "Textbooks and Instructional Materials", "Teaching Supplies / Chalk Allowance"]],
  ["Semi-Expendable Property (below ₱50,000)", ["Semi-Expendable - Machinery", "Semi-Expendable - Office Equipment", "Semi-Expendable - ICT Equipment", "Semi-Expendable - Communications Equipment", "Semi-Expendable - Printing Equipment", "Semi-Expendable - Furniture & Fixtures", "Semi-Expendable - Other Property"]],
  ["Utilities & Communication", ["Telephone Expenses - Landline"]],
  ["Awards & General Services", ["Awards and Rewards Expenses", "Other General Services"]],
  ["Professional & Financial Services", ["Other Professional Services", "Insurance Expenses", "Bank Charges"]],
  ["Repairs & Maintenance", ["R&M - Furniture and Fixtures", "R&M - Office Equipment", "R&M - ICT Equipment"]],
  ["Other Operating Expenses", ["Advertising Expenses", "Printing and Publication Expenses", "Rent/Lease Expenses", "Other Maintenance and Operating Expenses"]],
];
// Categorical chart colors (dataviz reference palette, validated for color-blind separation), assigned in this fixed order.
const SERIES_COLORS = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300"];
const TAX_RATE = 0.05;
const cents = (value) => Math.round(value * 100) / 100 || 0; // || 0 turns -0 into 0 so it never prints as "-₱0"
const defaultAmount = (category) => cents(number(budget.profile.nep) * number(category.rate) / 100);
// allocation null = not realigned yet, so it follows the default amount (and any NEP change).
const alignedAmount = (category) => (category.allocation == null ? defaultAmount(category) : number(category.allocation));
const realignment = (diff) => (Math.abs(diff) < 0.005 ? `<span class="muted">—</span>` : `<span class="${diff > 0 ? "gain" : "loss"}">${diff > 0 ? "+" : "−"}${money(Math.abs(diff))}</span>`);
const findSchool = (name) => SCHOOLS.find((school) => school.name === name);
const escapeHtml = (value) => String(value ?? "").replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]));
const initials = (name) => name.split(/\s+/).filter((word) => /^[A-Z]/.test(word)).slice(0, 2).map((word) => word[0]).join("") || "?";
const logoBadge = (name, size) => {
  const file = findSchool(name)?.logo;
  return `<span class="logo-badge ${size}" aria-hidden="true">${escapeHtml(initials(name))}${file ? `<img src="./school_logo/${file}" alt="">` : ""}</span>`;
};

function emptyBudget() {
  return { profile: { school: "", schoolId: "", fiscalYear: new Date().getFullYear(), nep: 0 }, categories: [], items: [] };
}

// The budget is stored in SQLite by the main process (window.sob). The old browser-storage copy is read
// only when the database has no budget yet, so data entered before the switch moves over on first start.
function loadBudget() {
  try {
    const stored = JSON.parse(window.sob?.loadBudget?.() ?? localStorage.getItem(STORAGE_KEY));
    return stored?.profile && Array.isArray(stored.categories) && Array.isArray(stored.items) ? stored : emptyBudget();
  } catch { return emptyBudget(); }
}

let budget = loadBudget();
let screen = "setup";
let formMonth = months[0];
// A realignment waiting for the user to say where the difference comes from / goes: { id, allocation }.
let pendingMove = null;
let dashExpense = null; // expense shown in the dashboard detail panel
let appQuarter = null; // quarter (0–3) whose breakdown is open on the Quarterly APP screen

function saveBudget() {
  const json = JSON.stringify(budget);
  // Outside Electron (no preload) fall back to browser storage.
  if (!window.sob?.saveBudget) { localStorage.setItem(STORAGE_KEY, json); setStatus("Saved on this machine", "success"); return; }
  window.sob.saveBudget(json)
    .then(() => setStatus("Saved on this machine", "success"))
    .catch((error) => setStatus(`Not saved: ${error.message}`, "error"));
}

function setStatus(text, kind = "") {
  const node = document.querySelector("#status");
  if (!node) return;
  node.textContent = text;
  node.className = `status ${kind}`;
}

function shell() {
  app.innerHTML = `
    <aside class="sidebar">
      <div class="brand"><img class="brand-mark" src="./app_logo.png" alt=""><div><strong>Fund Planner</strong><small>School Budget Preparation</small></div></div>
      <div class="nav-label">BUDGET WORKFLOW</div>
      <button class="nav-link" data-screen="setup"><span class="nav-number">01</span><span><b>Budget setup</b><small>School fund and MOOE plan</small></span></button>
      <button class="nav-link" data-screen="sob"><span class="nav-number">02</span><span><b>Monthly SOB</b><small>Planned items and costs</small></span></button>
      <button class="nav-link" data-screen="app"><span class="nav-number">03</span><span><b>Quarterly APP</b><small>Procurement by quarter</small></span></button>
      <div class="nav-label nav-label-gap">OVERVIEW</div>
      <button class="nav-link" data-screen="dash"><span class="nav-number">◔</span><span><b>Fund dashboard</b><small>Movement of amounts</small></span></button>
      <div class="sidebar-bottom"><span id="side-logo"></span><div><small>BUDGET FILE</small><div id="side-source">Saved locally</div></div></div>
    </aside>
    <main class="main">
      <header class="topbar"><div class="top-actions"><span class="status" id="status">Saved on this machine</span><button id="new-budget" class="button ghost">Start new budget</button></div></header>
      <section class="content" id="content"></section>
    </main>`;
  // A missing logo file falls back to the initials underneath it (error events don't bubble, so capture).
  document.addEventListener("error", (event) => { if (event.target.closest?.(".logo-badge")) event.target.remove(); }, true);
  // Peso fields: group thousands while typing, pad to 2 decimals on leaving the field.
  document.addEventListener("input", (event) => { if (event.target.matches("[data-money]")) formatMoneyInput(event.target); });
  document.addEventListener("focusout", (event) => { if (event.target.matches("[data-money]") && event.target.value) event.target.value = amount(event.target.value); });
  document.querySelectorAll("[data-screen]").forEach((button) => button.addEventListener("click", () => { screen = button.dataset.screen; render(); }));
  document.querySelector("#new-budget").addEventListener("click", () => {
    if (!budget.categories.length && !budget.items.length && !budget.profile.school) return;
    if (window.confirm("Start a new budget? Your current data saved in Fund Planner will be replaced.")) {
      budget = emptyBudget(); screen = "setup"; saveBudget(); render();
    }
  });
}

function groupMoney(raw) {
  const [whole, ...decimals] = raw.replace(/[^\d.]/g, "").split(".");
  const grouped = whole.replace(/^0+(?=\d)/, "").replace(/\B(?=(\d{3})+(?!\d))/g, ",");
  return raw.includes(".") ? `${grouped}.${decimals.join("").slice(0, 2)}` : grouped;
}

function formatMoneyInput(input) {
  // Keep the caret after the same number of digits once commas move around.
  const digitsBefore = input.value.slice(0, input.selectionStart).replace(/[^\d.]/g, "").length;
  const text = groupMoney(input.value);
  input.value = text;
  let caret = 0;
  for (let seen = 0; caret < text.length && seen < digitsBefore; caret++) if (text[caret] !== ",") seen++;
  input.setSelectionRange(caret, caret);
}

function updateContext() {
  const school = budget.profile.school;
  document.querySelector("#side-source").textContent = school || "Saved locally";
  document.querySelector("#side-source").title = window.sob?.budgetPath?.() || "";
  document.querySelector("#side-logo").innerHTML = school ? logoBadge(school, "sm") : `<span class="source-dot"></span>`;
  document.querySelectorAll(".nav-link").forEach((button) => {
    const active = button.dataset.screen === screen;
    button.classList.toggle("active", active);
    if (active) button.setAttribute("aria-current", "page"); else button.removeAttribute("aria-current");
  });
}

function setupScreen() {
  const nep = number(budget.profile.nep);
  const allocated = total(budget.categories.map(alignedAmount));
  const defaultTotal = total(budget.categories.map(defaultAmount));
  // Round to centavos so float noise (e.g. 2,964,000.0000000005) never reads as over or −₱0.
  const remaining = cents(nep - allocated);
  const used = new Set(budget.categories.map((category) => (category.name || "").trim().toLowerCase()));
  const missingDefaults = DEFAULT_MOOE.filter(([name]) => !used.has(name.toLowerCase()));
  const otherExpenses = OTHER_MOOE.map(([group, names]) => [group, names.filter((name) => !used.has(name.toLowerCase()))]).filter(([, names]) => names.length);
  const percentOfNep = (value) => (nep ? `${(value / nep * 100).toFixed(2)}%` : "—");
  // Raising an expense must take the difference from somewhere (and lowering gives it somewhere), so the total stays at the NEP.
  const moveRow = (category) => {
    const target = pendingMove.allocation == null ? defaultAmount(category) : pendingMove.allocation;
    const delta = cents(target - alignedAmount(category));
    const others = budget.categories.filter((entry) => entry.id !== category.id);
    const options = delta > 0
      ? [...(remaining >= delta ? [`<option value="__remaining">Remaining to assign (${money(remaining)})</option>`] : []),
         ...others.filter((entry) => alignedAmount(entry) >= delta).map((entry) => `<option value="${entry.id}">${escapeHtml(entry.name || "Unnamed category")} (${money(alignedAmount(entry))})</option>`)]
      : [`<option value="__remaining">Leave in Remaining to assign</option>`, ...others.map((entry) => `<option value="${entry.id}">${escapeHtml(entry.name || "Unnamed category")} (${money(alignedAmount(entry))})</option>`)];
    return `<tr class="move-row"><td colspan="7"><div class="move-box">
      <span>${delta > 0 ? "Raise" : "Lower"} <b>${escapeHtml(category.name || "this expense")}</b> to <b>${money(target)}</b>:
        ${delta > 0 ? "take" : "give"} <b class="${delta > 0 ? "gain" : "loss"}">${money(Math.abs(delta))}</b> ${delta > 0 ? "from" : "to"}</span>
      ${options.length ? `<select id="move-source" aria-label="${delta > 0 ? "Take the amount from" : "Give the amount to"}">${options.join("")}</select><button class="button primary" id="move-apply">Apply</button>` : `<span class="loss">No expense has ${money(delta)} to spare. Lower the amount or free up money first.</span>`}
      <button class="button ghost" id="move-cancel">Cancel</button></div></td></tr>`;
  };
  const categoryRows = budget.categories.map((category) => `
    <tr>
      <td><input class="table-input name" data-category="${category.id}" data-field="name" value="${escapeHtml(category.name)}" aria-label="MOOE category"></td>
      <td><label class="amount-input rate"><input class="table-input" data-category="${category.id}" data-field="rate" type="number" min="0" max="100" step="0.01" value="${number(category.rate).toFixed(2)}" aria-label="Default percentage"><span>%</span></label></td>
      <td class="number">${money(defaultAmount(category))}</td>
      <td class="number">${realignment(alignedAmount(category) - defaultAmount(category))}</td>
      <td><label class="amount-input"><span>₱</span><input class="table-input" data-category="${category.id}" data-field="allocation" type="text" inputmode="decimal" data-money value="${amount(alignedAmount(category))}" placeholder="${amount(defaultAmount(category))}" aria-label="Aligned annual amount"></label></td>
      <td class="number">${percentOfNep(alignedAmount(category))}</td>
      <td><button class="icon-button remove-category" data-remove-category="${category.id}" aria-label="Remove category">×</button></td>
    </tr>${pendingMove?.id === category.id ? moveRow(category) : ""}`).join("");
  document.querySelector("#content").innerHTML = `
    <div class="page-heading"><div><span class="step-tag">STEP 01 <i>•</i> BUDGET PREPARATION</span><h1>Set up the school fund</h1><p>Enter the annual school allocation and decide how the MOOE fund will be assigned.</p></div><button class="button primary" id="export-allocation" ${budget.categories.length ? "" : "disabled"}>Export allocation</button></div>
    <section class="panel setup-panel"><div class="panel-heading">${budget.profile.school
      ? `<div class="school-identity">${logoBadge(budget.profile.school, "lg")}<div><h2>${escapeHtml(budget.profile.school)}</h2><p>School ID ${escapeHtml(budget.profile.schoolId) || "not assigned"} · Fiscal year ${number(budget.profile.fiscalYear) || "—"} · Appears on all exports</p></div></div>`
      : `<div><h2>School details</h2><p>Choose the school. Its name and ID appear in the budget and procurement exports.</p></div>`}</div>
      <div class="setup-fields">
        <label>School name<select data-profile="school">
          <option value="" ${budget.profile.school ? "" : "selected"} disabled>Select a school</option>
          ${budget.profile.school && !findSchool(budget.profile.school) ? `<option selected>${escapeHtml(budget.profile.school)}</option>` : ""}
          ${[["Elementary", false], ["Secondary", true]].map(([group, secondary]) => `<optgroup label="${group}">${SCHOOLS.filter((school) => isSecondary(school.name) === secondary).map((school) => `<option ${school.name === budget.profile.school ? "selected" : ""}>${escapeHtml(school.name)}</option>`).join("")}</optgroup>`).join("")}
        </select></label>
        <label>School ID<input data-profile="schoolId" value="${escapeHtml(budget.profile.schoolId)}" placeholder="Enter school ID"></label>
        <label>Fiscal year<input data-profile="fiscalYear" type="number" min="2020" max="2100" value="${number(budget.profile.fiscalYear) || new Date().getFullYear()}"></label>
        <label>Annual NEP allocation<div class="fund-field"><span>₱</span><input data-profile="nep" type="text" inputmode="decimal" data-money value="${amount(nep)}" placeholder="0.00"></div></label>
      </div>
    </section>
    <section class="overview-grid three"><article class="budget-card"><div class="card-kicker">ANNUAL NEP ALLOCATION</div><div class="large-value">${wholeMoney.format(nep)}</div><div class="card-foot">Enter the released annual amount</div></article><article class="stat-card"><div class="card-kicker">ALIGNED MOOE TOTAL</div><div class="stat-value">${wholeMoney.format(allocated)}</div><div class="card-foot">${budget.categories.length} categories entered</div></article><article class="stat-card${remaining < 0 ? " over" : ""}"><div class="card-kicker">${remaining < 0 ? "OVER ALLOCATION" : "REMAINING TO ASSIGN"}</div><div class="stat-value">${money(Math.abs(remaining))}</div><div class="card-foot">${remaining < 0 ? "Amount exceeds NEP allocation" : "Available for MOOE categories"}</div></article></section>
    <section class="panel allocation-panel"><div class="panel-heading"><div><h2>MOOE allocation</h2><p>The default % is the Division Office breakdown of the NEP. Realign by changing the aligned amount to fit what the school needs this year; clear an aligned amount to return it to the default.</p></div><select class="add-select" id="add-category" aria-label="Add a MOOE expense"><option value="" selected disabled>+ Add MOOE expense</option>${missingDefaults.length ? `<optgroup label="Default breakdown (restores its default %)">${missingDefaults.map(([name]) => `<option>${escapeHtml(name)}</option>`).join("")}</optgroup>` : ""}${otherExpenses.map(([group, names]) => `<optgroup label="${escapeHtml(group)}">${names.map((name) => `<option>${escapeHtml(name)}</option>`).join("")}</optgroup>`).join("")}<option value="__custom">Custom expense…</option></select></div>
      <div class="table-scroll"><table class="data-table allocation-table editable"><thead><tr><th>MOOE CATEGORY</th><th class="number">DEFAULT %</th><th class="number">DEFAULT AMOUNT</th><th class="number">REALIGNMENT</th><th class="number">ALIGNED AMOUNT</th><th class="number">ALIGNED %</th><th></th></tr></thead>
        <tbody>${categoryRows || `<tr><td colspan="7" class="empty-state">No MOOE categories yet.<br><button class="button primary" id="use-defaults">Use the default MOOE breakdown</button></td></tr>`}</tbody>
        ${budget.categories.length ? `<tfoot><tr><td>Total</td><td class="number">${total(budget.categories.map((category) => category.rate)).toFixed(2)}%</td><td class="number">${money(defaultTotal)}</td><td class="number">${realignment(allocated - defaultTotal)}</td><td class="number">${money(allocated)}</td><td class="number">${percentOfNep(allocated)}</td><td></td></tr></tfoot>` : ""}
      </table></div>
      <div class="table-total"><span>Aligned allocation total</span><strong>${money(allocated)}</strong><span class="balance-note${remaining < 0 ? " over" : ""}">${money(Math.abs(remaining))} ${remaining < 0 ? "over the NEP allocation" : "remaining"}</span></div>
    </section>`;
  document.querySelectorAll("[data-profile]").forEach((input) => input.addEventListener("change", () => {
    budget.profile[input.dataset.profile] = input.type === "number" || "money" in input.dataset ? number(input.value) : input.value;
    if (input.dataset.profile === "school") budget.profile.schoolId = findSchool(input.value)?.id ?? budget.profile.schoolId;
    saveBudget(); render();
  }));
  document.querySelector("#add-category").addEventListener("change", (event) => {
    const choice = event.target.value;
    const preset = DEFAULT_MOOE.find(([name]) => name === choice);
    const id = newId();
    // A restored default follows its default % again; any other expense starts at ₱0 for the school to align.
    budget.categories.push(preset ? { id, name: preset[0], rate: preset[1], allocation: null } : { id, name: choice === "__custom" ? "" : choice, rate: 0, allocation: 0 });
    saveBudget(); render();
    document.querySelector(`[data-category="${id}"][data-field="${choice === "__custom" ? "name" : "allocation"}"]`)?.focus();
  });
  document.querySelector("#use-defaults")?.addEventListener("click", () => {
    budget.categories = DEFAULT_MOOE.map(([name, rate]) => ({ id: newId(), name, rate, allocation: null })); saveBudget(); render();
  });
  document.querySelectorAll("[data-category]").forEach((input) => input.addEventListener("change", () => {
    const category = budget.categories.find((item) => item.id === input.dataset.category);
    const { field } = input.dataset;
    // An emptied aligned amount goes back to following the default breakdown.
    if (field === "allocation") {
      // An emptied aligned amount means "back to the default amount".
      const allocation = input.value.trim() ? number(input.value) : null;
      const target = allocation == null ? defaultAmount(category) : allocation;
      if (Math.abs(target - alignedAmount(category)) < 0.005) { category.allocation = allocation; saveBudget(); render(); return; }
      pendingMove = { id: category.id, allocation };
      render(); return;
    }
    category[field] = field === "name" ? input.value : number(input.value);
    saveBudget(); render();
  }));
  document.querySelectorAll("[data-remove-category]").forEach((button) => button.addEventListener("click", () => {
    const id = button.dataset.removeCategory;
    if (budget.items.some((item) => item.categoryId === id)) { setStatus("Remove or change the SOB items using this category first", "warning"); return; }
    const category = budget.categories.find((entry) => entry.id === id);
    if (alignedAmount(category) > 0 && !window.confirm(`Remove ${category.name || "this expense"}? Its ${money(alignedAmount(category))} goes back to Remaining to assign so you can realign it to other expenses.`)) return;
    budget.categories = budget.categories.filter((category) => category.id !== id); saveBudget(); render();
  }));
  document.querySelector("#move-cancel")?.addEventListener("click", () => { pendingMove = null; render(); });
  document.querySelector("#move-apply")?.addEventListener("click", () => {
    const category = budget.categories.find((entry) => entry.id === pendingMove.id);
    const source = budget.categories.find((entry) => entry.id === document.querySelector("#move-source").value);
    const delta = cents((pendingMove.allocation == null ? defaultAmount(category) : pendingMove.allocation) - alignedAmount(category));
    if (source) source.allocation = cents(alignedAmount(source) - delta);
    // Log each move like the BREAKDOWN sheet's paired −/+ rows, for the dashboard.
    const sourceName = source ? source.name || "Unnamed category" : "Remaining to assign";
    const categoryName = category.name || "Unnamed category";
    (budget.moves ||= []).push({ date: new Date().toISOString(), from: delta > 0 ? sourceName : categoryName, to: delta > 0 ? categoryName : sourceName, amount: Math.abs(delta) });
    category.allocation = pendingMove.allocation;
    pendingMove = null; saveBudget(); render();
  });
  document.querySelector("#move-source")?.focus();
  document.querySelector("#export-allocation")?.addEventListener("click", exportAllocation);
}

// Mirrors the BREAKDOWN sheet's month rows: each category's planned spending per month and what is left by December.
// Utility bills are planned automatically every month: aligned amount ÷ 12, with December taking
// the centavo remainder so the year adds up to the aligned amount exactly (balance reaches ₱0).
const UTILITY_BILLS = [[/water/i, "Water Bill"], [/electric/i, "Electricity Bill"], [/internet/i, "Internet Bill"], [/mobile/i, "Mobile / Communication Allowance"]];
function utilityItems() {
  return budget.categories.flatMap((category) => {
    const bill = UTILITY_BILLS.find(([pattern]) => pattern.test(category.name || ""))?.[1];
    const yearly = alignedAmount(category);
    if (!bill || !yearly) return [];
    const share = cents(yearly / 12);
    return months.map((month, index) => ({ id: `auto-${category.id}-${index}`, auto: true, month, categoryId: category.id, description: bill, quantity: 1, unit: "month", unitCost: index === 11 ? cents(yearly - share * 11) : share }));
  });
}
// Services set once in January and repeated every month (Feb–Dec); editing January changes them all.
const REPEAT_MONTHLY = /janitorial|security|representation|transportation and delivery/i;
function repeatedItems() {
  const repeating = new Set(budget.categories.filter((category) => REPEAT_MONTHLY.test(category.name || "")).map((category) => category.id));
  return budget.items.filter((item) => item.month === months[0] && repeating.has(item.categoryId))
    .flatMap((item) => months.slice(1).map((month, index) => ({ ...item, id: `rep-${item.id}-${index}`, month, auto: true, repeat: true })));
}
// Bank charges = purchase of checks: 2 checks at ₱500 each, every quarter (Jan, Apr, Jul, Oct).
function quarterlyItems() {
  return budget.categories.filter((category) => /bank charges/i.test(category.name || ""))
    .flatMap((category) => [0, 3, 6, 9].map((index) => ({ id: `qtr-${category.id}-${index}`, auto: true, quarterly: true, month: months[index], categoryId: category.id, description: "Purchase of Checks", quantity: 2, unit: "pcs", unitCost: 500 })));
}
// Semi-expendable property (below ₱50,000 per unit): typical items per kind as [name, unit, unit cost]; cost 0 = type it.
// Costs from the school APP 2023 workbook where it lists the item.
// Aircon by horsepower; only the 3.0 HP price is in the APP workbook, the rest are typed.
const AIRCON_SIZES = ["1.0", "1.5", "2.0", "2.5", "3.0"].map((hp) => [`Air Conditioner ${hp} HP`, "unit", hp === "3.0" ? 35000 : 0]);
const SEMI_EXPENDABLE_ITEMS = [
  [/machinery/i, [["Generator Set", "unit", 0], ["Water Pump", "unit", 8000], ["Grass Cutter", "unit", 0], ["Industrial Fan", "unit", 8000]]],
  [/office equipment/i, [...AIRCON_SIZES, ["Electric Fan", "unit", 0], ["Water Dispenser", "unit", 7500], ["Refrigerator", "unit", 0], ["Paper Shredder", "unit", 0], ["Bundy Clock", "unit", 0]]],
  [/ict equipment/i, [["Laptop", "unit", 36000], ["Desktop Computer", "unit", 0], ["Projector", "unit", 35000], ["Smart TV", "unit", 20000], ["CCTV", "unit", 20000], ["1TB SSD", "unit", 6000], ["Router / WiFi", "unit", 0]]],
  [/communications/i, [["Sound System", "set", 8000], ["Speaker", "unit", 0], ["Microphone", "unit", 0], ["Two-way Radio", "unit", 0], ["Telephone Unit", "unit", 0]]],
  [/printing/i, [["Printer", "unit", 11500], ["Printer 3 in 1", "unit", 15000], ["Photocopier", "unit", 0], ["Laminator", "unit", 0]]],
  [/furniture/i, [["Office Table", "unit", 0], ["Office Chair", "unit", 0], ["Filing Cabinet", "unit", 0], ["Bookshelf", "unit", 0], ["Whiteboard", "unit", 0], ["Monobloc Chair", "pcs", 0]]],
  [/semi-expendable/i, [["Television", "unit", 20000], ...AIRCON_SIZES, ["Industrial Fan", "unit", 8000], ["Water Dispenser", "unit", 7500], ["CCTV", "unit", 20000]]],
];
const SEMI_EXPENDABLE_LIMIT = 50000;
const semiItems = (category) => (/semi-expendable/i.test(category.name || "") ? SEMI_EXPENDABLE_ITEMS.find(([pattern]) => pattern.test(category.name))[1] : null);
const allItems = () => [...budget.items, ...utilityItems(), ...repeatedItems(), ...quarterlyItems()];
const itemCost = (item) => cents(number(item.quantity) * number(item.unitCost));
const monthSpent = (month) => total(allItems().filter((item) => item.month === month).map(itemCost));
// Office supplies the school buys every month: [description, unit, unit cost from the SOB/APP workbooks or 0 = type it].
const OFFICE_DEFAULTS = [
  ["Bond Paper A4", "ream", 285], ["Bond Paper Folio (Long)", "ream", 295], ["Ink Printer", "set", 1600], ["Vellum", "ream", 500],
  ["Folder", "pack", 240], ["Expanded Envelope", "pcs", 20], ["Ballpen", "box", 0], ["Staple Wire", "box", 0], ["Masking Tape", "roll", 0],
];
// Prepared item choices per MOOE expense: [pattern, default unit, items]. An item is a name (default unit, type the cost)
// or [name, unit, usual cost]. Costs come from the school SOB/APP workbooks. Order matters: specific patterns first.
const CATEGORY_HINTS = [
  [/r&m.*building|school bldg/i, "pcs", ["Faucet", ["Cement", "bag", 290], "Paint", "Plywood", "Electrical Wire", ["Labor", "job", 0]]],
  [/r&m.*transport/i, "job", ["Vehicle Repair", "Vehicle Maintenance", ["Tires", "pcs", 0]]],
  [/r&m.*furniture/i, "job", ["Chair Repair", "Table Repair", "Cabinet Repair"]],
  [/r&m.*ict/i, "job", ["Computer Repair", "Laptop Repair", "Printer Repair", "CCTV Repair", "Network Repair"]],
  [/r&m.*(office|plant|equipment)/i, "job", ["Aircon Cleaning", "Aircon Repair", "Electric Fan Repair", "Photocopier Repair", "Equipment Repair"]],
  [/travel/i, "pax", [["Travel Reimbursement", "pax", 5000], "Travel Allowance"]],
  [/training/i, "pax", [["Meals (Participants)", "pax", 350], ["Meals (Mentor)", "pax", 400], "Snacks", ["Training Materials", "pcs", 0]]],
  [/office supplies/i, "pcs", [...OFFICE_DEFAULTS, ["Trashbag", "pcs", 75], "Correction Tape", "Stapler", "Scissors", "Glue", "Record Book", "Fastener", ["Printer Ink Refill", "bottle", 0]]],
  [/other supplies/i, "unit", [["Printer 3 in 1", "unit", 15000], ["Learning Materials", "set", 0], ["Cleaning Materials", "set", 0], ["Sports Equipment", "set", 0], ["Electrical Supplies", "lot", 0]]],
  [/fuel/i, "liter", ["Gasoline", "Diesel", "Engine Oil"]],
  [/water/i, "month", ["Water Bill", ["Drinking Water Refill", "gallon", 0]]],
  [/electric/i, "month", ["Electricity Bill"]],
  [/postage|courier/i, "pcs", ["Mail", "Courier Fee"]],
  [/mobile/i, "month", ["Communication Allowance", "Load Card"]],
  [/landline/i, "month", ["Landline Bill"]],
  [/internet/i, "month", ["Internet Bill"]],
  [/janitorial/i, "pax", ["Utility Worker"]],
  [/security/i, "pax", ["Watchmen", "Security Guard"]],
  [/fidelity/i, "year", ["Fidelity Bond Premium"]],
  [/representation/i, "pax", ["Meals and Snacks", "Tokens"]],
  [/transportation and delivery|delivery/i, "trip", ["Delivery Fee", "Hauling"]],
  [/non-accountable forms/i, "pcs", ["School Forms (SF1–SF10)", "Report Card (SF9)", "Permanent Record (SF10)", "Certificates", "Diplomas"]],
  [/accountable forms/i, "pad", ["Official Receipts", "Disbursement Vouchers"]],
  [/animal|zoological/i, "kg", ["Animal Feeds", ["Veterinary Supplies", "set", 0]]],
  [/food supplies/i, "pack", ["Feeding Program Supplies", ["Rice", "kg", 0], "Snacks"]],
  [/drugs and medicines/i, "box", ["First Aid Kit", "Paracetamol", ["Alcohol", "bottle", 0], "Vitamins", "Antiseptic"]],
  [/medical, dental/i, "pcs", ["Thermometer", "Blood Pressure Monitor", ["Gloves", "box", 0], ["Face Masks", "box", 0], "Weighing Scale"]],
  [/textbooks/i, "pcs", ["Learning Modules", "Workbooks", "Reference Books", "Teaching Charts"]],
  [/teaching supplies|chalk/i, "pax", ["Chalk Allowance", ["Teaching Supplies", "set", 0]]],
  [/awards/i, "pcs", ["Certificates and Medals", "Plaques", "Trophies"]],
  [/general services/i, "month", ["Utility Services", "Gardening Services"]],
  [/professional services/i, "job", ["Professional Services", "Resource Speaker Honorarium"]],
  [/insurance/i, "year", ["Building Insurance", "Property Insurance"]],
  [/bank charges/i, "pcs", [["Purchase of Checks", "pcs", 500], ["Bank Service Fee", "month", 0]]],
  [/advertising/i, "pcs", ["Tarpaulin", "Posters"]],
  [/printing and publication/i, "pcs", ["Tarpaulin Printing", "Printing of Forms", "Bookbinding"]],
  [/rent|lease/i, "day", ["Venue Rental", "Equipment Rental", "Sound System Rental"]],
  [/other maintenance/i, "lot", ["Miscellaneous"]],
];
// The prepared choices for an expense as [name, unit, cost]; semi-expendable kinds use their own lists.
const itemChoices = (category) => {
  const semi = semiItems(category);
  if (semi) return semi;
  const [, unit, items] = categoryHint(category);
  return items.map((entry) => (Array.isArray(entry) ? entry : [entry, unit, 0]));
};
// Job-order workers (janitorial, security) are paid a fixed monthly rate: ₱10,000 down to ₱5,000.
const JOB_ORDER = /janitorial|security/i;
const JOB_ORDER_PAY = Array.from({ length: 11 }, (_, index) => 10000 - index * 500);
// Daily-paid job orders: ₱400 down to ₱250 a day, costed as 22 working days a month.
const WORK_DAYS = 22;
const JOB_ORDER_DAILY = [400, 350, 300, 250];
// Pay dropdown values: "8000" = monthly rate, "d400" = daily rate (× WORK_DAYS).
const payValue = (value) => (String(value).startsWith("d") ? { unitCost: number(String(value).slice(1)) * WORK_DAYS, dailyRate: number(String(value).slice(1)) } : { unitCost: number(value) });
// Pay note under a job-order unit cost: "400.00/day × 22 days" or "per month".
const payNote = (item, category) => (item.dailyRate ? `<small class="rate-note">${amount(item.dailyRate)}/day × ${WORK_DAYS} days</small>` : JOB_ORDER.test(category.name || "") ? `<small class="rate-note">per month</small>` : "");
const categoryHint = (category) => CATEGORY_HINTS.find(([pattern]) => pattern.test(category.name || "")) || [null, "", []];
const unitOptions = (preset = "") => `<option value="" ${preset ? "" : "selected"} disabled>Select</option>${UNITS.map(([group, units]) => `<optgroup label="${group}">${units.map((unit) => `<option ${unit === preset ? "selected" : ""}>${unit}</option>`).join("")}</optgroup>`).join("")}`;
let focusCategory = null;
// Teaching / non-teaching counts per month; a month without its own entry carries over the latest earlier month.
const staffFor = (month) => {
  for (let index = months.indexOf(month); index >= 0; index -= 1) {
    const entry = budget.staff?.[months[index]];
    if (entry) return { teaching: number(entry.teaching), nonTeaching: number(entry.nonTeaching), from: months[index] };
  }
  return { teaching: 0, nonTeaching: 0, from: null };
};

function utilizationPanel() {
  const items = allItems();
  const monthly = budget.categories.map((category) => months.map((month) => total(items.filter((item) => item.categoryId === category.id && item.month === month).map(itemCost))));
  const monthTotals = months.map((_, index) => total(monthly.map((row) => row[index])));
  const release = cents(number(budget.profile.nep) / 12);
  const balances = budget.categories.map((category, index) => cents(alignedAmount(category) - total(monthly[index])));
  const unplanned = cents(total(balances.filter((balance) => balance > 0)));
  const balanceCell = (balance) => balance > 0.004 ? `<span class="warn">${money(balance)} left</span>` : balance < -0.004 ? `<span class="loss">${money(-balance)} over</span>` : `<span class="gain">✓ Used up</span>`;
  const cell = (value) => (value ? amount(value) : `<span class="muted">—</span>`);
  const rows = budget.categories.map((category, index) => `<tr><td><span class="category-name">${escapeHtml(category.name || "Unnamed category")}</span></td><td class="number strong">${amount(alignedAmount(category))}</td>${monthly[index].map((value) => `<td class="number">${cell(value)}</td>`).join("")}<td class="number">${balanceCell(balances[index])}</td></tr>`).join("");
  return `<section class="panel app-panel"><div class="panel-heading"><div><h2>Fund utilization by month</h2><p>Planned spending per MOOE category from January to December. Every aligned fund should be used up by December.</p></div><span class="table-note">${unplanned > 0.004 ? `${wholeMoney.format(unplanned)} not yet planned` : "All funds planned"}</span></div>
    <div class="table-scroll"><table class="data-table utilization-table"><thead><tr><th>MOOE CATEGORY</th><th class="number">ALIGNED</th>${months.map((month) => `<th class="number">${month.slice(0, 3).toUpperCase()}</th>`).join("")}<th class="number">DECEMBER BALANCE</th></tr></thead>
      <tbody>${rows}</tbody>
      <tfoot><tr><td>Month total</td><td class="number">${amount(total(budget.categories.map(alignedAmount)))}</td>${monthTotals.map((value) => (value > release + 0.004 ? `<td class="number loss">▲ ${amount(value)}</td>` : `<td class="number">${cell(value)}</td>`)).join("")}<td class="number">${balanceCell(cents(total(balances)))}</td></tr>
      <tr class="release-row"><td>Monthly release (NEP ÷ 12)<small>▲ month total above the release</small></td><td></td>${months.map(() => `<td class="number">${amount(release)}</td>`).join("")}<td></td></tr></tfoot>
    </table></div></section>`;
}

// One month's SOB in the same columns as the JAN–DEC sheets. Every MOOE category gets its balance
// for the month, its items, an entry row for new items, and a subtotal with the running balance.
function monthBreakdown(month) {
  const index = months.indexOf(month);
  const nep = number(budget.profile.nep);
  const release = cents(nep / 12);
  const items = allItems();
  const spentBefore = (category) => total(items.filter((item) => (!category || item.categoryId === category.id) && months.indexOf(item.month) < index).map(itemCost));
  const grand = monthSpent(month);
  const tabs = months.map((name) => `<button type="button" class="month-tab${name === month ? " active" : ""}" data-month="${name}" aria-pressed="${name === month}">${name.slice(0, 3)}<small>${monthSpent(name) ? wholeMoney.format(monthSpent(name)) : "—"}</small></button>`).join("");
  let runningTotal = 0;
  const forms = [];
  const exhausted = [];
  const groups = budget.categories.map((category) => {
    const balance = cents(alignedAmount(category) - spentBefore(category));
    const rows = items.filter((item) => item.month === month && item.categoryId === category.id);
    const subtotal = total(rows.map(itemCost));
    const left = cents(balance - subtotal);
    runningTotal += left;
    // Used up in earlier months and nothing planned this month: leave it out of this and later months.
    if (balance <= 0.004 && spentBefore(category) > 0 && !rows.length) { exhausted.push(category.name || "Unnamed category"); return ""; }
    const formId = `add-${category.id}`;
    const [, presetUnit, suggestions] = categoryHint(category);
    // Auto-billed expenses only open an entry row on demand (+); electricity is the bill alone, no extra items.
    const autoBill = UTILITY_BILLS.some(([pattern]) => pattern.test(category.name || ""));
    const billOnly = /electric/i.test(category.name || "");
    // Office Supplies: the regular monthly items wait as ready rows until added for this month.
    const readyRows = /office supplies/i.test(category.name || "") ? OFFICE_DEFAULTS.filter(([name]) => !rows.some((item) => item.description.trim().toLowerCase() === name.toLowerCase())) : [];
    const entryRow = (formId, preset, hidden, stockNo) => {
      forms.push([formId, category.id]);
      return `<tr class="entry-row${preset ? " preset-row" : ""}" data-entry="${formId}"${hidden ? " hidden" : ""}><td><input class="cell-input" form="${formId}" name="aipNo" aria-label="AIP item number"></td><td><input class="cell-input" form="${formId}" name="page" aria-label="Page"></td><td></td><td></td><td class="number muted">${stockNo}</td>
        <td><input class="cell-input number" form="${formId}" name="quantity" required type="number" ${JOB_ORDER.test(category.name || "") ? `min="1" step="1" placeholder="Pax" aria-label="Number of workers"` : `min="0.01" step="0.01" placeholder="Qty" aria-label="Quantity"`}></td>
        <td><select class="cell-input" form="${formId}" name="unit" required aria-label="Unit of issue">${unitOptions(preset?.unit || presetUnit)}</select></td>
        <td>${itemChoices(category).length && !preset
          ? `<select class="cell-input" form="${formId}" name="description" required aria-label="Item description"><option value="" selected disabled>Choose item</option>${itemChoices(category).map(([name, unit, cost]) => `<option value="${escapeHtml(name)}" data-unit="${unit}" data-cost="${cost}">${escapeHtml(name)}</option>`).join("")}<option value="__other">Other (type it)…</option></select>`
          : `<input class="cell-input" form="${formId}" name="description" required ${preset ? `value="${escapeHtml(preset.description)}"` : ""} list="items-${category.id}" placeholder="${escapeHtml(suggestions[0] ? `e.g. ${suggestions[0]}` : "Item description")}" aria-label="Item description"><datalist id="items-${category.id}">${suggestions.map((item) => `<option value="${escapeHtml(item)}">`).join("")}</datalist>`}</td>
        <td>${JOB_ORDER.test(category.name || "")
          ? `<select class="cell-input number" form="${formId}" name="unitCost" required aria-label="Monthly pay per worker"><option value="" selected disabled>Pay</option><optgroup label="Per month">${JOB_ORDER_PAY.map((pay) => `<option value="${pay}">${wholeMoney.format(pay)}</option>`).join("")}</optgroup><optgroup label="Per day × ${WORK_DAYS} days">${JOB_ORDER_DAILY.map((rate) => `<option value="d${rate}">${wholeMoney.format(rate)}/day</option>`).join("")}</optgroup></select>`
          : `<input class="cell-input number" form="${formId}" name="unitCost" required type="text" inputmode="decimal" data-money placeholder="0.00" ${preset?.unitCost ? `value="${amount(preset.unitCost)}"` : ""} aria-label="Estimated unit cost">`}</td>
        <td class="number muted" data-preview="${formId}">—</td><td></td>
        <td><button class="icon-button add" form="${formId}" type="submit" aria-label="Add item to ${escapeHtml(category.name || "category")}">+</button></td></tr>`;
    };
    return `<tbody class="sob-group">
      <tr class="group-row"><td colspan="3"></td><td class="number">${amount(balance)}</td><th colspan="8" scope="rowgroup">${escapeHtml(category.name || "Unnamed category")}</th></tr>
      ${rows.map((item, row) => `<tr${item.auto ? ` class="auto-row"` : ""}><td>${escapeHtml(item.aipNo || "")}</td><td>${escapeHtml(item.page || "")}</td><td></td><td></td><td class="number">${row + 1}</td><td class="number">${number(item.quantity)}</td><td>${escapeHtml(item.unit)}</td><td><span class="category-name">${escapeHtml(item.description)}</span>${item.repeat ? ` <span class="auto-tag" title="Set in January; change or remove it there">From January</span>` : item.quarterly ? ` <span class="auto-tag" title="Planned automatically each quarter">Auto · quarterly</span>` : item.auto ? ` <span class="auto-tag" title="Planned automatically: aligned amount ÷ 12">Auto · monthly</span>` : ""}</td><td class="number">${amount(item.unitCost)}${payNote(item, category)}</td><td class="number">${amount(itemCost(item))}</td><td></td><td>${item.auto ? "" : `<button class="icon-button" data-remove-item="${item.id}" aria-label="Remove ${escapeHtml(item.description)}">×</button>`}</td></tr>`).join("")}
      ${billOnly ? "" : [
        ...readyRows.map(([description, unit, unitCost], at) => entryRow(`${formId}-d${at}`, { description, unit, unitCost }, false, rows.length + at + 1)),
        entryRow(formId, null, autoBill, rows.length + readyRows.length + 1),
      ].join("")}
      <tr class="subtotal-row"><td colspan="9" class="number">Subtotal</td><td class="number">${amount(subtotal)}</td><td class="number${left < 0 ? " loss" : ""}">${left < 0 ? "−" : ""}${amount(Math.abs(left))}</td><td>${autoBill && !billOnly ? `<button type="button" class="icon-button add" data-reveal="${formId}" aria-label="Add another ${escapeHtml(category.name)} item">+</button>` : ""}</td></tr>
    </tbody>`;
  }).join("");
  const gap = cents(release - grand);
  return `<section class="panel app-panel"><div class="panel-heading"><div><h2>School Operating Budget · ${month} ${number(budget.profile.fiscalYear) || ""}</h2><p>Type items straight into each MOOE category and press Enter or + to add them. Water, electricity, internet and mobile are planned automatically: their aligned amount ÷ 12 each month.</p></div><span class="table-note">${gap > 0.004 ? `${money(gap)} of ${money(release)} left this month` : gap < -0.004 ? `${money(-gap)} over this month` : "✓ Month fully planned"}</span></div>
    <div class="month-tabs" role="group" aria-label="Month">${tabs}</div>
    <div class="staff-row">${(() => {
      const staff = staffFor(month);
      return `<span class="staff-title">Personnel in ${month}</span>
        <label>Teaching<input class="cell-input number" type="number" min="0" step="1" data-staff="teaching" value="${staff.teaching}"></label>
        <label>Non-teaching<input class="cell-input number" type="number" min="0" step="1" data-staff="nonTeaching" value="${staff.nonTeaching}"></label>
        <span class="staff-total">Total <b>${staff.teaching + staff.nonTeaching}</b></span>
        ${staff.from && staff.from !== month ? `<span class="muted">Carried over from ${staff.from}. Change it if staffing changed this month.</span>` : ""}`;
    })()}</div>
    <p class="form-hint" id="month-room" aria-live="polite"></p>
    <div class="table-scroll"><table class="data-table sob-table">
      <thead><tr><th>AIP ITEM NO.</th><th>PAGE</th><th class="number">ANNUAL BUDGET</th><th class="number">BALANCE FOR THE MONTH</th><th class="number">STOCK NO.</th><th class="number">QTY.</th><th>UNIT OF ISSUE</th><th>ITEM DESCRIPTION</th><th class="number">EST. UNIT COST</th><th class="number">ESTIMATED COST</th><th class="number">RUNNING BALANCE</th><th></th></tr></thead>
      <tbody><tr class="budget-row"><td colspan="2"></td><td class="number">${amount(nep)}</td><td class="number">${amount(cents(nep - spentBefore()))}</td><td colspan="8"></td></tr></tbody>
      ${groups}
      ${exhausted.length ? `<tbody><tr class="exhausted-row"><td colspan="12">✓ Used up in earlier months, not shown: ${exhausted.map(escapeHtml).join(", ")}</td></tr></tbody>` : ""}
      <tfoot><tr><td colspan="9" class="number">GRAND TOTAL<small>Monthly allocation (NEP ÷ 12) ${money(release)}</small></td><td class="number">${amount(grand)}</td><td class="number">${amount(cents(runningTotal))}</td><td></td></tr></tfoot>
    </table></div>
    ${forms.map(([formId, categoryId]) => `<form id="${formId}" data-category-form="${categoryId}" hidden></form>`).join("")}
  </section>`;
}

function sobScreen() {
  if (!budget.categories.length) {
    document.querySelector("#content").innerHTML = `<div class="empty-workspace"><div class="empty-icon">1</div><h1>Add MOOE categories first</h1><p>Set up the school fund and expense categories before adding planned purchases to the monthly SOB.</p><button class="button primary" id="go-setup">Set up the budget</button></div>`;
    document.querySelector("#go-setup").addEventListener("click", () => { screen = "setup"; render(); }); return;
  }
  document.querySelector("#content").innerHTML = `
    <div class="page-heading"><div><span class="step-tag">STEP 02 <i>•</i> SCHOOL OPERATING BUDGET</span><h1>Plan each month’s expenses</h1><p>Each month may use up to the monthly allocation (NEP ÷ 12). The monthly entries become the quarterly APP.</p></div></div>
    <div id="month-breakdown"></div>
    ${utilizationPanel()}`;
  const breakdown = document.querySelector("#month-breakdown");
  const release = cents(number(budget.profile.nep) / 12);
  const fields = (form) => Object.fromEntries([...form.elements].map((field) => [field.name, field]));
  // Monthly SOB is capped at the monthly release (NEP ÷ 12); returns false when the item would exceed it.
  const checkRoom = (form) => {
    const hint = document.querySelector("#month-room");
    const category = form && budget.categories.find((entry) => entry.id === form.dataset.categoryForm);
    // A January item in a repeating service lands in every month, so it must fit the tightest month.
    const repeats = formMonth === months[0] && REPEAT_MONTHLY.test(category?.name || "");
    const tightest = repeats ? months.reduce((low, month) => (release - monthSpent(month) < release - monthSpent(low) ? month : low), months[0]) : formMonth;
    const room = cents(release - monthSpent(tightest));
    const cost = form ? (({ quantity, unitCost }) => cents(number(quantity.value) * payValue(unitCost.value).unitCost))(fields(form)) : 0;
    const over = cost - room > 0.004;
    // Semi-expendable means below ₱50,000 per unit; dearer items are capital outlay, not MOOE.
    const tooDear = Boolean(form && semiItems(category) && payValue(fields(form).unitCost.value).unitCost >= SEMI_EXPENDABLE_LIMIT);
    if (form) document.querySelector(`[data-preview="${form.id}"]`).textContent = cost ? amount(cost) : "—";
    hint.className = `form-hint${over || tooDear || !release ? " error" : ""}`;
    hint.textContent = !release ? "Enter the Annual NEP allocation in Budget setup first. The monthly SOB limit is NEP ÷ 12."
      : tooDear ? `Semi-expendable items must cost below ${money(SEMI_EXPENDABLE_LIMIT)} each. An item at ${money(SEMI_EXPENDABLE_LIMIT)} or more is equipment (capital outlay), not MOOE.`
      : over ? `This item costs ${money(cost)}${repeats ? " every month" : ""}, which is ${money(cost - room)} more than ${tightest} has left (${money(room)}). Lower the cost${repeats ? "" : " or plan it in another month"}.`
      : `${repeats ? `Repeats every month · tightest month ${tightest}` : formMonth}: ${money(room)} of the ${money(release)} monthly allocation still available${cost ? ` · ${money(room - cost)} after this item` : ""}.`;
    return release > 0 && !over && !tooDear;
  };
  const showMonth = () => {
    breakdown.innerHTML = monthBreakdown(formMonth);
    checkRoom();
    breakdown.querySelectorAll("[data-staff]").forEach((input) => input.addEventListener("change", () => {
      const current = staffFor(formMonth);
      (budget.staff ||= {})[formMonth] = { teaching: current.teaching, nonTeaching: current.nonTeaching, [input.dataset.staff]: Math.max(0, Math.round(number(input.value))) };
      saveBudget(); showMonth();
    }));
    breakdown.querySelectorAll("[data-month]").forEach((tab) => tab.addEventListener("click", () => { formMonth = tab.dataset.month; showMonth(); }));
    breakdown.querySelectorAll("[data-reveal]").forEach((button) => button.addEventListener("click", () => {
      button.hidden = true;
      breakdown.querySelector(`[data-entry="${button.dataset.reveal}"]`).hidden = false;
      breakdown.querySelector(`[form="${button.dataset.reveal}"][name="quantity"]`).focus();
    }));
    breakdown.querySelectorAll("[data-remove-item]").forEach((button) => button.addEventListener("click", () => { budget.items = budget.items.filter((item) => item.id !== button.dataset.removeItem); saveBudget(); render(); }));
    breakdown.querySelector(".sob-table").addEventListener("input", (event) => { if (event.target.form) checkRoom(event.target.form); });
    breakdown.querySelectorAll("select[name=description]").forEach((select) => select.addEventListener("change", () => {
      const { form } = select;
      if (select.value === "__other") {
        const input = Object.assign(document.createElement("input"), { className: "cell-input", name: "description", required: true, placeholder: "Type the item" });
        input.setAttribute("form", form.id);
        select.replaceWith(input);
        input.focus();
        return;
      }
      const { unit, cost } = select.selectedOptions[0].dataset;
      const { unit: unitField, unitCost } = fields(form);
      if (unit) unitField.value = unit;
      if (Number(cost)) unitCost.value = amount(cost);
      checkRoom(form);
    }));
    breakdown.querySelectorAll("[data-category-form]").forEach((form) => form.addEventListener("submit", (event) => {
      event.preventDefault();
      if (!checkRoom(form)) { fields(form).unitCost.focus(); return; }
      const { aipNo, page, quantity, unit, description, unitCost } = fields(form);
      budget.items.push({ id: newId(), month: formMonth, categoryId: form.dataset.categoryForm, aipNo: aipNo.value.trim(), page: page.value.trim(), description: description.value.trim(), quantity: number(quantity.value), unit: unit.value, ...payValue(unitCost.value) });
      focusCategory = form.dataset.categoryForm;
      saveBudget(); render();
    }));
    if (focusCategory) document.querySelector(`[form^="add-${focusCategory}"][name="quantity"]`)?.focus();
    focusCategory = null;
  };
  showMonth();
}


function consolidatedItems() {
  const items = new Map();
  allItems().forEach((item) => {
    const category = budget.categories.find((entry) => entry.id === item.categoryId)?.name || "Uncategorized";
    const key = `${category}\u0000${item.description}\u0000${item.unit}`;
    if (!items.has(key)) items.set(key, { category, description: item.description, unit: item.unit, quantities: [0, 0, 0, 0], amounts: [0, 0, 0, 0] });
    const entry = items.get(key); const quarter = quarterForMonth(item.month);
    entry.quantities[quarter] += number(item.quantity); entry.amounts[quarter] += itemCost(item); // same per-item rounding as the SOB, so totals match
  });
  return [...items.values()].sort((a, b) => a.category.localeCompare(b.category) || a.description.localeCompare(b.description));
}

function appScreen() {
  const items = consolidatedItems();
  const quarterTotals = [0, 1, 2, 3].map((quarter) => total(items.map((item) => item.amounts[quarter])));
  const rows = items.map((item) => `<tr><td class="muted">${escapeHtml(item.category)}</td><td><span class="category-name">${escapeHtml(item.description)}</span></td><td>${escapeHtml(item.unit)}</td>${[0, 1, 2, 3].map((quarter) => `<td class="number">${item.quantities[quarter] || "—"}</td><td class="number">${money(item.amounts[quarter])}</td>`).join("")}<td class="number strong">${money(total(item.amounts))}</td></tr>`).join("");
  document.querySelector("#content").innerHTML = `
    <div class="page-heading"><div><span class="step-tag">STEP 03 <i>•</i> ANNUAL PROCUREMENT PLAN</span><h1>Review the quarterly APP</h1><p>Fund Planner groups your monthly SOB entries into four procurement quarters.</p></div><button class="button primary" id="export-app" ${items.length ? "" : "disabled"}>Export quarterly APP</button></div>
    <section class="quarter-grid">${["Q1", "Q2", "Q3", "Q4"].map((quarter, index) => `<button type="button" class="quarter-card${appQuarter === index ? " active" : ""}" data-quarter="${index}" aria-pressed="${appQuarter === index}"><div class="quarter-label">${quarter} <span>${["JAN — MAR", "APR — JUN", "JUL — SEP", "OCT — DEC"][index]}</span></div><div class="quarter-amount">${wholeMoney.format(quarterTotals[index])}</div><div class="quarter-foot">${items.filter((item) => item.amounts[index] > 0).length} planned items · ${appQuarter === index ? "Hide breakdown" : "Show breakdown"}</div></button>`).join("")}</section>
    ${appQuarter == null ? "" : quarterBreakdown(appQuarter)}
    <section class="panel app-panel"><div class="panel-heading"><div><h2>Quarterly procurement plan</h2><p>Generated from the monthly SOB entries you entered.</p></div><span class="table-note">${items.length} consolidated items</span></div><div class="table-scroll"><table class="data-table app-table"><thead><tr><th>MOOE CATEGORY</th><th>ITEM</th><th>UNIT</th>${["Q1", "Q2", "Q3", "Q4"].map((quarter) => `<th class="number">${quarter} QTY</th><th class="number">${quarter} AMOUNT</th>`).join("")}<th class="number">YEAR TOTAL</th></tr></thead><tbody>${rows || `<tr><td colspan="12" class="empty-state">Add planned SOB items to create the quarterly APP.</td></tr>`}</tbody></table></div><div class="table-total"><span>Annual planned cost</span><strong>${money(total(quarterTotals))}</strong><span class="balance-note">${items.length} items</span></div></section>`;
  document.querySelector("#export-app")?.addEventListener("click", exportApp);
  document.querySelectorAll("[data-quarter]").forEach((card) => card.addEventListener("click", () => {
    const quarter = Number(card.dataset.quarter);
    appQuarter = appQuarter === quarter ? null : quarter;
    render();
    document.querySelector(".quarter-breakdown")?.scrollIntoView({ behavior: "smooth", block: "start" });
  }));
}

// One quarter of the APP: items grouped by MOOE category with each month's amount, the quarter's qty and amount.
function quarterBreakdown(quarter) {
  const quarterMonths = months.slice(quarter * 3, quarter * 3 + 3);
  const items = allItems().filter((item) => quarterMonths.includes(item.month));
  // Same item at a different price stays its own line, like the APP workbook (qty × unit cost = amount).
  const groups = budget.categories.map((category) => {
    const lines = new Map();
    items.filter((item) => item.categoryId === category.id).forEach((item) => {
      const key = `${item.description}\u0000${item.unit}\u0000${number(item.unitCost)}`;
      if (!lines.has(key)) lines.set(key, { description: item.description, unit: item.unit, unitCost: number(item.unitCost), dailyRate: item.dailyRate, qtyByMonth: [0, 0, 0], amount: 0 });
      const line = lines.get(key);
      line.qtyByMonth[quarterMonths.indexOf(item.month)] += number(item.quantity);
      line.amount += itemCost(item);
    });
    return { category, lines: [...lines.values()].sort((a, b) => a.description.localeCompare(b.description)) };
  }).filter((group) => group.lines.length);
  const qty = (value) => (value ? value.toLocaleString("en-PH", { maximumFractionDigits: 2 }) : `<span class="muted">—</span>`);
  const quarterTotal = total(items.map(itemCost));
  let itemNo = 0;
  const body = groups.map(({ category, lines }) => `<tbody class="sob-group"><tr class="group-row"><th colspan="9" scope="rowgroup">${escapeHtml(category.name || "Unnamed category")}</th></tr>
      ${lines.map((line) => `<tr><td class="number muted">${++itemNo}</td><td><span class="category-name">${escapeHtml(line.description)}</span></td><td>${escapeHtml(line.unit)}</td><td class="number">${amount(line.unitCost)}${payNote(line, category)}</td>${line.qtyByMonth.map((value) => `<td class="number">${qty(value)}</td>`).join("")}<td class="number strong">${qty(total(line.qtyByMonth))}</td><td class="number strong">${amount(line.amount)}</td></tr>`).join("")}
      <tr class="subtotal-row"><td colspan="8" class="number">Sub-total</td><td class="number">${amount(total(lines.map((line) => line.amount)))}</td></tr></tbody>`).join("");
  return `<section class="panel app-panel quarter-breakdown"><div class="panel-heading"><div><h2>${["1st", "2nd", "3rd", "4th"][quarter]} Quarter · ${quarterMonths[0]} to ${quarterMonths[2]}</h2><p>Every item to buy or pay for this quarter: how many, at what unit cost, and in which month.</p></div><span class="table-note">${itemNo} items · ${money(quarterTotal)}</span></div>
    <div class="table-scroll"><table class="data-table quarter-table"><thead><tr><th class="number">ITEM NO.</th><th>ITEM DESCRIPTION</th><th>UNIT</th><th class="number">UNIT COST</th>${quarterMonths.map((month) => `<th class="number">${month.slice(0, 3).toUpperCase()} QTY</th>`).join("")}<th class="number">QUARTER QTY</th><th class="number">AMOUNT</th></tr></thead>
      ${body || `<tbody><tr><td colspan="9" class="empty-state">Nothing planned for this quarter yet.</td></tr></tbody>`}
      <tfoot><tr><td colspan="8" class="number">Quarter total</td><td class="number">${amount(quarterTotal)}</td></tr></tfoot></table></div></section>`;
}

async function exportAllocation() {
  try {
    const output = await window.sob.saveAllocation({ source: "Fund Planner", school: budget.profile.school, schoolId: budget.profile.schoolId, fiscalYear: budget.profile.fiscalYear, annual: number(budget.profile.nep), aligned: budget.categories.map((category) => ({ name: category.name || "Unnamed category", rate: number(category.rate) / 100, amount: alignedAmount(category) })) });
    if (output) setStatus(`Allocation exported · ${output.split(/[\\/]/).pop()}`, "success");
  } catch (error) { setStatus(error.message || "Could not export allocation", "error"); }
}

async function exportApp() {
  try {
    const output = await window.sob.exportApp({ source: "Fund Planner", school: budget.profile.school || "School", schoolId: budget.profile.schoolId, fiscalYear: budget.profile.fiscalYear, items: consolidatedItems() });
    if (output) setStatus(`APP exported · ${output.split(/[\\/]/).pop()}`, "success");
  } catch (error) { setStatus(error.message || "Could not export APP", "error"); }
}

let shownScreen = null;
function render() {
  updateContext();
  if (screen === "setup") setupScreen(); else if (screen === "sob") sobScreen(); else if (screen === "dash") dashboardScreen(); else appScreen();
  // Animate only on a tab change; render() also runs on every edit.
  if (shownScreen !== null && shownScreen !== screen && !matchMedia("(prefers-reduced-motion: reduce)").matches) {
    document.querySelector("#content").animate([{ opacity: 0, transform: "translateY(10px)" }, { opacity: 1, transform: "none" }], { duration: 260, easing: "cubic-bezier(.2, .7, .2, 1)" });
  }
  shownScreen = screen;
}

// Fund dashboard: the BREAKDOWN sheet as pictures. Where money moved (default → aligned),
// how each month spends against NEP ÷ 12, how the year's balance runs down to ₱0, and utilization per expense.
function dashboardScreen() {
  const nep = number(budget.profile.nep);
  const release = cents(nep / 12);
  const items = allItems();
  const aligned = total(budget.categories.map(alignedAmount));
  const perMonth = months.map((month) => monthSpent(month));
  const planned = total(perMonth);
  const moves = budget.moves || [];
  const short = (value) => wholeMoney.format(value);
  // 5% tax on each monthly allocation, shown only (the BREAKDOWN sheet "with less tax" column); nothing is deducted from the plan.
  const taxMonth = cents(release * TAX_RATE);
  const taxYear = cents(taxMonth * 12);

  const tiles = [
    ["ANNUAL NEP", short(nep), "Released annual allocation"],
    ["ALIGNED", short(aligned), cents(nep - aligned) ? `${money(Math.abs(cents(nep - aligned)))} ${nep - aligned > 0 ? "not yet assigned" : "over the NEP"}` : "Matches the NEP"],
    ["PLANNED JAN–DEC", short(planned), nep ? `${(planned / nep * 100).toFixed(1)}% of the NEP` : "—"],
    ["NOT YET PLANNED", short(Math.max(0, cents(aligned - planned))), cents(aligned - planned) > 0 ? "Should reach ₱0 by December" : "✓ Every aligned peso is planned"],
    ["REALIGNED", short(total(moves.map((move) => move.amount))), `${moves.length} move${moves.length === 1 ? "" : "s"} recorded`],
  ].map(([label, value, foot]) => `<article class="stat-card"><div class="card-kicker">${label}</div><div class="stat-value">${value}</div><div class="card-foot">${foot}</div></article>`).join("");

  // Realignment per expense: diverging bars around zero, + to the right, − to the left, always with a signed label.
  const shifts = budget.categories.map((category) => ({ name: category.name || "Unnamed category", delta: cents(alignedAmount(category) - defaultAmount(category)) })).filter((row) => row.delta);
  const maxShift = Math.max(1, ...shifts.map((row) => Math.abs(row.delta)));
  const shiftRows = shifts.map(({ name, delta }) => `<div class="shift-row" title="${escapeHtml(name)}: ${delta > 0 ? "+" : "−"}${money(Math.abs(delta))}">
      <span class="shift-name">${escapeHtml(name)}</span>
      <span class="shift-track"><span class="shift-bar ${delta > 0 ? "plus" : "minus"}" style="width:${(Math.abs(delta) / maxShift * 50).toFixed(2)}%"></span></span>
      <span class="shift-value ${delta > 0 ? "gain" : "loss"}">${delta > 0 ? "+" : "−"}${money(Math.abs(delta))}</span></div>`).join("");
  const moveRows = [...moves].reverse().map((move) => `<tr><td class="muted">${new Date(move.date).toLocaleDateString("en-PH", { month: "short", day: "numeric" })}</td><td class="move-cell"><span class="loss">${escapeHtml(move.from)}</span> → <span class="gain">${escapeHtml(move.to)}</span></td><td class="number strong">${money(move.amount)}</td></tr>`).join("");

  // Spending by expense per month: stacked bars, the 6 biggest expenses get fixed colors, the rest fold into "Other".
  const spendOf = (category, month) => total(items.filter((item) => item.categoryId === category.id && (!month || item.month === month)).map(itemCost));
  const spenders = budget.categories.map((category) => ({ category, year: spendOf(category) })).filter((row) => row.year > 0).sort((a, b) => b.year - a.year);
  const series = [...spenders.slice(0, 6).map(({ category }, index) => ({ name: category.name || "Unnamed category", id: category.id, color: SERIES_COLORS[index], ids: [category.id] })),
    ...(spenders.length > 6 ? [{ name: "Other expenses", id: null, color: "#a3a8b8", ids: spenders.slice(6).map(({ category }) => category.id) }] : [])];
  // Default to the biggest expense the school plans item by item (not an automatic monthly bill).
  const isBill = (category) => UTILITY_BILLS.some(([pattern]) => pattern.test(category.name || ""));
  const detail = budget.categories.find((category) => category.id === dashExpense) || spenders.find(({ category }) => !isBill(category))?.category || spenders[0]?.category || budget.categories[0];
  const W = 1000, H = 280, left = 78, bottom = 32, top = 28, right = 12;
  const yMax = Math.max(release, ...perMonth) * 1.12 || 1;
  const y = (value) => top + (H - top - bottom) * (1 - value / yMax);
  const slot = (W - left - right) / 12;
  const barW = slot * 0.58;
  const stacked = months.map((month, index) => {
    const x = left + index * slot + (slot - barW) / 2;
    let base = 0;
    const segments = series.map((serie) => {
      const value = total(items.filter((item) => serie.ids.includes(item.categoryId) && item.month === month).map(itemCost));
      if (!value) return "";
      const y0 = y(base), y1 = y(base + value);
      base += value;
      // 2px surface gap between stacked segments.
      return `<rect class="seg" style="--i:${index}" x="${x.toFixed(1)}" y="${y1.toFixed(1)}" width="${barW.toFixed(1)}" height="${Math.max(1, y0 - y1 - 2).toFixed(1)}" rx="2" fill="${serie.color}"${serie.id ? ` data-expense="${serie.id}"` : ""}><title>${month} · ${escapeHtml(serie.name)}: ${money(value)}</title></rect>`;
    }).join("");
    const monthTotal = perMonth[index];
    return `${segments}${monthTotal ? `<text class="bar-total${monthTotal > release + 0.004 ? " over" : ""}" x="${(x + barW / 2).toFixed(1)}" y="${(y(monthTotal) - 6).toFixed(1)}" text-anchor="middle">${monthTotal > release + 0.004 ? "▲ " : ""}${short(monthTotal).replace("₱", "")}</text>` : ""}
      <text class="axis" x="${(x + barW / 2).toFixed(1)}" y="${H - 10}" text-anchor="middle">${month.slice(0, 3)}</text>`;
  }).join("");
  const grid = [0, 0.5, 1].map((share) => yMax / 1.12 * share).map((value) => `<line class="grid" x1="${left}" x2="${W - right}" y1="${y(value).toFixed(1)}" y2="${y(value).toFixed(1)}"></line><text class="axis" x="${left - 8}" y="${(y(value) + 4).toFixed(1)}" text-anchor="end">${short(value)}</text>`).join("");
  const stackChart = `<svg viewBox="0 0 ${W} ${H}" class="chart" role="img" aria-label="Spending per month, split by MOOE expense">${grid}${stacked}
    <line class="ref" x1="${left}" x2="${W - right}" y1="${y(release).toFixed(1)}" y2="${y(release).toFixed(1)}"></line>
    <text class="ref-label" x="${W - right}" y="${(y(release) - 6).toFixed(1)}" text-anchor="end">Monthly allocation ${short(release)}</text></svg>`;
  const stackLegend = series.map((serie) => `<button type="button" class="legend-item${serie.id && serie.id === detail?.id ? " active" : ""}"${serie.id ? ` data-expense="${serie.id}"` : " disabled"}><i class="key" style="background:${serie.color}"></i>${escapeHtml(serie.name)}</button>`).join("");

  // Expense detail: one expense's spending per month and how many of each item it buys each month.
  const detailItems = detail ? items.filter((item) => item.categoryId === detail.id) : [];
  const detailMonthly = months.map((month) => total(detailItems.filter((item) => item.month === month).map(itemCost)));
  const detailAligned = detail ? alignedAmount(detail) : 0;
  const detailUsed = total(detailMonthly);
  const share = cents(detailAligned / 12);
  const dMax = Math.max(share, ...detailMonthly) * 1.15 || 1;
  const dy = (value) => top + (H - top - bottom) * (1 - value / dMax);
  const detailColor = series.find((serie) => serie.id === detail?.id)?.color || "#4f46e5";
  const detailBars = detailMonthly.map((value, index) => {
    const x = left + index * slot + (slot - barW) / 2;
    return `${value ? `<rect class="seg" style="--i:${index}" x="${x.toFixed(1)}" y="${dy(value).toFixed(1)}" width="${barW.toFixed(1)}" height="${Math.max(1, H - bottom - dy(value)).toFixed(1)}" rx="4" fill="${detailColor}"><title>${months[index]}: ${money(value)}</title></rect><text class="bar-total" x="${(x + barW / 2).toFixed(1)}" y="${(dy(value) - 6).toFixed(1)}" text-anchor="middle">${short(value).replace("₱", "")}</text>` : ""}
      <text class="axis" x="${(x + barW / 2).toFixed(1)}" y="${H - 10}" text-anchor="middle">${months[index].slice(0, 3)}</text>`;
  }).join("");
  const detailChart = `<svg viewBox="0 0 ${W} ${H}" class="chart" role="img" aria-label="Monthly spending for ${escapeHtml(detail?.name || "the selected expense")}">
    ${[0, 0.5, 1].map((part) => dMax / 1.15 * part).map((value) => `<line class="grid" x1="${left}" x2="${W - right}" y1="${dy(value).toFixed(1)}" y2="${dy(value).toFixed(1)}"></line><text class="axis" x="${left - 8}" y="${(dy(value) + 4).toFixed(1)}" text-anchor="end">${short(value)}</text>`).join("")}
    ${detailBars}
    ${share ? `<line class="ref" x1="${left}" x2="${W - right}" y1="${dy(share).toFixed(1)}" y2="${dy(share).toFixed(1)}"></line><text class="ref-label" x="${W - right}" y="${(dy(share) - 6).toFixed(1)}" text-anchor="end">Even monthly share ${short(share)}</text>` : ""}</svg>`;
  // Item × month quantities, e.g. Bond Paper A4: 40 Jan · 20 Feb · 30 Mar.
  const itemLines = new Map();
  detailItems.forEach((item) => {
    const key = `${item.description}\u0000${item.unit}`;
    if (!itemLines.has(key)) itemLines.set(key, { description: item.description, unit: item.unit, qty: months.map(() => 0), amount: 0 });
    const line = itemLines.get(key);
    line.qty[months.indexOf(item.month)] += number(item.quantity);
    line.amount += itemCost(item);
  });
  const qtyCell = (value) => (value ? value.toLocaleString("en-PH", { maximumFractionDigits: 2 }) : `<span class="muted">—</span>`);
  const itemRows = [...itemLines.values()].sort((a, b) => b.amount - a.amount).map((line) => `<tr><td><span class="category-name">${escapeHtml(line.description)}</span></td><td>${escapeHtml(line.unit)}</td>${line.qty.map((value) => `<td class="number">${qtyCell(value)}</td>`).join("")}<td class="number strong">${qtyCell(total(line.qty))}</td><td class="number strong">${money(line.amount)}</td></tr>`).join("");
  const topItem = [...itemLines.values()].sort((a, b) => b.amount - a.amount)[0];

  // Utilization per expense: planned against the aligned amount.
  const useRows = budget.categories.map((category) => {
    const target = alignedAmount(category);
    const used = total(items.filter((item) => item.categoryId === category.id).map(itemCost));
    const share = target ? Math.min(1, used / target) : used ? 1 : 0;
    const rest = cents(target - used);
    return `<tr><td><span class="category-name">${escapeHtml(category.name || "Unnamed category")}</span></td>
      <td class="meter-cell"><span class="meter"><span class="meter-fill${rest < 0 ? " over" : ""}" style="width:${(share * 100).toFixed(1)}%"></span></span></td>
      <td class="number">${target ? `${Math.round(used / target * 100)}%` : "—"}</td><td class="number">${money(used)}</td><td class="number">${money(target)}</td>
      <td class="number">${rest > 0 ? `<span class="warn">${money(rest)} left</span>` : rest < 0 ? `<span class="loss">${money(-rest)} over</span>` : `<span class="gain">✓ Used up</span>`}</td></tr>`;
  }).join("");

  document.querySelector("#content").innerHTML = `
    <div class="page-heading"><div><span class="step-tag">OVERVIEW <i>•</i> FUND DASHBOARD</span><h1>Movement of the school fund</h1><p>From the default breakdown, through realignment, to monthly spending and the balance left at year end.</p></div></div>
    <section class="overview-grid five">${tiles}</section>
    <section class="panel tax-panel"><div class="panel-heading"><div><h2>5% tax withheld from each monthly allocation</h2><p>For reference only. The plan, the SOB and the monthly limit are not reduced.</p></div><span class="table-note">${money(taxYear)} for the year</span></div>
      <div class="table-scroll"><table class="data-table tax-table"><thead><tr><th></th><th class="number">PER MONTH</th><th class="number">FOR THE YEAR (× 12)</th></tr></thead>
        <tbody><tr><td>Monthly allocation (NEP ÷ 12)</td><td class="number">${amount(release)}</td><td class="number">${amount(cents(release * 12))}</td></tr>
          <tr class="tax-row"><td>Less 5% tax</td><td class="number">−${amount(taxMonth)}</td><td class="number">−${amount(taxYear)}</td></tr>
          <tr class="net-row"><td>Net after tax</td><td class="number">${amount(cents(release - taxMonth))}</td><td class="number">${amount(cents(release * 12 - taxYear))}</td></tr></tbody></table></div></section>
    <section class="panel"><div class="panel-heading"><div><h2>Spending by expense, per month</h2><p>What each month&rsquo;s money goes to. Click an expense below or a bar segment to see its details. ▲ marks a month above the monthly allocation.</p></div><span class="table-note">${money(planned)} planned</span></div>
      <div class="chart-wrap">${spenders.length ? stackChart : `<p class="empty-state">No spending planned yet. Add items in Monthly SOB.</p>`}</div>
      <div class="legend">${stackLegend}</div></section>
    <section class="panel expense-detail"><div class="panel-heading"><div><h2>${escapeHtml(detail?.name || "Expense detail")}</h2><p>Spending per month and how many of each item are bought each month.</p></div>
      ${budget.categories.length ? `<select class="add-select" id="dash-expense" aria-label="Choose an expense">${budget.categories.map((category) => `<option value="${category.id}" ${category.id === detail?.id ? "selected" : ""}>${escapeHtml(category.name || "Unnamed category")}</option>`).join("")}</select>` : ""}</div>
      <div class="detail-stats">
        <div><span>Aligned</span><b>${money(detailAligned)}</b></div>
        <div><span>Planned so far</span><b>${money(detailUsed)}</b><small>${detailAligned ? `${Math.round(detailUsed / detailAligned * 100)}% of aligned` : "—"}</small></div>
        <div><span>Left to plan</span><b class="${detailAligned - detailUsed < -0.004 ? "loss" : ""}">${detailAligned - detailUsed < -0.004 ? "−" : ""}${money(Math.abs(cents(detailAligned - detailUsed)))}</b></div>
        <div><span>Biggest item</span><b>${topItem ? escapeHtml(topItem.description) : "—"}</b><small>${topItem ? `${qtyCell(total(topItem.qty))} ${escapeHtml(topItem.unit)} · ${money(topItem.amount)}` : ""}</small></div>
      </div>
      <div class="chart-wrap">${detailChart}</div>
      <div class="table-scroll"><table class="data-table detail-table"><thead><tr><th>ITEM</th><th>UNIT</th>${months.map((month) => `<th class="number">${month.slice(0, 3).toUpperCase()}</th>`).join("")}<th class="number">TOTAL QTY</th><th class="number">AMOUNT</th></tr></thead>
        <tbody>${itemRows || `<tr><td colspan="16" class="empty-state">No items planned for this expense yet.</td></tr>`}</tbody></table></div></section>
    <div class="dash-grid">
      <section class="panel"><div class="panel-heading"><div><h2>Realignment from the default breakdown</h2><p>How much each expense gained (+) or gave up (−) compared with the Division Office default.</p></div></div>
        <div class="shift-list">${shiftRows || `<p class="empty-state">No realignment yet. Every expense is at its default amount.</p>`}</div></section>
      <section class="panel"><div class="panel-heading"><div><h2>Realignment log</h2><p>Each move made in Budget setup, newest first.</p></div><span class="table-note">${moves.length} moves</span></div>
        <div class="table-scroll"><table class="data-table"><thead><tr><th>DATE</th><th>FROM → TO</th><th class="number">AMOUNT</th></tr></thead><tbody>${moveRows || `<tr><td colspan="3" class="empty-state">No moves recorded yet. Moves are logged when you apply a realignment in Budget setup.</td></tr>`}</tbody></table></div></section>
    </div>
    <section class="panel"><div class="panel-heading"><div><h2>Utilization per expense</h2><p>Planned spending (all months) against each aligned amount.</p></div></div>
      <div class="table-scroll"><table class="data-table"><thead><tr><th>MOOE EXPENSE</th><th>PLANNED SO FAR</th><th class="number">USED</th><th class="number">PLANNED</th><th class="number">ALIGNED</th><th class="number">BALANCE</th></tr></thead><tbody>${useRows || `<tr><td colspan="6" class="empty-state">Add MOOE categories in Budget setup.</td></tr>`}</tbody></table></div></section>`;
  const showExpense = (id) => { dashExpense = id; render(); document.querySelector(".expense-detail")?.scrollIntoView({ behavior: "smooth", block: "start" }); };
  document.querySelector("#dash-expense")?.addEventListener("change", (event) => showExpense(event.target.value));
  document.querySelectorAll("#content [data-expense]").forEach((element) => element.addEventListener("click", () => showExpense(element.dataset.expense)));
}

shell();
render();
// First start on SQLite: copy the budget kept in browser storage into the database.
if (window.sob?.loadBudget && !window.sob.loadBudget() && localStorage.getItem(STORAGE_KEY)) saveBudget();
