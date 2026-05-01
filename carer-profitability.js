import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260427";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const dateRangeMessage = document.getElementById("dateRangeMessage");
const carerHoursReportLink = document.getElementById("carerHoursReportLink");
const payrollReportLink = document.getElementById("payrollReportLink");
const carerHoursCsvInput = document.getElementById("carerHoursCsvInput");
const payrollCsvInput = document.getElementById("payrollCsvInput");
const incomeRateInput = document.getElementById("incomeRateInput");
const contractedRateInput = document.getElementById("contractedRateInput");
const taxOnCostInput = document.getElementById("taxOnCostInput");
const pensionOnCostInput = document.getElementById("pensionOnCostInput");
const exportReportBtn = document.getElementById("exportReportBtn");
const uploadStatusMessage = document.getElementById("uploadStatusMessage");
const reportOutputPanel = document.getElementById("reportOutputPanel");
const reportSummaryMessage = document.getElementById("reportSummaryMessage");
const grandTotalGrid = document.getElementById("grandTotalGrid");
const areaSummaryBody = document.getElementById("areaSummaryBody");
const carerSummaryBody = document.getElementById("carerSummaryBody");
const selectedTotalMessage = document.getElementById("selectedTotalMessage");
const selectedTotalGrid = document.getElementById("selectedTotalGrid");
const selectAllCarersBtn = document.getElementById("selectAllCarersBtn");
const clearSelectedCarersBtn = document.getElementById("clearSelectedCarersBtn");

const CARER_HOURS_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carersHoursRpt.php";
const PAYROLL_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carerPayroll.php";
const ASSUMPTIONS_STORAGE_KEY = "thrive.carerProfitability.assumptions.v1";

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let carerHoursRows = [];
let payrollRows = [];
let latestReport = null;
let reportPeriod = null;
let selectedCarerKeys = new Set();
let hasManualSelection = false;
const tableSortState = {
  area: { key: "", direction: "desc" },
  carer: { key: "", direction: "desc" },
};

const currencyFormatter = new Intl.NumberFormat("en-GB", {
  style: "currency",
  currency: "GBP",
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
});

const decimalFormatter = new Intl.NumberFormat("en-GB", {
  maximumFractionDigits: 2,
});

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setUploadStatus(message, isError = false) {
  uploadStatusMessage.textContent = message;
  uploadStatusMessage.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "reports").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function formatDateParam(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}-${month}-${year}`;
}

function formatMonthStamp(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${year}-${month}`;
}

function getLastMonthRange() {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const end = new Date(today.getFullYear(), today.getMonth(), 0);
  return { start, end };
}

function buildCarerHoursUrl(start, end) {
  const url = new URL(CARER_HOURS_BASE_URL);
  url.searchParams.set("jobtype", "All");
  url.searchParams.set("start", formatDateParam(start));
  url.searchParams.set("finish", formatDateParam(end));
  return url.toString();
}

function buildPayrollUrl(start, end) {
  const url = new URL(PAYROLL_BASE_URL);
  url.searchParams.set("dateStart", formatDateParam(start));
  url.searchParams.set("dateFinish", formatDateParam(end));
  url.searchParams.set("carers_id", "All");
  url.searchParams.set("searchpayrollCycle", "");
  url.searchParams.set("holidayOption", "true");
  url.searchParams.set("searchJobType", "All");
  url.searchParams.set("calBill", "true");
  return url.toString();
}

function updateReportLinks() {
  reportPeriod = getLastMonthRange();
  const { start, end } = reportPeriod;
  const startText = formatDateParam(start);
  const endText = formatDateParam(end);

  if (dateRangeMessage) {
    dateRangeMessage.textContent = `Last month: ${startText} to ${endText}`;
  }
  if (carerHoursReportLink) {
    carerHoursReportLink.href = buildCarerHoursUrl(start, end);
  }
  if (payrollReportLink) {
    payrollReportLink.href = buildPayrollUrl(start, end);
  }
}

function cleanCell(value) {
  return String(value || "").replace(/\r/g, "").trim();
}

function normalizeHeader(value) {
  return cleanCell(value)
    .replace(/^\uFEFF/, "")
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, "");
}

function normalizeName(value) {
  return cleanCell(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function isCentralRole(area) {
  return normalizeName(area) === "central";
}

function getField(row, candidates) {
  const normalizedCandidates = candidates.map(normalizeHeader);
  for (const [header, value] of Object.entries(row || {})) {
    if (normalizedCandidates.includes(normalizeHeader(header))) {
      return cleanCell(value);
    }
  }
  return "";
}

function parseCsvText(text) {
  const raw = String(text || "").replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;

  for (let index = 0; index < raw.length; index += 1) {
    const char = raw[index];
    const next = raw[index + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        value += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(value);
      value = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        index += 1;
      }
      row.push(value);
      rows.push(row);
      row = [];
      value = "";
      continue;
    }

    value += char;
  }

  if (value.length || row.length) {
    row.push(value);
    rows.push(row);
  }

  const nonEmptyRows = rows.filter((candidate) => candidate.some((cell) => cleanCell(cell)));
  if (!nonEmptyRows.length) {
    return { headers: [], rows: [], errors: ["CSV file is empty."] };
  }

  const headers = nonEmptyRows[0].map((header) => cleanCell(header));
  const records = [];
  const errors = [];

  for (let rowIndex = 1; rowIndex < nonEmptyRows.length; rowIndex += 1) {
    const source = nonEmptyRows[rowIndex];
    const record = {};
    for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
      record[headers[columnIndex]] = cleanCell(source[columnIndex] || "");
    }
    if (source.length !== headers.length) {
      errors.push(`Row ${rowIndex + 1} has ${source.length} values; expected ${headers.length}.`);
    }
    records.push(record);
  }

  return { headers, rows: records, errors };
}

function readFileText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(String(reader.result || "")));
    reader.addEventListener("error", () => reject(reader.error || new Error("Could not read file.")));
    reader.readAsText(file);
  });
}

function parseHours(value) {
  const raw = cleanCell(value).replace(",", ".");
  if (!raw) {
    return 0;
  }
  const timeMatch = raw.match(/^(-?\d+):(\d{1,2})$/);
  if (timeMatch) {
    const hours = Number(timeMatch[1]);
    const minutes = Number(timeMatch[2]);
    return Number.isFinite(hours) && Number.isFinite(minutes) ? hours + minutes / 60 : 0;
  }
  const number = Number(raw.replace(/[^0-9.-]/g, ""));
  return Number.isFinite(number) ? number : 0;
}

function parseCurrency(value) {
  const raw = cleanCell(value);
  const isNegative = /^\(.*\)$/.test(raw);
  const number = Number(raw.replace(/[£,\s()]/g, ""));
  if (!Number.isFinite(number)) {
    return 0;
  }
  return isNegative ? -number : number;
}

function parseSortNumber(value) {
  const number = Number(String(value || "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(number) ? number : null;
}

function hasCarerIdentity({ id, firstName, surname, fallbackName }) {
  return Boolean(cleanCell(id) || cleanCell(firstName) || cleanCell(surname) || cleanCell(fallbackName));
}

function getAssumptions() {
  return {
    incomeRate: Math.max(0, Number(incomeRateInput?.value || 0) || 0),
    contractedRate: Math.max(0, Number(contractedRateInput?.value || 0) || 0),
    taxPercent: Math.max(0, Number(taxOnCostInput?.value || 0) || 0),
    pensionPercent: Math.max(0, Number(pensionOnCostInput?.value || 0) || 0),
  };
}

function saveAssumptions() {
  try {
    localStorage.setItem(ASSUMPTIONS_STORAGE_KEY, JSON.stringify(getAssumptions()));
  } catch (error) {
    console.warn("Could not save carer profitability assumptions.", error);
  }
}

function loadAssumptions() {
  try {
    const stored = JSON.parse(localStorage.getItem(ASSUMPTIONS_STORAGE_KEY) || "{}");
    if (Number.isFinite(Number(stored.incomeRate)) && incomeRateInput) {
      incomeRateInput.value = String(stored.incomeRate);
    }
    if (Number.isFinite(Number(stored.contractedRate)) && contractedRateInput) {
      contractedRateInput.value = String(stored.contractedRate);
    }
    if (Number.isFinite(Number(stored.taxPercent)) && taxOnCostInput) {
      taxOnCostInput.value = String(stored.taxPercent);
    }
    if (Number.isFinite(Number(stored.pensionPercent)) && pensionOnCostInput) {
      pensionOnCostInput.value = String(stored.pensionPercent);
    }
  } catch (error) {
    console.warn("Could not load carer profitability assumptions.", error);
  }
}

function calculateFinancials(row, assumptions) {
  const revenue = row.confirmedHours * assumptions.incomeRate;
  const baseLabourCost = row.contractedHours * assumptions.contractedRate;
  const onCost = baseLabourCost * ((assumptions.taxPercent + assumptions.pensionPercent) / 100);
  const labourWithOnCost = baseLabourCost + onCost;
  const totalCost = labourWithOnCost + row.travelExpense;
  const profit = revenue - totalCost;
  const utilisation = row.contractedHours > 0 ? (row.confirmedHours / row.contractedHours) * 100 : null;

  return {
    ...row,
    revenue,
    baseLabourCost,
    onCost,
    labourWithOnCost,
    totalCost,
    profit,
    utilisation,
  };
}

function buildReport() {
  if (!carerHoursRows.length) {
    latestReport = null;
    return null;
  }

  const assumptions = getAssumptions();
  const carersByKey = new Map();
  const carersById = new Map();
  const carersByName = new Map();

  carerHoursRows.forEach((row) => {
    const id = getField(row, ["No.", "No", "Carer ID", "ID"]);
    const firstName = getField(row, ["First Name", "Forename"]);
    const surname = getField(row, ["Surname", "Last Name"]);
    const fallbackName = getField(row, ["Carer Name", "Name"]);

    if (!hasCarerIdentity({ id, firstName, surname, fallbackName })) {
      return;
    }

    const name = cleanCell(`${firstName} ${surname}`) || fallbackName;
    const area = getField(row, ["Area", "Care Area", "Region"]) || "Unassigned";
    const confirmedHours = parseHours(getField(row, ["Confirmed (HH:MM)", "Confirmed", "Confirmed Hrs", "Confirmed Hours"]));
    const sourceContractedHours = parseHours(getField(row, ["Contracted Hrs", "Contracted Hours", "Contracted"]));
    const contractedHours = isCentralRole(area) ? 0 : sourceContractedHours;
    const key = id ? `id:${id}` : `name:${normalizeName(name)}`;
    const existing = carersByKey.get(key) || {
      key,
      id,
      name,
      area,
      confirmedHours: 0,
      contractedHours: 0,
      payrollTotalHours: 0,
      travelExpense: 0,
    };

    existing.confirmedHours += confirmedHours;
    existing.contractedHours += contractedHours;
    carersByKey.set(key, existing);
  });

  for (const [key, carer] of carersByKey.entries()) {
    if (carer.id) {
      carersById.set(carer.id, key);
    }
    const normalized = normalizeName(carer.name);
    if (normalized) {
      carersByName.set(normalized, key);
    }
  }

  payrollRows.forEach((row, index) => {
    const id = getField(row, ["Carer ID", "No.", "No", "ID"]);
    const name = getField(row, ["Carer Name", "Name"]);
    const payrollTotalHours = parseHours(getField(row, ["Total Hrs (Decimal)", "Total Hrs", "Total Hours", "Hours"]));
    const travelExpense = parseCurrency(getField(row, ["Total Travel", "Travel Total", "Travel Expenses", "Travel"]));
    const key = (id && carersById.get(id)) || carersByName.get(normalizeName(name));

    if (key && carersByKey.has(key)) {
      const carer = carersByKey.get(key);
      carer.payrollTotalHours += payrollTotalHours;
      carer.travelExpense += travelExpense;
      return;
    }

    if (travelExpense !== 0 || payrollTotalHours !== 0) {
      const payrollOnlyKey = id ? `payroll-id:${id}` : `payroll-name:${normalizeName(name) || index}`;
      carersByKey.set(payrollOnlyKey, {
        key: payrollOnlyKey,
        id,
        name: name || `Payroll row ${index + 1}`,
        area: "Payroll only",
        confirmedHours: 0,
        contractedHours: payrollTotalHours,
        payrollTotalHours,
        travelExpense,
      });
    }
  });

  const carers = Array.from(carersByKey.values())
    .map((row) => ({
      ...row,
      idNumber: parseSortNumber(row.id),
      contractedHours: Math.max(row.contractedHours, row.payrollTotalHours),
    }))
    .map((row) => calculateFinancials(row, assumptions))
    .sort((left, right) => left.area.localeCompare(right.area) || left.name.localeCompare(right.name));

  const areasByName = new Map();
  carers.forEach((carer) => {
    const area = areasByName.get(carer.area) || {
      area: carer.area,
      carerCount: 0,
      confirmedHours: 0,
      contractedHours: 0,
      payrollTotalHours: 0,
      travelExpense: 0,
    };
    area.carerCount += 1;
    area.confirmedHours += carer.confirmedHours;
    area.contractedHours += carer.contractedHours;
    area.payrollTotalHours += carer.payrollTotalHours;
    area.travelExpense += carer.travelExpense;
    areasByName.set(carer.area, area);
  });

  const areas = Array.from(areasByName.values())
    .map((row) => calculateFinancials(row, assumptions))
    .sort((left, right) => left.area.localeCompare(right.area));

  const grandTotal = calculateFinancials(
    carers.reduce(
      (total, carer) => ({
        confirmedHours: total.confirmedHours + carer.confirmedHours,
        contractedHours: total.contractedHours + carer.contractedHours,
        payrollTotalHours: total.payrollTotalHours + carer.payrollTotalHours,
        travelExpense: total.travelExpense + carer.travelExpense,
        carerCount: total.carerCount + 1,
      }),
      { confirmedHours: 0, contractedHours: 0, payrollTotalHours: 0, travelExpense: 0, carerCount: 0 }
    ),
    assumptions
  );

  latestReport = { assumptions, carers, areas, grandTotal };
  return latestReport;
}

function formatHours(value) {
  return decimalFormatter.format(Number(value || 0));
}

function formatCurrency(value) {
  return currencyFormatter.format(Number(value || 0));
}

function formatPercent(value) {
  return value === null || value === undefined ? "n/a" : `${decimalFormatter.format(value)}%`;
}

function renderKpiCard(label, value) {
  const card = document.createElement("article");
  card.className = "profitability-kpi-card";
  const title = document.createElement("h3");
  title.textContent = label;
  const amount = document.createElement("p");
  amount.className = "profitability-kpi-value";
  amount.textContent = value;
  card.append(title, amount);
  return card;
}

function renderKpiSet(container, total) {
  container.innerHTML = "";
  container.append(
    renderKpiCard("Confirmed Hours", formatHours(total.confirmedHours)),
    renderKpiCard("Contracted Hours", formatHours(total.contractedHours)),
    renderKpiCard("Utilisation", formatPercent(total.utilisation)),
    renderKpiCard("Revenue", formatCurrency(total.revenue)),
    renderKpiCard("Labour + On-Cost", formatCurrency(total.labourWithOnCost)),
    renderKpiCard("Travel", formatCurrency(total.travelExpense)),
    renderKpiCard("Estimated Profit", formatCurrency(total.profit))
  );
}

function buildTotalFromCarers(carers, assumptions) {
  return calculateFinancials(
    carers.reduce(
      (total, carer) => ({
        confirmedHours: total.confirmedHours + carer.confirmedHours,
        contractedHours: total.contractedHours + carer.contractedHours,
        payrollTotalHours: total.payrollTotalHours + carer.payrollTotalHours,
        travelExpense: total.travelExpense + carer.travelExpense,
        carerCount: total.carerCount + 1,
      }),
      { confirmedHours: 0, contractedHours: 0, payrollTotalHours: 0, travelExpense: 0, carerCount: 0 }
    ),
    assumptions
  );
}

function numericSortValue(row, key) {
  const value = row?.[key];
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function sortedRows(rows, tableKey) {
  const sort = tableSortState[tableKey];
  if (!sort?.key) {
    return rows;
  }

  const direction = sort.direction === "asc" ? 1 : -1;
  return [...rows].sort((left, right) => {
    const leftValue = numericSortValue(left, sort.key);
    const rightValue = numericSortValue(right, sort.key);

    if (leftValue === null && rightValue === null) {
      return 0;
    }
    if (leftValue === null) {
      return 1;
    }
    if (rightValue === null) {
      return -1;
    }
    return (leftValue - rightValue) * direction;
  });
}

function updateSortButtons() {
  document.querySelectorAll("[data-sort-table][data-sort-key]").forEach((button) => {
    const table = button.getAttribute("data-sort-table");
    const key = button.getAttribute("data-sort-key");
    const sort = tableSortState[table];
    const isActive = sort?.key === key;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
    button.dataset.sortDirection = isActive ? sort.direction : "";
  });
}

function appendCell(row, value) {
  const cell = document.createElement("td");
  cell.textContent = value;
  row.appendChild(cell);
}

function renderAreaRows(rows) {
  areaSummaryBody.innerHTML = "";
  sortedRows(rows, "area").forEach((area) => {
    const tr = document.createElement("tr");
    appendCell(tr, area.area);
    appendCell(tr, String(area.carerCount));
    appendCell(tr, formatHours(area.confirmedHours));
    appendCell(tr, formatHours(area.contractedHours));
    appendCell(tr, formatPercent(area.utilisation));
    appendCell(tr, formatCurrency(area.revenue));
    appendCell(tr, formatCurrency(area.labourWithOnCost));
    appendCell(tr, formatCurrency(area.travelExpense));
    appendCell(tr, formatCurrency(area.profit));
    areaSummaryBody.appendChild(tr);
  });
}

function renderCarerRows(rows) {
  carerSummaryBody.innerHTML = "";
  sortedRows(rows, "carer").forEach((carer) => {
    const tr = document.createElement("tr");
    const selectCell = document.createElement("td");
    selectCell.className = "selection-cell";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = selectedCarerKeys.has(carer.key);
    checkbox.dataset.carerKey = carer.key;
    checkbox.setAttribute("aria-label", `Include ${carer.name} in selected totals`);
    selectCell.appendChild(checkbox);
    tr.appendChild(selectCell);
    appendCell(tr, carer.area);
    appendCell(tr, carer.id || "");
    appendCell(tr, carer.name);
    appendCell(tr, formatHours(carer.confirmedHours));
    appendCell(tr, formatHours(carer.contractedHours));
    appendCell(tr, formatPercent(carer.utilisation));
    appendCell(tr, formatCurrency(carer.travelExpense));
    appendCell(tr, formatCurrency(carer.profit));
    carerSummaryBody.appendChild(tr);
  });
}

function syncSelectedCarers(carers) {
  const validKeys = new Set(carers.map((carer) => carer.key));
  selectedCarerKeys = new Set(Array.from(selectedCarerKeys).filter((key) => validKeys.has(key)));

  if (!hasManualSelection) {
    selectedCarerKeys = new Set(validKeys);
  }

  if (selectAllCarersBtn) {
    selectAllCarersBtn.disabled = !carers.length || selectedCarerKeys.size === carers.length;
  }
  if (clearSelectedCarersBtn) {
    clearSelectedCarersBtn.disabled = !carers.length || selectedCarerKeys.size === 0;
  }
}

function renderSelectedTotals(report) {
  if (!selectedTotalGrid || !selectedTotalMessage) {
    return;
  }

  const selectedCarers = report.carers.filter((carer) => selectedCarerKeys.has(carer.key));
  const total = buildTotalFromCarers(selectedCarers, report.assumptions);
  const label =
    selectedCarers.length === report.carers.length
      ? "all carers"
      : `${selectedCarers.length} selected carer${selectedCarers.length === 1 ? "" : "s"}`;

  selectedTotalMessage.textContent = selectedCarers.length
    ? `Totals for ${label}.`
    : "No carers selected.";
  renderKpiSet(selectedTotalGrid, total);
}

function renderReport() {
  const report = buildReport();

  if (!report) {
    reportOutputPanel.hidden = true;
    exportReportBtn.disabled = true;
    syncSelectedCarers([]);
    if (selectedTotalGrid) {
      selectedTotalGrid.innerHTML = "";
    }
    if (selectedTotalMessage) {
      selectedTotalMessage.textContent = "";
    }
    return;
  }

  renderKpiSet(grandTotalGrid, report.grandTotal);
  syncSelectedCarers(report.carers);
  renderAreaRows(report.areas);
  renderCarerRows(report.carers);
  renderSelectedTotals(report);
  updateSortButtons();
  reportSummaryMessage.textContent = `${report.carers.length} carer row(s), ${report.areas.length} area(s). Payroll hours replace contracted hours where higher; central contracted hours are ignored.`;
  reportOutputPanel.hidden = false;
  exportReportBtn.disabled = false;
}

async function handleCsvUpload(file, label) {
  if (!file) {
    return { rows: [], errors: [] };
  }
  const text = await readFileText(file);
  const parsed = parseCsvText(text);
  if (!parsed.rows.length) {
    throw new Error(`${label} did not contain any data rows.`);
  }
  return parsed;
}

async function handleCarerHoursUpload() {
  try {
    const parsed = await handleCsvUpload(carerHoursCsvInput.files?.[0], "Carer hours CSV");
    carerHoursRows = parsed.rows;
    hasManualSelection = false;
    selectedCarerKeys = new Set();
    renderReport();
    const warning = parsed.errors.length ? ` ${parsed.errors[0]}` : "";
    setUploadStatus(`Loaded ${carerHoursRows.length} carer hours row(s).${warning}`);
  } catch (error) {
    carerHoursRows = [];
    renderReport();
    setUploadStatus(error?.message || "Could not load the carer hours CSV.", true);
  }
}

async function handlePayrollUpload() {
  try {
    const parsed = await handleCsvUpload(payrollCsvInput.files?.[0], "Payroll CSV");
    payrollRows = parsed.rows;
    renderReport();
    const warning = parsed.errors.length ? ` ${parsed.errors[0]}` : "";
    setUploadStatus(`Loaded ${payrollRows.length} payroll row(s).${warning}`);
  } catch (error) {
    payrollRows = [];
    renderReport();
    setUploadStatus(error?.message || "Could not load the payroll CSV.", true);
  }
}

function escapeCsvValue(value) {
  const text = String(value ?? "");
  if (!/[",\n]/.test(text)) {
    return text;
  }
  return `"${text.replaceAll('"', '""')}"`;
}

function toCsvNumber(value) {
  return Number(value || 0).toFixed(2);
}

function buildExportRows() {
  if (!latestReport) {
    return [];
  }

  const rows = [];
  const addRow = (level, row) => {
    rows.push({
      Level: level,
      Area: row.area || "",
      "Carer ID": row.id || "",
      Carer: row.name || row.carer || "",
      "Carer Count": row.carerCount || "",
      "Confirmed Hours": toCsvNumber(row.confirmedHours),
      "Contracted Hours": toCsvNumber(row.contractedHours),
      "Payroll Total Hours": toCsvNumber(row.payrollTotalHours),
      "Utilisation %": row.utilisation === null || row.utilisation === undefined ? "" : toCsvNumber(row.utilisation),
      Revenue: toCsvNumber(row.revenue),
      "Base Labour Cost": toCsvNumber(row.baseLabourCost),
      "On Cost": toCsvNumber(row.onCost),
      "Travel Expense": toCsvNumber(row.travelExpense),
      "Estimated Profit": toCsvNumber(row.profit),
      "Confirmed Hour Income": toCsvNumber(latestReport.assumptions.incomeRate),
      "Contracted Hour Cost": toCsvNumber(latestReport.assumptions.contractedRate),
      "Employer Tax/On-Cost %": toCsvNumber(latestReport.assumptions.taxPercent),
      "Pension %": toCsvNumber(latestReport.assumptions.pensionPercent),
    });
  };

  addRow("Grand Total", { ...latestReport.grandTotal, carer: "Grand Total" });
  latestReport.areas.forEach((area) => addRow("Area", area));
  latestReport.carers.forEach((carer) => addRow("Carer", carer));
  return rows;
}

function downloadCsv(filename, csvText) {
  const blob = new Blob([csvText], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function exportReport() {
  const rows = buildExportRows();
  if (!rows.length) {
    setUploadStatus("Upload the carer hours CSV before exporting.", true);
    return;
  }

  const headers = Object.keys(rows[0]);
  const lines = [headers.map(escapeCsvValue).join(",")];
  rows.forEach((row) => {
    lines.push(headers.map((header) => escapeCsvValue(row[header])).join(","));
  });

  const stamp = reportPeriod?.start ? formatMonthStamp(reportPeriod.start) : formatMonthStamp(new Date());
  downloadCsv(`carer-profitability-${stamp}.csv`, `\uFEFF${lines.join("\n")}`);
  setUploadStatus("Report CSV exported.");
}

async function fetchCurrentUser() {
  return directoryApi.getCurrentUser();
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await fetchCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();

    if (!canAccessPage(role, "reports")) {
      redirectToUnauthorized("reports");
      return;
    }

    renderTopNavigation({ role });
    loadAssumptions();
    updateReportLinks();

    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("reports");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

carerHoursCsvInput?.addEventListener("change", () => {
  void handleCarerHoursUpload();
});

payrollCsvInput?.addEventListener("change", () => {
  void handlePayrollUpload();
});

[incomeRateInput, contractedRateInput, taxOnCostInput, pensionOnCostInput].forEach((input) => {
  input?.addEventListener("input", () => {
    saveAssumptions();
    renderReport();
  });
});

carerSummaryBody?.addEventListener("change", (event) => {
  const checkbox = event.target?.closest?.('input[type="checkbox"][data-carer-key]');
  if (!checkbox || !carerSummaryBody.contains(checkbox)) {
    return;
  }

  hasManualSelection = true;
  const key = checkbox.dataset.carerKey;
  if (checkbox.checked) {
    selectedCarerKeys.add(key);
  } else {
    selectedCarerKeys.delete(key);
  }

  if (latestReport) {
    syncSelectedCarers(latestReport.carers);
    renderSelectedTotals(latestReport);
  }
});

selectAllCarersBtn?.addEventListener("click", () => {
  if (!latestReport) {
    return;
  }
  hasManualSelection = true;
  selectedCarerKeys = new Set(latestReport.carers.map((carer) => carer.key));
  renderReport();
});

clearSelectedCarersBtn?.addEventListener("click", () => {
  hasManualSelection = true;
  selectedCarerKeys = new Set();
  if (latestReport) {
    renderReport();
  }
});

document.querySelectorAll("[data-sort-table][data-sort-key]").forEach((button) => {
  button.addEventListener("click", () => {
    const table = button.getAttribute("data-sort-table");
    const key = button.getAttribute("data-sort-key");
    if (!tableSortState[table]) {
      return;
    }

    const current = tableSortState[table];
    tableSortState[table] = {
      key,
      direction: current.key === key && current.direction === "desc" ? "asc" : "desc",
    };

    if (latestReport) {
      if (table === "area") {
        renderAreaRows(latestReport.areas);
      }
      if (table === "carer") {
        renderCarerRows(latestReport.carers);
      }
    }
    updateSortButtons();
  });
});

exportReportBtn?.addEventListener("click", exportReport);

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

void init();
