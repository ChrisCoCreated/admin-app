import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260601";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const periodPresetSelect = document.getElementById("periodPresetSelect");
const dateRangeMessage = document.getElementById("dateRangeMessage");
const carerHoursReportLink = document.getElementById("carerHoursReportLink");
const utilisationCsvInput = document.getElementById("utilisationCsvInput");
const uploadStatusMessage = document.getElementById("uploadStatusMessage");
const exportReportBtn = document.getElementById("exportReportBtn");
const reportOutputPanel = document.getElementById("reportOutputPanel");
const reportSummaryMessage = document.getElementById("reportSummaryMessage");
const areaFilterSelect = document.getElementById("areaFilterSelect");
const showAllPeopleBtn = document.getElementById("showAllPeopleBtn");
const hideAllPeopleBtn = document.getElementById("hideAllPeopleBtn");
const summaryGrid = document.getElementById("summaryGrid");
const utilisationReportBody = document.getElementById("utilisationReportBody");

const CARER_HOURS_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carersHoursRpt.php";
const DEFAULT_AREA = "East Kent";
const FIELD_ALIASES = {
  surname: ["Surname"],
  firstName: ["First Name", "Firstname"],
  area: ["Area"],
  scheduledHours: ["Scheduled Hrs (HH:MM)", "Scheduled Hrs", "Scheduled Hours", "Scheduled"],
  confirmedHours: ["Confirmed (HH:MM)", "Confirmed Hrs", "Confirmed Hours", "Confirmed"],
  contractedHours: ["Contracted Hrs (HH:MM)", "Contracted Hrs", "Contracted Hours", "Contracted"],
};

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let latestRows = [];
let hiddenPeople = new Set();
let latestPeriod = getDateRangeForPreset("this_month");

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setUploadStatus(message, isError = false) {
  uploadStatusMessage.textContent = message;
  uploadStatusMessage.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "finance").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function cleanCell(value) {
  return String(value ?? "").replace(/^\uFEFF/, "").replace(/\r/g, "").trim();
}

function normalizeHeader(value) {
  return cleanCell(value)
    .toLowerCase()
    .replace(/&amp;/g, "and")
    .replace(/[^a-z0-9]+/g, "");
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

function getField(row, headerMap, aliases) {
  for (const alias of aliases) {
    const actualHeader = headerMap.get(normalizeHeader(alias));
    if (actualHeader) {
      return cleanCell(row[actualHeader]);
    }
  }
  return "";
}

function parseHours(value) {
  const raw = cleanCell(value).replace(/,/g, "");
  if (!raw) {
    return 0;
  }

  const timeMatch = raw.match(/^(-?\d+(?:\.\d+)?):(\d{1,2})$/);
  if (timeMatch) {
    const hours = Number(timeMatch[1]);
    const minutes = Number(timeMatch[2]);
    if (Number.isFinite(hours) && Number.isFinite(minutes)) {
      const sign = hours < 0 ? -1 : 1;
      return hours + sign * (minutes / 60);
    }
  }

  const numeric = Number(raw);
  return Number.isFinite(numeric) ? numeric : 0;
}

function formatHours(value) {
  const totalMinutes = Math.round(Number(value || 0) * 60);
  const sign = totalMinutes < 0 ? "-" : "";
  const absoluteMinutes = Math.abs(totalMinutes);
  const hours = Math.floor(absoluteMinutes / 60);
  const minutes = String(absoluteMinutes % 60).padStart(2, "0");
  return `${sign}${hours}:${minutes}`;
}

function formatPercent(value) {
  if (value === null || value === undefined || !Number.isFinite(value)) {
    return "n/a";
  }
  return `${value.toFixed(1)}%`;
}

function utilisation(numerator, contractedHours) {
  return contractedHours > 0 ? (numerator / contractedHours) * 100 : null;
}

function personKey(row) {
  return [row.surname, row.firstName, row.area].map((value) => normalizeHeader(value)).join("|");
}

function mapRows(parsed) {
  const headerMap = new Map(parsed.headers.map((header) => [normalizeHeader(header), header]));
  const requiredFields = [
    ["Surname", FIELD_ALIASES.surname],
    ["First Name", FIELD_ALIASES.firstName],
    ["Area", FIELD_ALIASES.area],
    ["Scheduled Hrs (HH:MM)", FIELD_ALIASES.scheduledHours],
    ["Confirmed (HH:MM)", FIELD_ALIASES.confirmedHours],
    ["Contracted Hrs (HH:MM)", FIELD_ALIASES.contractedHours],
  ];
  const missingColumns = requiredFields
    .filter(([, aliases]) => !aliases.some((alias) => headerMap.has(normalizeHeader(alias))))
    .map(([label]) => label);
  if (missingColumns.length) {
    throw new Error(`Missing required column(s): ${missingColumns.join(", ")}.`);
  }

  return parsed.rows
    .map((row) => {
      const surname = getField(row, headerMap, FIELD_ALIASES.surname);
      const firstName = getField(row, headerMap, FIELD_ALIASES.firstName);
      const area = getField(row, headerMap, FIELD_ALIASES.area);
      const scheduledHours = parseHours(getField(row, headerMap, FIELD_ALIASES.scheduledHours));
      const confirmedHours = parseHours(getField(row, headerMap, FIELD_ALIASES.confirmedHours));
      const contractedHours = parseHours(getField(row, headerMap, FIELD_ALIASES.contractedHours));
      return {
        surname,
        firstName,
        area,
        scheduledHours,
        confirmedHours,
        contractedHours,
        projectedUtilisation: utilisation(scheduledHours, contractedHours),
        actualUtilisation: utilisation(confirmedHours, contractedHours),
      };
    })
    .filter((row) => row.surname || row.firstName || row.area);
}

function addMonths(baseDate, offset) {
  return new Date(baseDate.getFullYear(), baseDate.getMonth() + offset, 1);
}

function getDateRangeForPreset(preset) {
  const today = new Date();
  let offset = 0;
  if (preset === "last_month") {
    offset = -1;
  }
  if (preset === "next_month") {
    offset = 1;
  }
  const start = addMonths(today, offset);
  const end = new Date(start.getFullYear(), start.getMonth() + 1, 0);
  return { start, end };
}

function formatDateParam(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  return `${day}-${month}-${date.getFullYear()}`;
}

function formatReadableDate(date) {
  return date.toLocaleDateString("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function buildCarerHoursUrl(start, end) {
  const url = new URL(CARER_HOURS_BASE_URL);
  url.searchParams.set("jobtype", "All");
  url.searchParams.set("start", formatDateParam(start));
  url.searchParams.set("finish", formatDateParam(end));
  return url.toString();
}

function refreshPeriodLink() {
  latestPeriod = getDateRangeForPreset(periodPresetSelect?.value || "this_month");
  if (dateRangeMessage) {
    dateRangeMessage.textContent = `Selected period: ${formatReadableDate(latestPeriod.start)} to ${formatReadableDate(latestPeriod.end)}.`;
  }
  if (carerHoursReportLink) {
    carerHoursReportLink.href = buildCarerHoursUrl(latestPeriod.start, latestPeriod.end);
  }
}

function getAreas(rows) {
  return Array.from(new Set(rows.map((row) => row.area).filter(Boolean))).sort((left, right) =>
    left.localeCompare(right, undefined, { sensitivity: "base" })
  );
}

function selectedArea() {
  return areaFilterSelect?.value || DEFAULT_AREA;
}

function filteredRows() {
  const area = selectedArea();
  return latestRows.filter((row) => area === "All" || row.area === area);
}

function visibleRows() {
  return filteredRows().filter((row) => !hiddenPeople.has(personKey(row)));
}

function renderAreaOptions() {
  const areas = getAreas(latestRows);
  areaFilterSelect.innerHTML = "";

  const allOption = document.createElement("option");
  allOption.value = "All";
  allOption.textContent = "All areas";
  areaFilterSelect.appendChild(allOption);

  for (const area of areas) {
    const option = document.createElement("option");
    option.value = area;
    option.textContent = area;
    areaFilterSelect.appendChild(option);
  }

  areaFilterSelect.value = areas.includes(DEFAULT_AREA) ? DEFAULT_AREA : "All";
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

function totalsForRows(rows) {
  const totals = rows.reduce(
    (summary, row) => ({
      scheduledHours: summary.scheduledHours + row.scheduledHours,
      confirmedHours: summary.confirmedHours + row.confirmedHours,
      contractedHours: summary.contractedHours + row.contractedHours,
    }),
    { scheduledHours: 0, confirmedHours: 0, contractedHours: 0 }
  );
  return {
    ...totals,
    projectedUtilisation: utilisation(totals.scheduledHours, totals.contractedHours),
    actualUtilisation: utilisation(totals.confirmedHours, totals.contractedHours),
  };
}

function renderSummary(rows) {
  const totals = totalsForRows(rows);
  summaryGrid.innerHTML = "";
  summaryGrid.append(
    renderKpiCard("People Shown", String(rows.length)),
    renderKpiCard("Scheduled", formatHours(totals.scheduledHours)),
    renderKpiCard("Confirmed", formatHours(totals.confirmedHours)),
    renderKpiCard("Contracted", formatHours(totals.contractedHours)),
    renderKpiCard("Projected Utilisation", formatPercent(totals.projectedUtilisation)),
    renderKpiCard("Actual Utilisation", formatPercent(totals.actualUtilisation))
  );
}

function makeCopyButton(value) {
  const button = document.createElement("button");
  button.className = "utilisation-copy-button";
  button.type = "button";
  button.textContent = formatPercent(value);
  button.title = "Copy utilisation percentage";
  button.addEventListener("click", async () => {
    const text = formatPercent(value);
    try {
      await navigator.clipboard.writeText(text);
      button.classList.add("is-copied");
      button.textContent = "Copied";
      window.setTimeout(() => {
        button.classList.remove("is-copied");
        button.textContent = text;
      }, 900);
    } catch (error) {
      console.error("Could not copy utilisation percentage", error);
      setUploadStatus("Could not copy percentage.", true);
    }
  });
  return button;
}

function renderRows() {
  const rows = filteredRows();
  const includedRows = rows.filter((row) => !hiddenPeople.has(personKey(row)));
  utilisationReportBody.innerHTML = "";

  if (!rows.length) {
    utilisationReportBody.innerHTML = '<tr><td colspan="8" class="muted">No rows for this area.</td></tr>';
  } else {
    for (const row of rows) {
      const key = personKey(row);
      const isHidden = hiddenPeople.has(key);
      const tr = document.createElement("tr");
      tr.className = isHidden ? "utilisation-person-hidden" : "";

      const nameCell = document.createElement("td");
      nameCell.className = "profitability-sticky-column";
      nameCell.textContent = [row.surname, row.firstName].filter(Boolean).join(", ");
      tr.appendChild(nameCell);

      const visibleCell = document.createElement("td");
      const visibleButton = document.createElement("button");
      visibleButton.className = "secondary utilisation-visibility-button";
      visibleButton.type = "button";
      visibleButton.textContent = isHidden ? "Show" : "Hide";
      visibleButton.addEventListener("click", () => {
        if (hiddenPeople.has(key)) {
          hiddenPeople.delete(key);
        } else {
          hiddenPeople.add(key);
        }
        renderReport();
      });
      visibleCell.appendChild(visibleButton);
      tr.appendChild(visibleCell);

      for (const value of [
        row.area,
        formatHours(row.scheduledHours),
        formatHours(row.confirmedHours),
        formatHours(row.contractedHours),
      ]) {
        const td = document.createElement("td");
        td.textContent = value;
        tr.appendChild(td);
      }

      const projectedCell = document.createElement("td");
      projectedCell.appendChild(makeCopyButton(row.projectedUtilisation));
      tr.appendChild(projectedCell);

      const actualCell = document.createElement("td");
      actualCell.appendChild(makeCopyButton(row.actualUtilisation));
      tr.appendChild(actualCell);

      utilisationReportBody.appendChild(tr);
    }
  }

  renderSummary(includedRows);
  reportSummaryMessage.textContent =
    `Showing ${includedRows.length} of ${rows.length} row(s) for ${selectedArea() === "All" ? "all areas" : selectedArea()}. ` +
    `${hiddenPeople.size} hidden across the uploaded report.`;
}

function renderReport() {
  if (!latestRows.length) {
    reportOutputPanel.hidden = true;
    exportReportBtn?.setAttribute("disabled", "disabled");
    return;
  }
  reportOutputPanel.hidden = false;
  exportReportBtn?.removeAttribute("disabled");
  renderRows();
}

function escapeCsvValue(value) {
  const text = String(value ?? "");
  if (!/[",\n]/.test(text)) {
    return text;
  }
  return `"${text.replaceAll('"', '""')}"`;
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

function handleExport() {
  const rows = visibleRows();
  if (!rows.length) {
    setUploadStatus("There are no visible rows to export.", true);
    return;
  }
  const headers = [
    "Surname",
    "First Name",
    "Area",
    "Scheduled Hrs (HH:MM)",
    "Confirmed (HH:MM)",
    "Contracted Hrs (HH:MM)",
    "Projected Utilisation %",
    "Actual Utilisation %",
  ];
  const lines = [headers.map(escapeCsvValue).join(",")];
  for (const row of rows) {
    lines.push(
      [
        row.surname,
        row.firstName,
        row.area,
        formatHours(row.scheduledHours),
        formatHours(row.confirmedHours),
        formatHours(row.contractedHours),
        formatPercent(row.projectedUtilisation),
        formatPercent(row.actualUtilisation),
      ]
        .map(escapeCsvValue)
        .join(",")
    );
  }
  const stamp = `${latestPeriod.start.getFullYear()}${String(latestPeriod.start.getMonth() + 1).padStart(2, "0")}`;
  downloadCsv(`utilisation-report-${stamp}-${normalizeHeader(selectedArea()) || "all"}.csv`, `\uFEFF${lines.join("\n")}`);
  setUploadStatus("Utilisation report CSV downloaded.");
}

async function handleCsvUpload() {
  try {
    const file = utilisationCsvInput?.files?.[0];
    if (!file) {
      setUploadStatus("Choose a carer hours CSV first.", true);
      return;
    }

    const text = await file.text();
    const parsed = parseCsvText(text);
    latestRows = mapRows(parsed).sort(
      (left, right) =>
        left.area.localeCompare(right.area, undefined, { sensitivity: "base" }) ||
        left.surname.localeCompare(right.surname, undefined, { sensitivity: "base" }) ||
        left.firstName.localeCompare(right.firstName, undefined, { sensitivity: "base" })
    );
    hiddenPeople = new Set();
    renderAreaOptions();
    renderReport();

    const warning = parsed.errors.length ? ` ${parsed.errors.length} CSV warning(s) were ignored.` : "";
    setUploadStatus(`Loaded ${latestRows.length} utilisation row(s).${warning}`);
  } catch (error) {
    console.error("Utilisation CSV import failed", error);
    latestRows = [];
    hiddenPeople = new Set();
    renderReport();
    setUploadStatus(error?.message || "Could not read the utilisation CSV.", true);
  }
}

async function init() {
  try {
    setStatus("Checking access...");
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();

    if (!canAccessPage(role, "finance")) {
      redirectToUnauthorized("finance");
      return;
    }

    renderTopNavigation({ role });
    document.body.classList.remove("auth-pending");
    setStatus(`Signed in as ${profile?.email || "unknown"} (${role || "unknown role"}).`);
    refreshPeriodLink();
    renderReport();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("finance");
      return;
    }
    console.error("Utilisation report failed to initialise", error);
    setStatus(error?.message || "Unable to load the utilisation report.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

periodPresetSelect?.addEventListener("change", refreshPeriodLink);
utilisationCsvInput?.addEventListener("change", handleCsvUpload);
areaFilterSelect?.addEventListener("change", renderReport);
exportReportBtn?.addEventListener("click", handleExport);
showAllPeopleBtn?.addEventListener("click", () => {
  for (const row of filteredRows()) {
    hiddenPeople.delete(personKey(row));
  }
  renderReport();
});
hideAllPeopleBtn?.addEventListener("click", () => {
  for (const row of filteredRows()) {
    hiddenPeople.add(personKey(row));
  }
  renderReport();
});

void init();
