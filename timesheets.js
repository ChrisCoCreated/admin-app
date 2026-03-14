import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260314";

const MAX_RANGE_DAYS = 60;

const signOutBtn = document.getElementById("signOutBtn");
const carerSearchInput = document.getElementById("carerSearchInput");
const carerSelect = document.getElementById("carerSelect");
const presetButtons = Array.from(document.querySelectorAll("[data-preset]"));
const customRangeFields = document.getElementById("customRangeFields");
const dateStartInput = document.getElementById("dateStartInput");
const dateFinishInput = document.getElementById("dateFinishInput");
const fetchTimesheetsBtn = document.getElementById("fetchTimesheetsBtn");
const selectedRangeMessage = document.getElementById("selectedRangeMessage");
const statusMessage = document.getElementById("statusMessage");
const summaryPanel = document.getElementById("summaryPanel");
const summaryMessage = document.getElementById("summaryMessage");
const summaryTotalRows = document.getElementById("summaryTotalRows");
const summaryConfirmed = document.getElementById("summaryConfirmed");
const summaryUnconfirmed = document.getElementById("summaryUnconfirmed");
const summaryClients = document.getElementById("summaryClients");
const timesheetsTableBody = document.getElementById("timesheetsTableBody");
const timesheetsTableFoot = document.getElementById("timesheetsTableFoot");
const totalsScheduledCell = document.getElementById("totalsScheduledCell");
const totalsActualCell = document.getElementById("totalsActualCell");
const totalsVarianceCell = document.getElementById("totalsVarianceCell");
const totalsActualisationCell = document.getElementById("totalsActualisationCell");
const totalsConfirmedCell = document.getElementById("totalsConfirmedCell");
const emptyState = document.getElementById("emptyState");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let allCarers = [];
let filteredCarers = [];
let selectedPreset = "yesterday";

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "timesheets").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function toDateInputValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function toStartOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function addDays(baseDate, days) {
  return new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate() + days);
}

function getMondayOfWeek(baseDate) {
  const date = toStartOfDay(baseDate);
  const day = date.getDay();
  const daysSinceMonday = (day + 6) % 7;
  return addDays(date, -daysSinceMonday);
}

function getLastMonthRange() {
  const today = new Date();
  const start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
  const end = new Date(today.getFullYear(), today.getMonth(), 0);
  return { start, end };
}

function getLastWeekRange() {
  const thisWeekMonday = getMondayOfWeek(new Date());
  const start = addDays(thisWeekMonday, -7);
  const end = addDays(start, 6);
  return { start, end };
}

function getYesterdayRange() {
  const end = addDays(toStartOfDay(new Date()), -1);
  return { start: end, end };
}

function getInclusiveDayCount(start, end) {
  return Math.floor((end.getTime() - start.getTime()) / 86_400_000) + 1;
}

function formatReadableDate(date) {
  return date.toLocaleDateString("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
  });
}

function formatDateTime(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "-";
  }
  const normalized = raw.includes("T") ? raw : raw.replace(" ", "T");
  const parsed = new Date(normalized);
  if (Number.isNaN(parsed.getTime())) {
    return raw;
  }
  return parsed.toLocaleString("en-GB", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function parseDateValue(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return null;
  }
  const parsed = new Date(raw.includes("T") ? raw : `${raw}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function parseTimeToMinutes(value) {
  const raw = String(value || "").trim();
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(raw);
  if (!match) {
    return null;
  }
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  if (!Number.isFinite(hours) || !Number.isFinite(minutes)) {
    return null;
  }
  return hours * 60 + minutes;
}

function getDurationMinutesFromClockTimes(startValue, endValue) {
  const start = parseTimeToMinutes(startValue);
  const end = parseTimeToMinutes(endValue);
  if (start === null || end === null) {
    return 0;
  }
  const diff = end - start;
  return diff >= 0 ? diff : 0;
}

function getDurationMinutesFromDateTimes(startValue, endValue) {
  const start = parseDateValue(startValue);
  const end = parseDateValue(endValue);
  if (!start || !end) {
    return 0;
  }
  const diff = Math.round((end.getTime() - start.getTime()) / 60_000);
  return diff >= 0 ? diff : 0;
}

function formatDuration(minutes) {
  const totalMinutes = Math.max(0, Math.round(Number(minutes) || 0));
  const hours = Math.floor(totalMinutes / 60);
  const remainingMinutes = totalMinutes % 60;
  return `${hours}h ${String(remainingMinutes).padStart(2, "0")}m`;
}

function formatVariance(minutes) {
  const numeric = Math.round(Number(minutes) || 0);
  if (numeric === 0) {
    return "0h 00m";
  }
  const sign = numeric > 0 ? "+" : "-";
  return `${sign}${formatDuration(Math.abs(numeric))}`;
}

function formatPercentage(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return "-";
  }
  return `${numeric.toFixed(1)}%`;
}

function getActualisationPercent(actualMinutes, scheduledMinutes) {
  if (!scheduledMinutes) {
    return 0;
  }
  return (actualMinutes / scheduledMinutes) * 100;
}

function getAppliedRange(appliedQuery) {
  if (appliedQuery?.date) {
    const date = parseDateValue(appliedQuery.date);
    return date ? { start: date, finish: date } : null;
  }
  const start = parseDateValue(appliedQuery?.datestart);
  const finish = parseDateValue(appliedQuery?.datefinish);
  if (!start || !finish) {
    return null;
  }
  return { start, finish };
}

function getContractedMinutesForPeriod(carer, appliedQuery) {
  const range = getAppliedRange(appliedQuery);
  if (!range) {
    return 0;
  }

  const contracted = carer?.raw?.contracted_hrs || carer?.raw?.contracted_hours || null;
  const weekdayMinutes = {
    0: Number(contracted?.sunday || 0) * 60,
    1: Number(contracted?.monday || 0) * 60,
    2: Number(contracted?.tuesday || 0) * 60,
    3: Number(contracted?.wednesday || 0) * 60,
    4: Number(contracted?.thursday || 0) * 60,
    5: Number(contracted?.friday || 0) * 60,
    6: Number(contracted?.saturday || 0) * 60,
  };

  if (Object.values(weekdayMinutes).some((value) => Number.isFinite(value) && value > 0)) {
    let total = 0;
    for (let cursor = new Date(range.start); cursor.getTime() <= range.finish.getTime(); cursor = addDays(cursor, 1)) {
      total += Number(weekdayMinutes[cursor.getDay()] || 0);
    }
    return total;
  }

  const weeklyHours = Number(carer?.contractedHours || 0);
  if (!Number.isFinite(weeklyHours) || weeklyHours <= 0) {
    return 0;
  }
  return (weeklyHours * 60 * getInclusiveDayCount(range.start, range.finish)) / 7;
}

function enrichTimesheet(item) {
  const scheduledMinutes = getDurationMinutesFromClockTimes(item.dueIn, item.dueOut);
  const actualMinutes = getDurationMinutesFromDateTimes(item.logIn, item.logOut);
  const varianceMinutes = actualMinutes - scheduledMinutes;
  return {
    ...item,
    scheduledMinutes,
    actualMinutes,
    varianceMinutes,
    actualisationPercent: getActualisationPercent(actualMinutes, scheduledMinutes),
  };
}

function getCarerLabel(carer) {
  const name = String(carer?.name || "Unnamed carer").trim();
  const id = String(carer?.id || "").trim();
  const area = String(carer?.area || "").trim();
  return [name, id ? `#${id}` : "", area].filter(Boolean).join(" - ");
}

function renderCarerOptions() {
  const existingValue = String(carerSelect?.value || "");
  const query = normalizeText(carerSearchInput?.value);

  filteredCarers = allCarers.filter((carer) => {
    if (!query) {
      return true;
    }
    return (
      normalizeText(carer?.name).includes(query) ||
      normalizeText(carer?.id).includes(query) ||
      normalizeText(carer?.area).includes(query) ||
      normalizeText(carer?.postcode).includes(query)
    );
  });

  carerSelect.innerHTML = "";
  if (!filteredCarers.length) {
    carerSelect.innerHTML = '<option value="">No carers match that search</option>';
    return;
  }

  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = "Select a carer";
  carerSelect.appendChild(placeholderOption);

  for (const carer of filteredCarers) {
    const option = document.createElement("option");
    option.value = String(carer.id || "");
    option.textContent = getCarerLabel(carer);
    carerSelect.appendChild(option);
  }

  const canRestore = filteredCarers.some((carer) => String(carer.id || "") === existingValue);
  carerSelect.value = canRestore ? existingValue : "";
}

function applyPreset(preset) {
  selectedPreset = preset;

  for (const button of presetButtons) {
    button.classList.toggle("active", button.dataset.preset === preset);
  }

  const isCustom = preset === "custom";
  customRangeFields.hidden = !isCustom;

  if (!isCustom) {
    const range =
      preset === "last_month" ? getLastMonthRange() : preset === "last_week" ? getLastWeekRange() : getYesterdayRange();
    dateStartInput.value = toDateInputValue(range.start);
    dateFinishInput.value = toDateInputValue(range.end);
  }

  updateRangeMessage();
}

function updateRangeMessage() {
  if (selectedPreset === "custom") {
    const start = dateStartInput.value;
    const finish = dateFinishInput.value;
    selectedRangeMessage.textContent =
      start && finish ? `Range: ${start} to ${finish}` : "Range: choose a bespoke start and finish date";
    return;
  }

  if (selectedPreset === "last_month") {
    selectedRangeMessage.textContent = "Range: last month";
    return;
  }

  if (selectedPreset === "last_week") {
    selectedRangeMessage.textContent = "Range: last week";
    return;
  }

  selectedRangeMessage.textContent = "Range: yesterday";
}

function getSelectedCarer() {
  const carerId = String(carerSelect?.value || "").trim();
  if (!carerId) {
    return null;
  }
  return allCarers.find((carer) => String(carer.id || "") === carerId) || null;
}

function getRequestQuery() {
  if (selectedPreset === "custom") {
    const start = String(dateStartInput.value || "").trim();
    const finish = String(dateFinishInput.value || "").trim();
    if (!start || !finish) {
      throw new Error("Choose both start and finish dates for a bespoke range.");
    }
    const startDate = new Date(`${start}T00:00:00`);
    const finishDate = new Date(`${finish}T00:00:00`);
    if (Number.isNaN(startDate.getTime()) || Number.isNaN(finishDate.getTime())) {
      throw new Error("Enter valid bespoke dates.");
    }
    if (startDate.getTime() > finishDate.getTime()) {
      throw new Error("The bespoke start date must be on or before the finish date.");
    }
    if (getInclusiveDayCount(startDate, finishDate) > MAX_RANGE_DAYS) {
      throw new Error("Bespoke ranges can be up to 60 days.");
    }
    return {
      datestart: start,
      datefinish: finish,
    };
  }

  const range =
    selectedPreset === "last_month" ? getLastMonthRange() : selectedPreset === "last_week" ? getLastWeekRange() : getYesterdayRange();
  const start = toDateInputValue(range.start);
  const finish = toDateInputValue(range.end);
  if (start === finish) {
    return { date: start };
  }
  return {
    datestart: start,
    datefinish: finish,
  };
}

function renderSummary(timesheets, selectedCarer, appliedQuery) {
  const confirmedCount = timesheets.filter((item) => item.timeConfirmed).length;
  const uniqueClients = new Set(
    timesheets.map((item) => String(item.clientId || item.clientName || "").trim()).filter(Boolean)
  );
  const scheduledMinutes = timesheets.reduce((sum, item) => sum + item.scheduledMinutes, 0);
  const actualMinutes = timesheets.reduce((sum, item) => sum + item.actualMinutes, 0);
  const contractedMinutes = getContractedMinutesForPeriod(selectedCarer, appliedQuery);
  const actualisationPercent = getActualisationPercent(actualMinutes, scheduledMinutes);

  const rangeText = appliedQuery.date
    ? formatReadableDate(new Date(`${appliedQuery.date}T00:00:00`))
    : `${formatReadableDate(new Date(`${appliedQuery.datestart}T00:00:00`))} to ${formatReadableDate(
        new Date(`${appliedQuery.datefinish}T00:00:00`)
      )}`;

  summaryTotalRows.textContent = formatDuration(scheduledMinutes);
  summaryConfirmed.textContent = formatDuration(actualMinutes);
  summaryUnconfirmed.textContent = formatDuration(contractedMinutes);
  summaryClients.textContent = formatPercentage(actualisationPercent);
  summaryMessage.textContent = `${selectedCarer?.name || "Selected carer"} for ${rangeText}. ${timesheets.length} row(s), ${uniqueClients.size} client(s), ${confirmedCount} confirmed.`;
  summaryPanel.hidden = false;
}

function renderTimesheets(timesheets) {
  timesheetsTableBody.innerHTML = "";
  if (timesheetsTableFoot) {
    timesheetsTableFoot.hidden = true;
  }

  if (!timesheets.length) {
    emptyState.textContent = "No timesheets were returned for that selection.";
    summaryPanel.hidden = true;
    return;
  }

  emptyState.textContent = "";
  const scheduledTotal = timesheets.reduce((sum, item) => sum + item.scheduledMinutes, 0);
  const actualTotal = timesheets.reduce((sum, item) => sum + item.actualMinutes, 0);
  const varianceTotal = actualTotal - scheduledTotal;
  const confirmedCount = timesheets.filter((item) => item.timeConfirmed).length;

  for (const item of timesheets) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item.date || "-")}</td>
      <td>${escapeHtml(item.carerName || "-")}</td>
      <td>${escapeHtml(item.clientName || "-")}</td>
      <td>${escapeHtml(item.jobType || "-")}</td>
      <td>${escapeHtml(formatDuration(item.scheduledMinutes))}</td>
      <td>${escapeHtml(formatDuration(item.actualMinutes))}</td>
      <td>${escapeHtml(formatVariance(item.varianceMinutes))}</td>
      <td>${escapeHtml(formatPercentage(item.actualisationPercent))}</td>
      <td>${item.timeConfirmed ? "Yes" : "No"}</td>
    `;
    timesheetsTableBody.appendChild(tr);
  }

  if (timesheetsTableFoot) {
    timesheetsTableFoot.hidden = false;
  }
  if (totalsScheduledCell) {
    totalsScheduledCell.textContent = formatDuration(scheduledTotal);
  }
  if (totalsActualCell) {
    totalsActualCell.textContent = formatDuration(actualTotal);
  }
  if (totalsVarianceCell) {
    totalsVarianceCell.textContent = formatVariance(varianceTotal);
  }
  if (totalsActualisationCell) {
    totalsActualisationCell.textContent = formatPercentage(getActualisationPercent(actualTotal, scheduledTotal));
  }
  if (totalsConfirmedCell) {
    totalsConfirmedCell.textContent = `${confirmedCount}/${timesheets.length}`;
  }
}

async function fetchTimesheets() {
  const selectedCarer = getSelectedCarer();
  if (!selectedCarer) {
    setStatus("Choose a carer before fetching timesheets.", true);
    return;
  }

  try {
    fetchTimesheetsBtn.disabled = true;
    const query = getRequestQuery();
    setStatus(`Loading timesheets for ${selectedCarer.name || "selected carer"}...`);
    const payload = await directoryApi.listTimesheets({
      carer_id: selectedCarer.id,
      ...query,
      per_page: 200,
    });
    const timesheets = Array.isArray(payload?.timesheets) ? payload.timesheets.map(enrichTimesheet) : [];

    renderTimesheets(timesheets);
    if (timesheets.length) {
      renderSummary(timesheets, selectedCarer, query);
    } else {
      summaryPanel.hidden = true;
    }
    setStatus(`Loaded ${timesheets.length} timesheet row(s).`);
  } catch (error) {
    console.error(error);
    timesheetsTableBody.innerHTML = "";
    summaryPanel.hidden = true;
    emptyState.textContent = "Choose a carer and fetch a range to see timesheets.";
    setStatus(error?.message || "Could not load timesheets.", true);
  } finally {
    fetchTimesheetsBtn.disabled = false;
  }
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();
    if (!canAccessPage(role, "timesheets")) {
      redirectToUnauthorized("timesheets");
      return;
    }

    renderTopNavigation({ role });
    applyPreset("yesterday");

    const carersPayload = await directoryApi.listCarers({ limit: 500 });
    allCarers = Array.isArray(carersPayload?.carers) ? carersPayload.carers : [];
    renderCarerOptions();
    setStatus(`Loaded ${allCarers.length} carer(s). Choose one to begin.`);
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not initialise timesheets page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

carerSearchInput?.addEventListener("input", renderCarerOptions);
dateStartInput?.addEventListener("change", updateRangeMessage);
dateFinishInput?.addEventListener("change", updateRangeMessage);
fetchTimesheetsBtn?.addEventListener("click", () => {
  void fetchTimesheets();
});

for (const button of presetButtons) {
  button.addEventListener("click", () => {
    applyPreset(String(button.dataset.preset || "yesterday"));
  });
}

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

void init();
