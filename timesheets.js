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

  const rangeText = appliedQuery.date
    ? formatReadableDate(new Date(`${appliedQuery.date}T00:00:00`))
    : `${formatReadableDate(new Date(`${appliedQuery.datestart}T00:00:00`))} to ${formatReadableDate(
        new Date(`${appliedQuery.datefinish}T00:00:00`)
      )}`;

  summaryTotalRows.textContent = String(timesheets.length);
  summaryConfirmed.textContent = String(confirmedCount);
  summaryUnconfirmed.textContent = String(timesheets.length - confirmedCount);
  summaryClients.textContent = String(uniqueClients.size);
  summaryMessage.textContent = `${selectedCarer?.name || "Selected carer"} for ${rangeText}`;
  summaryPanel.hidden = false;
}

function renderTimesheets(timesheets) {
  timesheetsTableBody.innerHTML = "";

  if (!timesheets.length) {
    emptyState.textContent = "No timesheets were returned for that selection.";
    summaryPanel.hidden = true;
    return;
  }

  emptyState.textContent = "";

  for (const item of timesheets) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(item.date || "-")}</td>
      <td>${escapeHtml(item.carerName || "-")}</td>
      <td>${escapeHtml(item.clientName || "-")}</td>
      <td>${escapeHtml(item.jobType || "-")}</td>
      <td>${escapeHtml(item.dueIn || "-")}</td>
      <td>${escapeHtml(item.dueOut || "-")}</td>
      <td>${escapeHtml(formatDateTime(item.logIn))}</td>
      <td>${escapeHtml(formatDateTime(item.logOut))}</td>
      <td>${item.timeConfirmed ? "Yes" : "No"}</td>
    `;
    timesheetsTableBody.appendChild(tr);
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
    const timesheets = Array.isArray(payload?.timesheets) ? payload.timesheets : [];

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
