import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const signOutBtn = document.getElementById("signOutBtn");
const staffPostcodeInput = document.getElementById("staffPostcodeInput");
const carerSearchInput = document.getElementById("carerSearchInput");
const carerSearchResults = document.getElementById("carerSearchResults");
const carerSearchStatus = document.getElementById("carerSearchStatus");
const clientPostcodeInput = document.getElementById("clientPostcodeInput");
const clientSearchInput = document.getElementById("clientSearchInput");
const areaFilters = document.getElementById("areaFilters");
const clientSearchResults = document.getElementById("clientSearchResults");
const clientSearchStatus = document.getElementById("clientSearchStatus");
const addClientBtn = document.getElementById("addClientBtn");
const calculateRunBtn = document.getElementById("calculateRunBtn");
const clearRunBtn = document.getElementById("clearRunBtn");
const runDateInput = document.getElementById("runDateInput");
const runStartTimeInput = document.getElementById("runStartTimeInput");
const visitDurationInput = document.getElementById("visitDurationInput");
const runNameInput = document.getElementById("runNameInput");
const saveRunBtn = document.getElementById("saveRunBtn");
const exportRunsBtn = document.getElementById("exportRunsBtn");
const savedRunsStatus = document.getElementById("savedRunsStatus");
const savedRunsBody = document.getElementById("savedRunsBody");
const mappingStatus = document.getElementById("mappingStatus");
const clientPostcodesList = document.getElementById("clientPostcodesList");
const noClientsMessage = document.getElementById("noClientsMessage");
const runResults = document.getElementById("runResults");
const runTotals = document.getElementById("runTotals");
const runCostSummary = document.getElementById("runCostSummary");
const runCostBreakdown = document.getElementById("runCostBreakdown");
const runSchedule = document.getElementById("runSchedule");
const runOrderList = document.getElementById("runOrderList");
const runLegsBody = document.getElementById("runLegsBody");
const mapsDirectionsLink = document.getElementById("mapsDirectionsLink");

const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
const RUN_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/routes/run` : "/api/routes/run";

const selectedClientStops = [];
let allClients = [];
let allCarers = [];
let selectedArea = "ALL";
const FIXED_AREAS = ["Central", "London Plus", "East Kent"];
const DEFAULT_VISIT_DURATION_MINUTES = 60;
const SAVED_RUNS_KEY = "thrive.mapping.savedRuns.v1";
let scheduleRows = [];
let scheduleGapMinutes = [];
let lastRun = null;
let savedRuns = [];

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

function normalizePostcode(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function normalizeLocationQuery(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function compactLocationQuery(value) {
  return normalizeLocationQuery(value).toLowerCase();
}

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function setStatus(message, isError = false) {
  mappingStatus.textContent = message;
  mappingStatus.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "mapping").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function setBusy(isBusy) {
  calculateRunBtn.disabled = isBusy;
  addClientBtn.disabled = isBusy;
  clearRunBtn.disabled = isBusy;
  if (saveRunBtn) {
    saveRunBtn.disabled = isBusy || !lastRun;
  }
}

function setClientSearchStatus(message, isError = false) {
  clientSearchStatus.textContent = message;
  clientSearchStatus.classList.toggle("error", isError);
}

function setCarerSearchStatus(message, isError = false) {
  if (!carerSearchStatus) {
    return;
  }
  carerSearchStatus.textContent = message;
  carerSearchStatus.classList.toggle("error", isError);
}

function updateSavedRunsStatus() {
  if (!savedRunsStatus || !exportRunsBtn) {
    return;
  }
  const count = savedRuns.length;
  savedRunsStatus.textContent = count === 0 ? "No saved runs yet." : `${count} saved run(s) ready for export.`;
  exportRunsBtn.disabled = count === 0;
  renderSavedRunsTable();
}

function loadSavedRuns() {
  try {
    const raw = localStorage.getItem(SAVED_RUNS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    savedRuns = Array.isArray(parsed) ? parsed : [];
  } catch {
    savedRuns = [];
  }
  updateSavedRunsStatus();
}

function persistSavedRuns() {
  try {
    localStorage.setItem(SAVED_RUNS_KEY, JSON.stringify(savedRuns));
    updateSavedRunsStatus();
  } catch {
    setStatus("Could not persist saved runs locally.", true);
  }
}

function renderSavedRunsTable() {
  if (!savedRunsBody) {
    return;
  }

  savedRunsBody.innerHTML = "";
  if (!savedRuns.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="5" class="muted">No saved runs.</td>`;
    savedRunsBody.appendChild(tr);
    return;
  }

  for (const run of savedRuns) {
    const saved = new Date(run.savedAt);
    const savedText = Number.isNaN(saved.getTime()) ? String(run.savedAt || "") : saved.toLocaleString();
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(savedText)}</td>
      <td>${escapeHtml(String(run.name || ""))}</td>
      <td>${escapeHtml(String(run.runStartTime || ""))}</td>
      <td>${escapeHtml(String(run.clientStopCount || ""))}</td>
      <td>£${escapeHtml(String(run.grandTotal || "0.00"))}</td>
    `;
    savedRunsBody.appendChild(tr);
  }
}

function hideRun() {
  lastRun = null;
  scheduleRows = [];
  scheduleGapMinutes = [];
  runResults.hidden = true;
  runOrderList.innerHTML = "";
  runLegsBody.innerHTML = "";
  if (runSchedule) {
    runSchedule.innerHTML = "";
  }
  if (runCostSummary) {
    runCostSummary.textContent = "";
  }
  if (runCostBreakdown) {
    runCostBreakdown.innerHTML = "";
  }
  if (saveRunBtn) {
    saveRunBtn.disabled = true;
  }
}

function renderClientPostcodes() {
  clientPostcodesList.innerHTML = "";

  selectedClientStops.forEach((stop, index) => {
    const li = document.createElement("li");
    li.className = "postcode-pill";

    const text = document.createElement("span");
    text.textContent = stop.label ? `${stop.label} - ${stop.address}` : stop.address;

    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "pill-remove-btn";
    remove.textContent = "Remove";
    remove.addEventListener("click", () => {
      selectedClientStops.splice(index, 1);
      renderClientPostcodes();
      hideRun();
      setStatus("Client stop removed.");
    });

    const moveUp = document.createElement("button");
    moveUp.type = "button";
    moveUp.className = "pill-action-btn";
    moveUp.textContent = "Up";
    moveUp.disabled = index === 0;
    moveUp.addEventListener("click", () => {
      moveStop(index, index - 1);
    });

    const moveDown = document.createElement("button");
    moveDown.type = "button";
    moveDown.className = "pill-action-btn";
    moveDown.textContent = "Down";
    moveDown.disabled = index === selectedClientStops.length - 1;
    moveDown.addEventListener("click", () => {
      moveStop(index, index + 1);
    });

    li.append(text, moveUp, moveDown, remove);
    clientPostcodesList.appendChild(li);
  });

  noClientsMessage.hidden = selectedClientStops.length > 0;
}

function moveStop(fromIndex, toIndex) {
  if (fromIndex < 0 || fromIndex >= selectedClientStops.length) {
    return;
  }
  if (toIndex < 0 || toIndex >= selectedClientStops.length) {
    return;
  }
  const [moved] = selectedClientStops.splice(fromIndex, 1);
  selectedClientStops.splice(toIndex, 0, moved);
  renderClientPostcodes();
  hideRun();
  setStatus("Client stop order updated.");
}

function addClientStop(address, label = "") {
  const input = normalizeLocationQuery(address);
  if (!input || !compactLocationQuery(input)) {
    setStatus("Enter a client stop before adding.", true);
    return false;
  }

  selectedClientStops.push({
    label: normalizeLocationQuery(label),
    address: input,
  });
  renderClientPostcodes();
  hideRun();
  setStatus(label ? `Added ${label}.` : `Added ${input}.`);
  return true;
}

function addClientPostcodeFromInput() {
  if (addClientStop(clientPostcodeInput.value, "Manual")) {
    clientPostcodeInput.value = "";
  }
}

function getClientLabel(client) {
  return String(client.name || "Unnamed client").trim();
}

function getClientArea(client) {
  return normalizeLocationQuery(client.area || client.location || "");
}

function getFilteredClients() {
  const query = normalizeText(clientSearchInput?.value);
  return allClients.filter((client) => {
    const area = getClientArea(client) || "Unassigned";
    const matchesArea = selectedArea === "ALL" || area === selectedArea;
    if (!matchesArea) {
      return false;
    }

    if (!query) {
      return true;
    }

    return (
      normalizeText(client.name).includes(query) ||
      normalizeText(client.id).includes(query) ||
      normalizeText(getClientArea(client)).includes(query) ||
      normalizeText(client.postcode).includes(query) ||
      normalizeText(client.address).includes(query) ||
      normalizeText(client.town).includes(query) ||
      normalizeText(client.county).includes(query)
    );
  });
}

function buildClientAddress(client) {
  const parts = [client.address, client.town, client.county, client.postcode]
    .map((value) => normalizeLocationQuery(value))
    .filter(Boolean);

  if (parts.length) {
    return parts.join(", ");
  }

  return normalizeLocationQuery(client.postcode || client.location || "");
}

function getAreaOptions() {
  const normalizedKnown = new Map(FIXED_AREAS.map((area) => [area.toLowerCase(), area]));
  const found = new Set();

  for (const client of allClients) {
    const area = getClientArea(client);
    const known = normalizedKnown.get(area.toLowerCase());
    if (known) {
      found.add(known);
    }
  }

  const ordered = FIXED_AREAS.filter((area) => found.has(area));
  return ["ALL", ...ordered];
}

function renderAreaFilters() {
  if (!areaFilters) {
    return;
  }

  areaFilters.innerHTML = "";
  const options = getAreaOptions();
  if (!options.length) {
    return;
  }

  if (!options.includes(selectedArea)) {
    selectedArea = "ALL";
  }

  for (const area of options) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `area-filter-btn${selectedArea === area ? " active" : ""}`;
    btn.textContent = area === "ALL" ? "All Areas" : area;
    btn.addEventListener("click", () => {
      selectedArea = area;
      renderAreaFilters();
      renderClientSearchResults();
    });
    areaFilters.appendChild(btn);
  }
}

function getFilteredCarers() {
  const query = normalizeText(carerSearchInput?.value);
  if (!query) {
    return allCarers.slice(0, 12);
  }

  return allCarers
    .filter((carer) => {
      return (
        normalizeText(carer.name).includes(query) ||
        normalizeText(carer.id).includes(query) ||
        normalizeText(carer.postcode).includes(query)
      );
    })
    .slice(0, 25);
}

function renderCarerSearchResults() {
  if (!carerSearchResults) {
    return;
  }
  carerSearchResults.innerHTML = "";

  if (!allCarers.length) {
    setCarerSearchStatus("No carers loaded.");
    return;
  }

  const filtered = getFilteredCarers();
  if (!filtered.length) {
    setCarerSearchStatus("No matching carers.");
    return;
  }

  setCarerSearchStatus(`Showing ${filtered.length} carer(s).`);
  for (const carer of filtered) {
    const li = document.createElement("li");
    li.className = "carer-result-row";

    const info = document.createElement("div");
    info.className = "client-result-info";
    info.innerHTML = `
      <strong>${escapeHtml(String(carer.name || "Unnamed carer"))}</strong>
      <span>${escapeHtml(String(carer.id || "-"))}</span>
      <span>Postcode: ${escapeHtml(String(carer.postcode || "Not available"))}</span>
    `;

    const selectBtn = document.createElement("button");
    selectBtn.type = "button";
    selectBtn.className = "secondary";
    selectBtn.textContent = "Set staff";
    selectBtn.disabled = !normalizePostcode(carer.postcode);
    selectBtn.addEventListener("click", () => {
      const postcode = normalizePostcode(carer.postcode);
      if (!postcode) {
        setStatus("Selected carer has no postcode.", true);
        return;
      }
      staffPostcodeInput.value = postcode;
      setStatus(`Staff postcode set from ${carer.name || "carer"}.`);
    });

    li.append(info, selectBtn);
    carerSearchResults.appendChild(li);
  }
}

function renderClientSearchResults() {
  clientSearchResults.innerHTML = "";
  const filtered = getFilteredClients();

  if (!allClients.length) {
    setClientSearchStatus("No clients loaded.", true);
    return;
  }

  if (!filtered.length) {
    setClientSearchStatus("No matching clients.");
    return;
  }

  const visibleCount = Math.min(filtered.length, 3);
  setClientSearchStatus(`Showing ${visibleCount} of ${filtered.length} client(s). Scroll for more.`);

  for (const client of filtered) {
    const li = document.createElement("li");
    li.className = "client-result-row";

    const routeAddress = buildClientAddress(client);
    const info = document.createElement("div");
    info.className = "client-result-info";
    info.innerHTML = `
      <strong>${escapeHtml(String(client.name || "Unnamed client"))}</strong>
      <span>${escapeHtml(routeAddress || "Not available")}</span>
    `;

    const addBtn = document.createElement("button");
    addBtn.type = "button";
    addBtn.className = "secondary";
    addBtn.textContent = "Add client";
    addBtn.disabled = !routeAddress;
    addBtn.addEventListener("click", () => {
      addClientStop(routeAddress, getClientLabel(client));
    });

    li.append(info, addBtn);
    clientSearchResults.appendChild(li);
  }
}

function buildMapsDirectionsUrl(staffPostcode, orderedClients) {
  const url = new URL("https://www.google.com/maps/dir/");
  url.searchParams.set("api", "1");
  url.searchParams.set("travelmode", "driving");
  url.searchParams.set("origin", staffPostcode);
  url.searchParams.set("destination", staffPostcode);
  if (orderedClients.length) {
    url.searchParams.set("waypoints", orderedClients.join("|"));
  }
  return url.toString();
}

function getDurationMinutes(durationSeconds) {
  const seconds = Number(durationSeconds || 0);
  if (!Number.isFinite(seconds) || seconds <= 0) {
    return 0;
  }
  return Math.max(1, Math.round(seconds / 60));
}

function parseTimeInputToMinutes(value) {
  const match = /^(\d{2}):(\d{2})(?::(\d{2}))?$/.exec(String(value || "").trim());
  if (!match) {
    return null;
  }
  const hours = Number(match[1]);
  const mins = Number(match[2]);
  if (!Number.isFinite(hours) || !Number.isFinite(mins)) {
    return null;
  }
  return hours * 60 + mins;
}

function getVisitDurationMinutes() {
  const parsed = Number(visitDurationInput?.value);
  if (!Number.isFinite(parsed) || parsed < 0) {
    return DEFAULT_VISIT_DURATION_MINUTES;
  }
  return roundUpToNearestTenMinutes(parsed);
}

function toDateInputValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function getNextMondayDate() {
  const now = new Date();
  const date = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const day = date.getDay(); // 0=Sun ... 1=Mon
  let daysUntilMonday = (8 - day) % 7;
  if (daysUntilMonday === 0) {
    daysUntilMonday = 7;
  }
  date.setDate(date.getDate() + daysUntilMonday);
  return date;
}

function parseDateInputValue(value) {
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || "").trim());
  if (!match) {
    return null;
  }
  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || !Number.isFinite(day)) {
    return null;
  }
  const date = new Date(year, month, day);
  if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) {
    return null;
  }
  return date;
}

function formatMinutesForTimeInput(totalMinutes) {
  let mins = Number(totalMinutes || 0);
  if (!Number.isFinite(mins)) {
    mins = 0;
  }
  mins = ((mins % 1440) + 1440) % 1440;
  const hours = Math.floor(mins / 60);
  const remainder = mins % 60;
  return `${String(hours).padStart(2, "0")}:${String(remainder).padStart(2, "0")}`;
}

function roundUpToNearestTenMinutes(totalMinutes) {
  const mins = Number(totalMinutes || 0);
  if (!Number.isFinite(mins)) {
    return 540;
  }
  return Math.ceil(mins / 10) * 10;
}

function resolveStartMinutes() {
  const parsed = parseTimeInputToMinutes(runStartTimeInput?.value);
  if (parsed === null) {
    return 540;
  }
  return roundUpToNearestTenMinutes(parsed);
}

function buildDepartureTimeIso() {
  const startMinutes = resolveStartMinutes();
  const selectedDate = parseDateInputValue(runDateInput?.value) || getNextMondayDate();
  const local = new Date(
    selectedDate.getFullYear(),
    selectedDate.getMonth(),
    selectedDate.getDate(),
    Math.floor(startMinutes / 60),
    startMinutes % 60,
    0,
    0
  );
  return local.toISOString();
}

function buildScheduleRows(run) {
  const clients = Array.isArray(run.orderedClients) ? run.orderedClients : [];
  const legs = Array.isArray(run.legs) ? run.legs : [];
  const staffPostcode = run.staffStart?.postcode || "Staff home";

  const labels = [`Start - ${staffPostcode}`];
  for (const client of clients) {
    labels.push(client.stopLabel || client.formattedAddress || client.query || "Client");
  }
  labels.push(`Return - ${staffPostcode}`);

  const visitDuration = getVisitDurationMinutes();
  scheduleGapMinutes = [];
  for (let i = 0; i < labels.length - 1; i += 1) {
    const legMinutes = getDurationMinutes(legs[i]?.durationSeconds);
    if (i === 0) {
      scheduleGapMinutes.push(legMinutes);
    } else {
      scheduleGapMinutes.push(visitDuration + legMinutes);
    }
  }

  const startMinutes = resolveStartMinutes();
  scheduleRows = labels.map((label, index) => ({
    label,
    minutes: index === 0 ? startMinutes : 0,
  }));

  for (let i = 1; i < scheduleRows.length; i += 1) {
    scheduleRows[i].minutes = roundUpToNearestTenMinutes(scheduleRows[i - 1].minutes + scheduleGapMinutes[i - 1]);
  }
}

function closestDelta(oldMinutes, newClockMinutes) {
  const candidates = [newClockMinutes, newClockMinutes + 1440, newClockMinutes - 1440];
  let best = candidates[0] - oldMinutes;
  for (const candidate of candidates.slice(1)) {
    const delta = candidate - oldMinutes;
    if (Math.abs(delta) < Math.abs(best)) {
      best = delta;
    }
  }
  return best;
}

function renderSchedule() {
  if (!runSchedule) {
    return;
  }

  runSchedule.innerHTML = "";
  for (let i = 0; i < scheduleRows.length; i += 1) {
    const row = scheduleRows[i];
    const wrap = document.createElement("div");
    wrap.className = "schedule-row";

    const label = document.createElement("span");
    label.className = "schedule-label";
    label.textContent = row.label;

    const timeInput = document.createElement("input");
    timeInput.type = "time";
    timeInput.step = "600";
    timeInput.value = formatMinutesForTimeInput(row.minutes);
    timeInput.addEventListener("change", () => {
      const parsed = parseTimeInputToMinutes(timeInput.value);
      if (parsed === null) {
        timeInput.value = formatMinutesForTimeInput(scheduleRows[i].minutes);
        return;
      }

      const delta = closestDelta(scheduleRows[i].minutes, parsed);
      scheduleRows[i].minutes = roundUpToNearestTenMinutes(scheduleRows[i].minutes + delta);
      for (let rowIndex = i + 1; rowIndex < scheduleRows.length; rowIndex += 1) {
        scheduleRows[rowIndex].minutes = roundUpToNearestTenMinutes(
          scheduleRows[rowIndex - 1].minutes + scheduleGapMinutes[rowIndex - 1]
        );
      }

      if (i === 0 && runStartTimeInput) {
        runStartTimeInput.value = formatMinutesForTimeInput(scheduleRows[0].minutes);
      }

      renderSchedule();
    });

    wrap.append(label, timeInput);
    runSchedule.appendChild(wrap);
  }
}

function renderRun(run) {
  lastRun = run;
  const staffPostcode = run.staffStart?.postcode || "";
  const orderedClients = Array.isArray(run.orderedClients) ? run.orderedClients : [];
  const legs = Array.isArray(run.legs) ? run.legs : [];

  runTotals.textContent = `Total distance: ${run.totalDistanceMiles || 0} mi | Total time: ${run.totalDurationText || "0 min"}`;
  renderCost(run);

  runOrderList.innerHTML = "";
  const startItem = document.createElement("li");
  startItem.textContent = `Start: ${staffPostcode}`;
  runOrderList.appendChild(startItem);

  for (const client of orderedClients) {
    const li = document.createElement("li");
    li.textContent = client.stopLabel || client.formattedAddress || client.query || "";
    runOrderList.appendChild(li);
  }

  const endItem = document.createElement("li");
  endItem.textContent = `Return: ${staffPostcode}`;
  runOrderList.appendChild(endItem);

  runLegsBody.innerHTML = "";
  for (const leg of legs) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(leg.legNumber || ""))}</td>
      <td>${escapeHtml(String(leg.from || ""))}</td>
      <td>${escapeHtml(String(leg.to || ""))}</td>
      <td>${escapeHtml(String(leg.distanceMiles || 0))} mi</td>
      <td>${escapeHtml(String(leg.durationText || ""))}</td>
    `;
    runLegsBody.appendChild(tr);
  }

  mapsDirectionsLink.href = buildMapsDirectionsUrl(
    staffPostcode,
    orderedClients.map((item) => item.query || item.formattedAddress).filter(Boolean)
  );
  buildScheduleRows(run);
  renderSchedule();
  if (saveRunBtn) {
    saveRunBtn.disabled = false;
  }
  runResults.hidden = false;
}

function calculateTravelPerHourMetrics(run) {
  const cost = run?.cost || {};
  const visitCount = Array.isArray(run?.orderedClients) ? run.orderedClients.length : 0;
  const visitHours = (visitCount * getVisitDurationMinutes()) / 60;
  if (!Number.isFinite(visitHours) || visitHours <= 0) {
    return {
      visitHours: 0,
      exceptionalTravelPerHour: null,
      totalTravelPerHour: null,
    };
  }

  const exceptionalTotal = Number(cost.totals?.exceptionalHomeTotal || 0);
  const totalTravel = Number(cost.totals?.grandTotal || 0);

  return {
    visitHours,
    exceptionalTravelPerHour: exceptionalTotal / visitHours,
    totalTravelPerHour: totalTravel / visitHours,
  };
}

function renderCost(run) {
  const cost = run?.cost || null;
  if (!runCostSummary || !runCostBreakdown) {
    return;
  }

  if (!cost) {
    runCostSummary.textContent = "Costing unavailable.";
    runCostBreakdown.innerHTML = "";
    return;
  }

  const modeText = cost.mode === "time" ? "time threshold" : "distance threshold";
  const basedOnText =
    cost.mode === "time"
      ? `${Number(cost.thresholds?.maxTimeMinutes || 0).toFixed(0)} mins`
      : `${Number(cost.thresholds?.maxDistanceMiles || 0).toFixed(2)} miles`;
  const modeLabel = modeText.charAt(0).toUpperCase() + modeText.slice(1);
  runCostSummary.textContent = `Costing mode: ${modeLabel} ${basedOnText}.`;
  runCostBreakdown.innerHTML = "";

  const homeSeconds = Number(cost.homeTravel?.paidDurationSeconds || 0);
  const runSeconds = Number(cost.runTravel?.durationSeconds || 0);
  const metrics = calculateTravelPerHourMetrics(run);
  const visitHoursText = metrics.visitHours > 0 ? metrics.visitHours.toFixed(2) : "0.00";
  const exceptionalPerHourText =
    metrics.exceptionalTravelPerHour === null ? "n/a" : `£${metrics.exceptionalTravelPerHour.toFixed(2)}`;
  const totalPerHourText =
    metrics.totalTravelPerHour === null ? "n/a" : `£${metrics.totalTravelPerHour.toFixed(2)}`;
  runCostBreakdown.innerHTML = `
    <section class="cost-section">
      <h3>Exceptional Travel Costs from Home</h3>
      <p>Paid distance/time: ${Number(cost.homeTravel?.paidDistanceMiles || 0).toFixed(2)} mi, ${formatSecondsAsTime(homeSeconds)}</p>
      <p>Time cost: £${Number(cost.components?.homeTimeCost || 0).toFixed(2)}</p>
      <p>Mileage cost: £${Number(cost.components?.homeMileageCost || 0).toFixed(2)}</p>
      <p class="cost-total">Exceptional Travel Total: £${Number(cost.totals?.exceptionalHomeTotal || 0).toFixed(2)}</p>
      <p class="cost-metric">Exceptional Travel £/hour: ${exceptionalPerHourText}</p>
    </section>
    <section class="cost-section">
      <h3>Run Travel</h3>
      <p>Distance/time: ${Number(cost.runTravel?.distanceMiles || 0).toFixed(2)} mi, ${formatSecondsAsTime(runSeconds)}</p>
      <p>Time cost: £${Number(cost.components?.runTimeCost || 0).toFixed(2)}</p>
      <p>Mileage cost: £${Number(cost.components?.runMileageCost || 0).toFixed(2)}</p>
      <p class="cost-total">Run Total: £${Number(cost.totals?.runTravelTotal || 0).toFixed(2)}</p>
    </section>
    <section class="cost-section grand">
      <h3>Grand Total</h3>
      <p class="cost-total">£${Number(cost.totals?.grandTotal || 0).toFixed(2)}</p>
      <p>Total visit hours: ${visitHoursText} h</p>
      <p class="cost-metric">Total Travel £/hour: ${totalPerHourText}</p>
    </section>
  `;
}

function formatSecondsAsTime(totalSeconds) {
  const safeSeconds = Number(totalSeconds || 0);
  if (!Number.isFinite(safeSeconds) || safeSeconds <= 0) {
    return "0 mins";
  }

  const totalMinutes = Math.round(safeSeconds / 60);
  if (totalMinutes < 60) {
    return `${totalMinutes} mins`;
  }

  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  if (minutes === 0) {
    return `${hours}h`;
  }
  return `${hours}h ${minutes}m`;
}

function collectRunExportRecord(name) {
  const run = lastRun || {};
  const cost = run.cost || {};
  const clients = Array.isArray(run.orderedClients) ? run.orderedClients : [];
  const legCount = Array.isArray(run.legs) ? run.legs.length : 0;
  const startTime = scheduleRows.length ? formatMinutesForTimeInput(scheduleRows[0].minutes) : (runStartTimeInput?.value || "09:00");
  const scheduleText = scheduleRows
    .map((row) => `${row.label} @ ${formatMinutesForTimeInput(row.minutes)}`)
    .join(" | ");
  const stopsText = clients
    .map((client) => client.stopLabel || client.formattedAddress || client.query || "")
    .filter(Boolean)
    .join(" | ");

  const travelPerHour = calculateTravelPerHourMetrics(run);

  return {
    savedAt: new Date().toISOString(),
    name,
    runStartTime: startTime,
    staffStart: String(run.staffStart?.postcode || ""),
    clientStopCount: clients.length,
    clientStops: stopsText,
    legCount,
    totalDistanceMiles: Number(run.totalDistanceMiles || 0).toFixed(2),
    totalDuration: String(run.totalDurationText || ""),
    costingMode: String(runCostSummary?.textContent || ""),
    exceptionalTravelTotal: Number(cost.totals?.exceptionalHomeTotal || 0).toFixed(2),
    runTravelTotal: Number(cost.totals?.runTravelTotal || 0).toFixed(2),
    grandTotal: Number(cost.totals?.grandTotal || 0).toFixed(2),
    totalVisitHours: travelPerHour.visitHours.toFixed(2),
    exceptionalTravelPerHour:
      travelPerHour.exceptionalTravelPerHour === null ? "" : travelPerHour.exceptionalTravelPerHour.toFixed(2),
    totalTravelPerHour:
      travelPerHour.totalTravelPerHour === null ? "" : travelPerHour.totalTravelPerHour.toFixed(2),
    schedule: scheduleText,
  };
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

function exportSavedRunsAsCsv() {
  if (!savedRuns.length) {
    setStatus("No saved runs to export.", true);
    return;
  }

  const headers = [
    "Saved At",
    "Run Name",
    "Run Start Time",
    "Staff Start",
    "Client Stop Count",
    "Client Stops",
    "Leg Count",
    "Total Distance (mi)",
    "Total Duration",
    "Costing Mode",
    "Exceptional Travel Total",
    "Run Travel Total",
    "Grand Total",
    "Total Visit Hours",
    "Exceptional Travel £/hour",
    "Total Travel £/hour",
    "Schedule",
  ];

  const lines = [headers.join(",")];
  for (const row of savedRuns) {
    const values = [
      row.savedAt,
      row.name,
      row.runStartTime,
      row.staffStart,
      row.clientStopCount,
      row.clientStops,
      row.legCount,
      row.totalDistanceMiles,
      row.totalDuration,
      row.costingMode,
      row.exceptionalTravelTotal,
      row.runTravelTotal,
      row.grandTotal,
      row.totalVisitHours,
      row.exceptionalTravelPerHour,
      row.totalTravelPerHour,
      row.schedule,
    ];
    lines.push(values.map(escapeCsvValue).join(","));
  }

  const stamp = new Date().toISOString().slice(0, 19).replaceAll(":", "-");
  downloadCsv(`runs-export-${stamp}.csv`, lines.join("\n"));
  setStatus(`Exported ${savedRuns.length} run(s).`);
}

function saveCurrentRunForExport() {
  if (!lastRun) {
    setStatus("Calculate a run before saving.", true);
    return;
  }

  const runName = normalizeLocationQuery(runNameInput?.value);
  if (!runName) {
    setStatus("Enter a run name before saving.", true);
    return;
  }

  const record = collectRunExportRecord(runName);
  savedRuns.push(record);
  persistSavedRuns();
  if (runNameInput) {
    runNameInput.value = "";
  }
  setStatus(`Saved run '${runName}' for export.`);
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

async function calculateRun() {
  const staffPostcode = normalizePostcode(staffPostcodeInput.value);
  if (!staffPostcode) {
    setStatus("Staff postcode is required.", true);
    return;
  }

  if (!selectedClientStops.length) {
    setStatus("Add at least one client stop.", true);
    return;
  }

  setBusy(true);
  hideRun();
  setStatus("Calculating route...");

  try {
    const token = await authController.acquireToken([FRONTEND_CONFIG.apiScope]);
    const response = await fetch(RUN_ENDPOINT, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        staffPostcode,
        departureTime: buildDepartureTimeIso(),
        clientLocations: selectedClientStops.map((stop) => ({
          query: stop.address,
          label: stop.label || stop.address,
        })),
      }),
    });

    const data = await response.json();
    if (!response.ok) {
      if (response.status === 403) {
        redirectToUnauthorized("mapping");
        return;
      }
      throw new Error(data?.detail || data?.error || `Run request failed (${response.status}).`);
    }

    renderRun(data.run || {});
    setStatus(`Run calculated with ${selectedClientStops.length} client stop(s).`);
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not calculate run.", true);
  } finally {
    setBusy(false);
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
    if (!canAccessPage(role, "mapping")) {
      redirectToUnauthorized("mapping");
      return;
    }
    renderTopNavigation({ role });

    if (runDateInput && !runDateInput.value) {
      runDateInput.value = toDateInputValue(getNextMondayDate());
    }
    loadSavedRuns();
    renderClientPostcodes();
    setCarerSearchStatus("Loading carers...");
    setClientSearchStatus("Loading clients...");
    const [clientsPayload, carersPayload] = await Promise.all([
      directoryApi.listClients({ limit: 1000 }),
      directoryApi.listCarers({ limit: 500 }),
    ]);
    allClients = Array.isArray(clientsPayload?.clients) ? clientsPayload.clients : [];
    allCarers = Array.isArray(carersPayload?.carers) ? carersPayload.carers : [];
    renderCarerSearchResults();
    renderAreaFilters();
    renderClientSearchResults();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("mapping");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
    setCarerSearchStatus(error?.message || "Could not load carers.", true);
    setClientSearchStatus(error?.message || "Could not load clients.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

addClientBtn?.addEventListener("click", () => {
  addClientPostcodeFromInput();
});

clientPostcodeInput?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    addClientPostcodeFromInput();
  }
});

clientSearchInput?.addEventListener("input", () => {
  renderClientSearchResults();
});

carerSearchInput?.addEventListener("input", () => {
  renderCarerSearchResults();
});

runNameInput?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    saveCurrentRunForExport();
  }
});

calculateRunBtn?.addEventListener("click", () => {
  void calculateRun();
});

saveRunBtn?.addEventListener("click", () => {
  saveCurrentRunForExport();
});

exportRunsBtn?.addEventListener("click", () => {
  exportSavedRunsAsCsv();
});

clearRunBtn?.addEventListener("click", () => {
  if (runDateInput) {
    runDateInput.value = toDateInputValue(getNextMondayDate());
  }
  if (visitDurationInput) {
    visitDurationInput.value = String(DEFAULT_VISIT_DURATION_MINUTES);
  }
  staffPostcodeInput.value = "";
  clientPostcodeInput.value = "";
  if (carerSearchInput) {
    carerSearchInput.value = "";
  }
  if (clientSearchInput) {
    clientSearchInput.value = "";
  }
  selectedArea = "ALL";
  selectedClientStops.length = 0;
  renderCarerSearchResults();
  renderClientPostcodes();
  renderAreaFilters();
  renderClientSearchResults();
  hideRun();
  setStatus("Cleared.");
});

runStartTimeInput?.addEventListener("change", () => {
  const parsed = parseTimeInputToMinutes(runStartTimeInput.value);
  if (parsed === null) {
    runStartTimeInput.value = "09:00";
    return;
  }
  runStartTimeInput.value = formatMinutesForTimeInput(roundUpToNearestTenMinutes(parsed));
  if (lastRun) {
    buildScheduleRows(lastRun);
    renderSchedule();
  }
});

visitDurationInput?.addEventListener("change", () => {
  const rounded = getVisitDurationMinutes();
  if (visitDurationInput) {
    visitDurationInput.value = String(rounded);
  }
  if (lastRun) {
    buildScheduleRows(lastRun);
    renderSchedule();
  }
});

runDateInput?.addEventListener("change", () => {
  const parsed = parseDateInputValue(runDateInput.value);
  if (!parsed) {
    runDateInput.value = toDateInputValue(getNextMondayDate());
  }
});

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

void init();
