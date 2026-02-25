import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";

const signOutBtn = document.getElementById("signOutBtn");
const staffPostcodeInput = document.getElementById("staffPostcodeInput");
const clientPostcodeInput = document.getElementById("clientPostcodeInput");
const clientSearchInput = document.getElementById("clientSearchInput");
const areaFilters = document.getElementById("areaFilters");
const clientSearchResults = document.getElementById("clientSearchResults");
const clientSearchStatus = document.getElementById("clientSearchStatus");
const addClientBtn = document.getElementById("addClientBtn");
const calculateRunBtn = document.getElementById("calculateRunBtn");
const clearRunBtn = document.getElementById("clearRunBtn");
const runStartTimeInput = document.getElementById("runStartTimeInput");
const runNameInput = document.getElementById("runNameInput");
const saveRunBtn = document.getElementById("saveRunBtn");
const exportRunsBtn = document.getElementById("exportRunsBtn");
const savedRunsStatus = document.getElementById("savedRunsStatus");
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
const CLIENTS_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/clients` : "/api/clients";
const ME_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/auth/me` : "/api/auth/me";

const selectedClientStops = [];
let allClients = [];
let selectedArea = "ALL";
const FIXED_AREAS = ["Central", "London Plus", "East Kent"];
const VISIT_DURATION_MINUTES = 60;
const SAVED_RUNS_KEY = "thrive.mapping.savedRuns.v1";
let scheduleRows = [];
let scheduleGapMinutes = [];
let lastRun = null;
let savedRuns = [];

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});

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

function updateSavedRunsStatus() {
  if (!savedRunsStatus || !exportRunsBtn) {
    return;
  }
  const count = savedRuns.length;
  savedRunsStatus.textContent = count === 0 ? "No saved runs yet." : `${count} saved run(s) ready for export.`;
  exportRunsBtn.disabled = count === 0;
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

  const key = compactLocationQuery(input);
  const exists = selectedClientStops.some((stop) => compactLocationQuery(stop.address) === key);
  if (exists) {
    setStatus("Client stop already added.", true);
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
  const match = /^(\d{2}):(\d{2})$/.exec(String(value || "").trim());
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

function roundToNearestTenMinutes(totalMinutes) {
  const mins = Number(totalMinutes || 0);
  if (!Number.isFinite(mins)) {
    return 540;
  }
  return Math.round(mins / 10) * 10;
}

function resolveStartMinutes() {
  const parsed = parseTimeInputToMinutes(runStartTimeInput?.value);
  if (parsed === null) {
    return 540;
  }
  return roundToNearestTenMinutes(parsed);
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

  scheduleGapMinutes = [];
  for (let i = 0; i < labels.length - 1; i += 1) {
    const legMinutes = getDurationMinutes(legs[i]?.durationSeconds);
    if (i === 0) {
      scheduleGapMinutes.push(legMinutes);
    } else {
      scheduleGapMinutes.push(VISIT_DURATION_MINUTES + legMinutes);
    }
  }

  const startMinutes = resolveStartMinutes();
  scheduleRows = labels.map((label, index) => ({
    label,
    minutes: index === 0 ? startMinutes : 0,
  }));

  for (let i = 1; i < scheduleRows.length; i += 1) {
    scheduleRows[i].minutes = scheduleRows[i - 1].minutes + scheduleGapMinutes[i - 1];
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
    timeInput.step = "60";
    timeInput.value = formatMinutesForTimeInput(row.minutes);
    timeInput.addEventListener("change", () => {
      const parsed = parseTimeInputToMinutes(timeInput.value);
      if (parsed === null) {
        timeInput.value = formatMinutesForTimeInput(scheduleRows[i].minutes);
        return;
      }

      const delta = closestDelta(scheduleRows[i].minutes, parsed);
      for (let rowIndex = i; rowIndex < scheduleRows.length; rowIndex += 1) {
        scheduleRows[rowIndex].minutes += delta;
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
  renderCost(run.cost || null);

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

function renderCost(cost) {
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
  runCostBreakdown.innerHTML = `
    <section class="cost-section">
      <h3>Exceptional Travel Costs from Home</h3>
      <p>Paid distance/time: ${Number(cost.homeTravel?.paidDistanceMiles || 0).toFixed(2)} mi, ${formatSecondsAsTime(homeSeconds)}</p>
      <p>Time cost: £${Number(cost.components?.homeTimeCost || 0).toFixed(2)}</p>
      <p>Mileage cost: £${Number(cost.components?.homeMileageCost || 0).toFixed(2)}</p>
      <p class="cost-total">Exceptional Travel Total: £${Number(cost.totals?.exceptionalHomeTotal || 0).toFixed(2)}</p>
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
        clientLocations: selectedClientStops.map((stop) => ({
          query: stop.address,
          label: stop.label || stop.address,
        })),
      }),
    });

    const data = await response.json();
    if (!response.ok) {
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

async function fetchClients() {
  const token = await authController.acquireToken([FRONTEND_CONFIG.apiScope]);
  const response = await fetch(CLIENTS_ENDPOINT, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Clients request failed (${response.status}): ${text || "Unknown error"}`);
  }

  const data = await response.json();
  return Array.isArray(data?.clients) ? data.clients : [];
}

async function fetchCurrentUser() {
  const token = await authController.acquireToken([FRONTEND_CONFIG.apiScope]);
  const response = await fetch(ME_ENDPOINT, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Profile request failed (${response.status}): ${text || "Unknown error"}`);
  }

  return response.json();
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
    if (role === "marketing") {
      window.location.href = "./marketing.html";
      return;
    }

    loadSavedRuns();
    renderClientPostcodes();
    setClientSearchStatus("Loading clients...");
    allClients = await fetchClients();
    renderAreaFilters();
    renderClientSearchResults();
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
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
  staffPostcodeInput.value = "";
  clientPostcodeInput.value = "";
  if (clientSearchInput) {
    clientSearchInput.value = "";
  }
  selectedArea = "ALL";
  selectedClientStops.length = 0;
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
  runStartTimeInput.value = formatMinutesForTimeInput(roundToNearestTenMinutes(parsed));
  if (lastRun) {
    buildScheduleRows(lastRun);
    renderSchedule();
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
