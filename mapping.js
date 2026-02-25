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
const mappingStatus = document.getElementById("mappingStatus");
const clientPostcodesList = document.getElementById("clientPostcodesList");
const noClientsMessage = document.getElementById("noClientsMessage");
const runResults = document.getElementById("runResults");
const runTotals = document.getElementById("runTotals");
const runCostSummary = document.getElementById("runCostSummary");
const runCostBreakdown = document.getElementById("runCostBreakdown");
const runOrderList = document.getElementById("runOrderList");
const runLegsBody = document.getElementById("runLegsBody");
const mapsDirectionsLink = document.getElementById("mapsDirectionsLink");

const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
const RUN_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/routes/run` : "/api/routes/run";
const CLIENTS_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/clients` : "/api/clients";

const selectedClientStops = [];
let allClients = [];
let selectedArea = "ALL";
const FIXED_AREAS = ["Central", "London Plus", "East Kent"];

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
}

function setClientSearchStatus(message, isError = false) {
  clientSearchStatus.textContent = message;
  clientSearchStatus.classList.toggle("error", isError);
}

function hideRun() {
  runResults.hidden = true;
  runOrderList.innerHTML = "";
  runLegsBody.innerHTML = "";
  if (runCostSummary) {
    runCostSummary.textContent = "";
  }
  if (runCostBreakdown) {
    runCostBreakdown.innerHTML = "";
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

    li.append(text, remove);
    clientPostcodesList.appendChild(li);
  });

  noClientsMessage.hidden = selectedClientStops.length > 0;
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

function renderRun(run) {
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
  const distanceBasis = Number(cost.thresholds?.maxDistanceMiles);
  const timeBasis = Number(cost.thresholds?.maxTimeMinutes);
  const distanceText = Number.isFinite(distanceBasis) ? `${distanceBasis.toFixed(2)} miles` : "not set";
  const timeText = Number.isFinite(timeBasis) ? `${timeBasis.toFixed(0)} mins` : "not set";
  runCostSummary.textContent = `Costing mode: ${modeText}. Calculation based on ${basedOnText}. Basis: MAX_DISTANCE ${distanceText}, MAX_TIME ${timeText}.`;
  runCostBreakdown.innerHTML = "";

  const homeSeconds = Number(cost.homeTravel?.paidDurationSeconds || 0);
  const runSeconds = Number(cost.runTravel?.durationSeconds || 0);
  runCostBreakdown.innerHTML = `
    <section class="cost-section">
      <h3>Exceptional Travel Costs from Home</h3>
      <p>Paid distance/time: ${Number(cost.homeTravel?.paidDistanceMiles || 0).toFixed(2)} mi, ${formatSecondsAsTime(homeSeconds)}</p>
      <p>Time cost: £${Number(cost.components?.homeTimeCost || 0).toFixed(2)}</p>
      <p>Mileage cost: £${Number(cost.components?.homeMileageCost || 0).toFixed(2)}</p>
      <p class="cost-total">Section total: £${Number(cost.totals?.exceptionalHomeTotal || 0).toFixed(2)}</p>
    </section>
    <section class="cost-section">
      <h3>Run Travel</h3>
      <p>Distance/time: ${Number(cost.runTravel?.distanceMiles || 0).toFixed(2)} mi, ${formatSecondsAsTime(runSeconds)}</p>
      <p>Time cost: £${Number(cost.components?.runTimeCost || 0).toFixed(2)}</p>
      <p>Mileage cost: £${Number(cost.components?.runMileageCost || 0).toFixed(2)}</p>
      <p class="cost-total">Section total: £${Number(cost.totals?.runTravelTotal || 0).toFixed(2)}</p>
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

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

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

calculateRunBtn?.addEventListener("click", () => {
  void calculateRun();
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

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

void init();
