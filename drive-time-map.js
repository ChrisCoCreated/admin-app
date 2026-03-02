import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const signOutBtn = document.getElementById("signOutBtn");
const locationInput = document.getElementById("locationInput");
const searchNameInput = document.getElementById("searchNameInput");
const drawDriveTimeBtn = document.getElementById("drawDriveTimeBtn");
const saveSearchBtn = document.getElementById("saveSearchBtn");
const showAllSearchesBtn = document.getElementById("showAllSearchesBtn");
const hideAllSearchesBtn = document.getElementById("hideAllSearchesBtn");
const savedSearchesList = document.getElementById("savedSearchesList");
const noSavedSearchesMessage = document.getElementById("noSavedSearchesMessage");
const driveTimeStatus = document.getElementById("driveTimeStatus");
const driveTimeMeta = document.getElementById("driveTimeMeta");
const driveTimeMapRoot = document.getElementById("driveTimeMap");

const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
const DRIVE_TIME_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/maps/drive-time` : "/api/maps/drive-time";
const SAVED_SEARCHES_KEY = "thrive.drivetime.saved.v1";
const AREA_COLORS = [
  { stroke: "#1f3c88", fill: "#31b7c8" },
  { stroke: "#8b2f6e", fill: "#c9439b" },
  { stroke: "#1a7f65", fill: "#49cfbf" },
  { stroke: "#6a3d1a", fill: "#e5a15f" },
  { stroke: "#4a4f9a", fill: "#8ea1ff" },
  { stroke: "#2f6b2a", fill: "#89c46b" },
];

let map = null;
let previewMarker = null;
let previewPolygon = null;
let currentResult = null;
let savedSearches = [];
const areaLayers = new Map();

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

function setStatus(message, isError = false) {
  if (!driveTimeStatus) {
    return;
  }
  driveTimeStatus.textContent = message;
  driveTimeStatus.classList.toggle("error", isError);
}

function setBusy(isBusy) {
  if (drawDriveTimeBtn) {
    drawDriveTimeBtn.disabled = isBusy;
  }
  if (saveSearchBtn) {
    saveSearchBtn.disabled = isBusy || !currentResult;
  }
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "drivetime").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

async function authedFetch(url, init = {}) {
  const token = await authController.acquireToken([FRONTEND_CONFIG.apiScope]);
  return fetch(url, {
    ...init,
    headers: {
      ...(init.headers || {}),
      Authorization: `Bearer ${token}`,
    },
  });
}

function normalizeSearchName(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function createSearchId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }
  return `search-${Date.now()}-${Math.random().toString(16).slice(2, 10)}`;
}

function getColorByIndex(index) {
  return AREA_COLORS[index % AREA_COLORS.length];
}

function sanitizePolygon(points) {
  if (!Array.isArray(points)) {
    return [];
  }
  return points
    .map((point) => ({
      lat: Number(point?.lat),
      lng: Number(point?.lng),
    }))
    .filter((point) => Number.isFinite(point.lat) && Number.isFinite(point.lng));
}

function sanitizeSavedSearch(raw, fallbackIndex = 0) {
  const polygon = sanitizePolygon(raw?.polygon);
  const centerLat = Number(raw?.center?.lat);
  const centerLng = Number(raw?.center?.lng);
  if (!polygon.length || !Number.isFinite(centerLat) || !Number.isFinite(centerLng)) {
    return null;
  }

  const colorIndex = Number(raw?.colorIndex);
  return {
    id: String(raw?.id || createSearchId()),
    name: normalizeSearchName(raw?.name || raw?.formattedAddress || `Saved area ${fallbackIndex + 1}`),
    query: normalizeSearchName(raw?.query || ""),
    formattedAddress: normalizeSearchName(raw?.formattedAddress || ""),
    minutes: Number(raw?.minutes || 20),
    center: { lat: centerLat, lng: centerLng },
    polygon,
    quality: raw?.quality || null,
    createdAt: String(raw?.createdAt || new Date().toISOString()),
    visible: raw?.visible !== false,
    colorIndex: Number.isInteger(colorIndex) && colorIndex >= 0 ? colorIndex : fallbackIndex,
  };
}

function loadSavedSearches() {
  try {
    const raw = localStorage.getItem(SAVED_SEARCHES_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    const items = Array.isArray(parsed) ? parsed : [];
    savedSearches = items
      .map((entry, index) => sanitizeSavedSearch(entry, index))
      .filter(Boolean)
      .slice(0, 100);
  } catch {
    savedSearches = [];
  }
}

function persistSavedSearches() {
  try {
    localStorage.setItem(SAVED_SEARCHES_KEY, JSON.stringify(savedSearches));
  } catch {
    setStatus("Could not save searches locally.", true);
  }
}

function initMap() {
  if (!driveTimeMapRoot || !window.L) {
    return;
  }

  map = window.L.map(driveTimeMapRoot, {
    zoomControl: true,
    attributionControl: true,
  }).setView([51.5072, -0.1276], 10);

  window.L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
    maxZoom: 19,
    attribution:
      '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> &copy; <a href="https://carto.com/attributions">CARTO</a>',
  }).addTo(map);
}

function formatQualityText(quality) {
  const successfulDirections = Number(quality?.successfulDirections || 0);
  const sampledDirections = Number(quality?.sampledDirections || 0);
  const minMinutes = quality?.minDurationMinutes;
  const maxMinutes = quality?.maxDurationMinutes;
  const durationRange =
    Number.isFinite(minMinutes) && Number.isFinite(maxMinutes)
      ? `sampled drive times: ${minMinutes}-${maxMinutes} mins`
      : "sampled drive times unavailable";
  return `${successfulDirections}/${sampledDirections} directions resolved, ${durationRange}.`;
}

function renderPreview(payload) {
  if (!map || !window.L) {
    throw new Error("Map is not available.");
  }

  const center = payload?.center;
  const polygonPoints = Array.isArray(payload?.polygon) ? payload.polygon : [];
  if (!center || typeof center.lat !== "number" || typeof center.lng !== "number" || !polygonPoints.length) {
    throw new Error("Drive-time response was missing map coordinates.");
  }

  if (previewMarker) {
    previewMarker.remove();
  }
  if (previewPolygon) {
    previewPolygon.remove();
  }

  previewMarker = window.L.marker([center.lat, center.lng], {
    title: payload?.formattedAddress || "Selected location",
  }).addTo(map);

  previewPolygon = window.L.polygon(
    polygonPoints.map((point) => [point.lat, point.lng]),
    {
      color: "#1d2a42",
      weight: 2,
      opacity: 0.85,
      dashArray: "5 5",
      fillColor: "#9db3df",
      fillOpacity: 0.2,
    }
  ).addTo(map);

  const polygonBounds = previewPolygon.getBounds();
  if (polygonBounds.isValid()) {
    map.fitBounds(polygonBounds.pad(0.15));
  } else {
    map.setView([center.lat, center.lng], 11);
  }
}

function removeAreaLayers(searchId) {
  const layers = areaLayers.get(searchId);
  if (!layers) {
    return;
  }
  if (layers.marker) {
    layers.marker.remove();
  }
  if (layers.polygon) {
    layers.polygon.remove();
  }
  areaLayers.delete(searchId);
}

function syncSearchLayer(search) {
  removeAreaLayers(search.id);
  if (!search.visible || !map || !window.L) {
    return;
  }

  const color = getColorByIndex(search.colorIndex);
  const marker = window.L.circleMarker([search.center.lat, search.center.lng], {
    radius: 6,
    color: color.stroke,
    weight: 2,
    fillColor: "#fff",
    fillOpacity: 1,
  }).addTo(map);
  marker.bindPopup(search.name);

  const polygon = window.L.polygon(
    search.polygon.map((point) => [point.lat, point.lng]),
    {
      color: color.stroke,
      weight: 2,
      opacity: 0.95,
      fillColor: color.fill,
      fillOpacity: 0.2,
    }
  ).addTo(map);
  polygon.bindPopup(`${search.name} (${search.minutes} mins)`);
  areaLayers.set(search.id, { marker, polygon });
}

function syncAllSearchLayers() {
  for (const search of savedSearches) {
    syncSearchLayer(search);
  }
}

function fitToSearch(search) {
  const layers = areaLayers.get(search.id);
  if (layers?.polygon) {
    const bounds = layers.polygon.getBounds();
    if (bounds.isValid()) {
      map.fitBounds(bounds.pad(0.15));
      return;
    }
  }
  map.setView([search.center.lat, search.center.lng], 11);
}

function renderSavedSearches() {
  if (!savedSearchesList || !noSavedSearchesMessage) {
    return;
  }

  savedSearchesList.innerHTML = "";
  for (const search of savedSearches) {
    const color = getColorByIndex(search.colorIndex);
    const li = document.createElement("li");
    li.className = "drive-time-search-item";

    const visibilityLabel = document.createElement("label");
    visibilityLabel.className = "drive-time-search-visibility";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = search.visible;
    checkbox.addEventListener("change", () => {
      search.visible = checkbox.checked;
      persistSavedSearches();
      syncSearchLayer(search);
      setStatus(`${search.name} ${search.visible ? "shown" : "hidden"} on map.`);
    });

    const colorDot = document.createElement("span");
    colorDot.className = "drive-time-color-dot";
    colorDot.style.backgroundColor = color.fill;
    colorDot.style.borderColor = color.stroke;

    const titleWrap = document.createElement("div");
    titleWrap.className = "drive-time-search-title-wrap";
    const title = document.createElement("span");
    title.className = "drive-time-search-title";
    title.textContent = search.name;
    const subtitle = document.createElement("span");
    subtitle.className = "drive-time-search-subtitle";
    subtitle.textContent = search.formattedAddress || search.query || "Unknown address";
    titleWrap.append(title, subtitle);

    visibilityLabel.append(checkbox, colorDot, titleWrap);

    const actions = document.createElement("div");
    actions.className = "drive-time-search-actions";

    const focusBtn = document.createElement("button");
    focusBtn.type = "button";
    focusBtn.className = "secondary";
    focusBtn.textContent = "Focus";
    focusBtn.addEventListener("click", () => {
      if (!search.visible) {
        search.visible = true;
        persistSavedSearches();
        syncSearchLayer(search);
        renderSavedSearches();
      }
      fitToSearch(search);
    });

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "secondary";
    removeBtn.textContent = "Remove";
    removeBtn.addEventListener("click", () => {
      removeAreaLayers(search.id);
      savedSearches = savedSearches.filter((item) => item.id !== search.id);
      persistSavedSearches();
      renderSavedSearches();
      setStatus(`Removed '${search.name}'.`);
    });

    actions.append(focusBtn, removeBtn);
    li.append(visibilityLabel, actions);
    savedSearchesList.appendChild(li);
  }

  noSavedSearchesMessage.hidden = savedSearches.length > 0;
}

function saveCurrentSearch() {
  if (!currentResult) {
    setStatus("Draw an area first, then save it.", true);
    return;
  }

  const explicitName = normalizeSearchName(searchNameInput?.value || "");
  const fallbackName = normalizeSearchName(currentResult.formattedAddress || currentResult.query || "Saved area");
  const area = sanitizeSavedSearch(
    {
      id: createSearchId(),
      name: explicitName || fallbackName,
      query: currentResult.query,
      formattedAddress: currentResult.formattedAddress,
      minutes: currentResult.minutes,
      center: currentResult.center,
      polygon: currentResult.polygon,
      quality: currentResult.quality,
      createdAt: new Date().toISOString(),
      visible: true,
      colorIndex: savedSearches.length,
    },
    savedSearches.length
  );
  if (!area) {
    setStatus("Current result cannot be saved.", true);
    return;
  }

  savedSearches.unshift(area);
  persistSavedSearches();
  syncSearchLayer(area);
  renderSavedSearches();
  if (previewMarker) {
    previewMarker.remove();
    previewMarker = null;
  }
  if (previewPolygon) {
    previewPolygon.remove();
    previewPolygon = null;
  }

  if (searchNameInput) {
    searchNameInput.value = "";
  }
  if (saveSearchBtn) {
    saveSearchBtn.disabled = false;
  }
  setStatus(`Saved '${area.name}'.`);
}

function setAllSearchVisibility(isVisible) {
  if (!savedSearches.length) {
    return;
  }
  for (const search of savedSearches) {
    search.visible = isVisible;
    syncSearchLayer(search);
  }
  persistSavedSearches();
  renderSavedSearches();
  setStatus(isVisible ? "Showing all saved searches." : "Hid all saved searches.");
}

async function drawDriveTimeArea() {
  const location = String(locationInput?.value || "").trim();
  if (!location) {
    setStatus("Enter a location first.", true);
    return;
  }

  setBusy(true);
  setStatus("Calculating 20-minute drive-time area...");
  if (driveTimeMeta) {
    driveTimeMeta.textContent = "";
  }

  try {
    const response = await authedFetch(DRIVE_TIME_ENDPOINT, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        location,
        minutes: 20,
      }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data?.error || "Could not calculate drive-time area.");
    }

    currentResult = {
      query: location,
      formattedAddress: String(data.formattedAddress || location),
      minutes: Number(data.minutes || 20),
      center: data.center,
      polygon: Array.isArray(data.polygon) ? data.polygon : [],
      quality: data.quality || null,
    };
    renderPreview(data);
    setStatus(`Showing 20-minute drive-time area for ${data.formattedAddress || location}.`);
    if (driveTimeMeta) {
      driveTimeMeta.textContent = formatQualityText(data.quality);
    }
    if (saveSearchBtn) {
      saveSearchBtn.disabled = false;
    }
  } catch (error) {
    console.error(error);
    currentResult = null;
    if (saveSearchBtn) {
      saveSearchBtn.disabled = true;
    }
    setStatus(error?.message || "Could not draw drive-time area.", true);
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
    if (!canAccessPage(role, "drivetime")) {
      redirectToUnauthorized("drivetime");
      return;
    }
    renderTopNavigation({ role, currentPathname: window.location.pathname });

    initMap();
    if (!map) {
      throw new Error("Could not initialize map renderer.");
    }
    loadSavedSearches();
    renderSavedSearches();
    syncAllSearchLayers();
    setStatus("Enter a location to calculate.");
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("drivetime");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize map page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

drawDriveTimeBtn?.addEventListener("click", () => {
  void drawDriveTimeArea();
});

saveSearchBtn?.addEventListener("click", () => {
  saveCurrentSearch();
});

showAllSearchesBtn?.addEventListener("click", () => {
  setAllSearchVisibility(true);
});

hideAllSearchesBtn?.addEventListener("click", () => {
  setAllSearchVisibility(false);
});

locationInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") {
    return;
  }
  event.preventDefault();
  void drawDriveTimeArea();
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
