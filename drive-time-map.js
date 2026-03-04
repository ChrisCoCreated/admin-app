import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const signOutBtn = document.getElementById("signOutBtn");
const locationInput = document.getElementById("locationInput");
const driveTimeMinutesInput = document.getElementById("driveTimeMinutesInput");
const customDepartureBtn = document.getElementById("customDepartureBtn");
const customDepartureWrap = document.getElementById("customDepartureWrap");
const departureDateInput = document.getElementById("departureDateInput");
const departureTimeInput = document.getElementById("departureTimeInput");
const useDefaultDepartureBtn = document.getElementById("useDefaultDepartureBtn");
const searchNameInput = document.getElementById("searchNameInput");
const drawDriveTimeBtn = document.getElementById("drawDriveTimeBtn");
const saveSearchBtn = document.getElementById("saveSearchBtn");
const showAllSearchesBtn = document.getElementById("showAllSearchesBtn");
const hideAllSearchesBtn = document.getElementById("hideAllSearchesBtn");
const exportSearchesBtn = document.getElementById("exportSearchesBtn");
const importSearchesBtn = document.getElementById("importSearchesBtn");
const importSearchesFileInput = document.getElementById("importSearchesFileInput");
const loadPeopleOverlayBtn = document.getElementById("loadPeopleOverlayBtn");
const showPeopleOverlayInput = document.getElementById("showPeopleOverlayInput");
const overlayTypeSelect = document.getElementById("overlayTypeSelect");
const overlayLocationSelect = document.getElementById("overlayLocationSelect");
const overlayStatusFilters = document.getElementById("overlayStatusFilters");
const peopleOverlayStatus = document.getElementById("peopleOverlayStatus");
const savedSearchesList = document.getElementById("savedSearchesList");
const noSavedSearchesMessage = document.getElementById("noSavedSearchesMessage");
const driveTimeStatus = document.getElementById("driveTimeStatus");
const driveTimeMeta = document.getElementById("driveTimeMeta");
const driveTimeMapRoot = document.getElementById("driveTimeMap");

const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
const DRIVE_TIME_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/maps/drive-time` : "/api/maps/drive-time";
const GEOCODE_BATCH_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/maps/geocode-batch` : "/api/maps/geocode-batch";
const SAVED_SEARCHES_KEY = "thrive.drivetime.saved.v1";
const GEOCODE_CACHE_KEY = "thrive.drivetime.geocode.v1";
const DEFAULT_DRIVE_TIME_MINUTES = 20;
const MIN_DRIVE_TIME_MINUTES = 1;
const MAX_DRIVE_TIME_MINUTES = 240;
const DEPARTURE_LEAD_MS = 120 * 1000;
const DEFAULT_DEPARTURE_WEEKDAY = 3; // Wednesday
const DEFAULT_DEPARTURE_HOUR = 10;
const OVERLAY_BATCH_SIZE = 300;
const MIN_POLYGON_POINTS = 3;
const AREA_COLORS = [
  { stroke: "#1f3c88", fill: "#31b7c8" },
  { stroke: "#8b2f6e", fill: "#c9439b" },
  { stroke: "#1a7f65", fill: "#49cfbf" },
  { stroke: "#6a3d1a", fill: "#e5a15f" },
  { stroke: "#4a4f9a", fill: "#8ea1ff" },
  { stroke: "#2f6b2a", fill: "#89c46b" },
];
const AREA_COLOR_LABELS = ["Brand Cyan", "Brand Pink", "Brand Mint", "Warm Sand", "Soft Indigo", "Leaf Green"];

let map = null;
let previewMarker = null;
let previewPolygon = null;
let previewVertexMarkers = [];
let currentResult = null;
let editingSearchId = null;
let savedSearches = [];
const areaLayers = new Map();
let peopleOverlayData = [];
const peopleOverlayLayers = new Map();
let overlayStatusSet = new Set();
let overlayLoaded = false;
let geocodeCache = {};
let useCustomDepartureTime = false;

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

function updateSaveButtonLabel() {
  if (!saveSearchBtn) {
    return;
  }
  saveSearchBtn.textContent = editingSearchId ? "Update area" : "Save search";
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

function clampMinutes(value) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed)) {
    return DEFAULT_DRIVE_TIME_MINUTES;
  }
  return Math.min(Math.max(Math.round(parsed), MIN_DRIVE_TIME_MINUTES), MAX_DRIVE_TIME_MINUTES);
}

function resolveDriveTimeMinutes() {
  const minutes = clampMinutes(driveTimeMinutesInput?.value);
  if (driveTimeMinutesInput) {
    driveTimeMinutesInput.value = String(minutes);
  }
  return minutes;
}

function toDateInputValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function toTimeInputValue(date) {
  const hours = String(date.getHours()).padStart(2, "0");
  const mins = String(date.getMinutes()).padStart(2, "0");
  return `${hours}:${mins}`;
}

function getDefaultDepartureDate() {
  const now = new Date();
  const nowMs = now.getTime();
  const currentWeekday = now.getDay();
  let dayOffset = (DEFAULT_DEPARTURE_WEEKDAY - currentWeekday + 7) % 7;
  let candidate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + dayOffset, DEFAULT_DEPARTURE_HOUR, 0, 0, 0);
  if (candidate.getTime() <= nowMs + DEPARTURE_LEAD_MS) {
    candidate = new Date(candidate.getFullYear(), candidate.getMonth(), candidate.getDate() + 7, DEFAULT_DEPARTURE_HOUR, 0, 0, 0);
  }
  return candidate;
}

function ensureDefaultDepartureInputs() {
  const candidate = getDefaultDepartureDate();
  if (departureDateInput) {
    departureDateInput.value = toDateInputValue(candidate);
  }
  if (departureTimeInput) {
    departureTimeInput.value = toTimeInputValue(candidate);
  }
}

function parseCustomDepartureDate() {
  const rawDate = String(departureDateInput?.value || "").trim();
  const rawTime = String(departureTimeInput?.value || "").trim();
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(rawDate);
  const timeMatch = /^(\d{2}):(\d{2})$/.exec(rawTime);
  if (!match || !timeMatch) {
    return null;
  }
  const year = Number(match[1]);
  const month = Number(match[2]) - 1;
  const day = Number(match[3]);
  const hours = Number(timeMatch[1]);
  const mins = Number(timeMatch[2]);
  const date = new Date(year, month, day, hours, mins, 0, 0);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDepartureLabel(date) {
  return date.toLocaleString(undefined, {
    weekday: "short",
    day: "2-digit",
    month: "short",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function setCustomDepartureVisibility(isVisible) {
  useCustomDepartureTime = isVisible;
  if (customDepartureWrap) {
    customDepartureWrap.hidden = !isVisible;
  }
  if (customDepartureBtn) {
    customDepartureBtn.textContent = isVisible ? "Hide custom" : "Custom time";
  }
}

function resolveDepartureTime() {
  if (useCustomDepartureTime) {
    const custom = parseCustomDepartureDate();
    if (!custom) {
      throw new Error("Enter a valid custom date and time.");
    }
    if (custom.getTime() <= Date.now() + DEPARTURE_LEAD_MS) {
      throw new Error("Custom time must be in the future.");
    }
    return {
      iso: custom.toISOString(),
      label: formatDepartureLabel(custom),
    };
  }

  const defaultDeparture = getDefaultDepartureDate();
  return {
    iso: defaultDeparture.toISOString(),
    label: formatDepartureLabel(defaultDeparture),
  };
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

function getColorLabel(index) {
  return AREA_COLOR_LABELS[index] || `Colour ${index + 1}`;
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
    minutes: clampMinutes(raw?.minutes),
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

function setPeopleOverlayStatus(message, isError = false) {
  if (!peopleOverlayStatus) {
    return;
  }
  peopleOverlayStatus.textContent = message;
  peopleOverlayStatus.classList.toggle("error", isError);
}

function normalizeText(value) {
  return String(value || "").trim().toLowerCase();
}

function normalizeLocation(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeStatus(value) {
  return normalizeText(value) || "unknown";
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function loadGeocodeCache() {
  try {
    const raw = localStorage.getItem(GEOCODE_CACHE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    geocodeCache = parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    geocodeCache = {};
  }
}

function persistGeocodeCache() {
  try {
    localStorage.setItem(GEOCODE_CACHE_KEY, JSON.stringify(geocodeCache));
  } catch {
    // Ignore cache persistence errors.
  }
}

function buildClientQuery(client) {
  const parts = [client?.address, client?.town, client?.county, client?.postcode]
    .map((part) => normalizeLocation(part))
    .filter(Boolean);
  if (parts.length) {
    return parts.join(", ");
  }
  return normalizeLocation(client?.postcode || client?.location || "");
}

function buildCarerQuery(carer) {
  return normalizeLocation(carer?.postcode || carer?.location || "");
}

function buildClientLocationLabel(client) {
  return (
    normalizeLocation(client?.area) ||
    normalizeLocation(client?.location) ||
    normalizeLocation(client?.town) ||
    normalizeLocation(client?.postcode) ||
    "Unassigned"
  );
}

function buildCarerLocationLabel(carer) {
  return normalizeLocation(carer?.location || carer?.postcode) || "Unassigned";
}

function getOverlayTypeLabel(type) {
  return type === "companion" ? "Companion" : "Client";
}

function clearPeopleOverlayLayers() {
  for (const layer of peopleOverlayLayers.values()) {
    layer.remove();
  }
  peopleOverlayLayers.clear();
}

function deriveOverlayStatusOptions(items) {
  const set = new Set(items.map((item) => normalizeStatus(item.status)));
  const ordered = ["active", "pending", "archived"].filter((status) => set.has(status));
  const extra = Array.from(set)
    .filter((status) => !ordered.includes(status))
    .sort((a, b) => a.localeCompare(b));
  return [...ordered, ...extra];
}

function deriveOverlayLocationOptions(items) {
  const unique = new Set(
    items
      .map((item) => normalizeLocation(item.locationLabel))
      .filter(Boolean)
  );
  return Array.from(unique).sort((a, b) => a.localeCompare(b));
}

function formatStatusLabel(status) {
  const value = normalizeStatus(status);
  return value.charAt(0).toUpperCase() + value.slice(1);
}

function renderOverlayStatusFilters() {
  if (!overlayStatusFilters) {
    return;
  }
  const options = deriveOverlayStatusOptions(peopleOverlayData);
  if (!overlayStatusSet.size) {
    overlayStatusSet = new Set(options);
  } else {
    overlayStatusSet = new Set(Array.from(overlayStatusSet).filter((status) => options.includes(status)));
    if (!overlayStatusSet.size) {
      overlayStatusSet = new Set(options);
    }
  }

  overlayStatusFilters.innerHTML = "";
  for (const status of options) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `status-filter-btn${overlayStatusSet.has(status) ? " active" : ""}`;
    btn.textContent = formatStatusLabel(status);
    btn.addEventListener("click", () => {
      if (overlayStatusSet.has(status)) {
        overlayStatusSet.delete(status);
      } else {
        overlayStatusSet.add(status);
      }
      if (!overlayStatusSet.size) {
        overlayStatusSet.add(status);
      }
      renderOverlayStatusFilters();
      applyPeopleOverlayFilters();
    });
    overlayStatusFilters.appendChild(btn);
  }

  const allBtn = document.createElement("button");
  allBtn.type = "button";
  allBtn.className = `status-filter-btn${overlayStatusSet.size === options.length ? " active" : ""}`;
  allBtn.textContent = "All";
  allBtn.addEventListener("click", () => {
    overlayStatusSet = new Set(options);
    renderOverlayStatusFilters();
    applyPeopleOverlayFilters();
  });
  overlayStatusFilters.appendChild(allBtn);
}

function renderOverlayLocationOptions() {
  if (!overlayLocationSelect) {
    return;
  }
  const current = String(overlayLocationSelect.value || "all");
  const options = deriveOverlayLocationOptions(peopleOverlayData);
  overlayLocationSelect.innerHTML = "";
  const all = document.createElement("option");
  all.value = "all";
  all.textContent = "All locations";
  overlayLocationSelect.appendChild(all);
  for (const location of options) {
    const option = document.createElement("option");
    option.value = location;
    option.textContent = location;
    overlayLocationSelect.appendChild(option);
  }
  overlayLocationSelect.value = options.includes(current) ? current : "all";
}

function getFilteredOverlayPeople() {
  const typeFilter = String(overlayTypeSelect?.value || "all");
  const locationFilter = String(overlayLocationSelect?.value || "all");
  return peopleOverlayData.filter((item) => {
    const matchesType = typeFilter === "all" || item.type === typeFilter;
    const matchesStatus = overlayStatusSet.has(normalizeStatus(item.status));
    const matchesLocation = locationFilter === "all" || normalizeLocation(item.locationLabel) === locationFilter;
    return matchesType && matchesStatus && matchesLocation;
  });
}

function ensureOverlayLayer(item) {
  if (!map || !window.L || !Number.isFinite(item?.lat) || !Number.isFinite(item?.lng)) {
    return null;
  }
  const existing = peopleOverlayLayers.get(item.id);
  if (existing) {
    return existing;
  }
  const style =
    item.type === "client"
      ? { color: "#1f3c88", fillColor: "#31b7c8" }
      : { color: "#8b2f6e", fillColor: "#c9439b" };
  const marker = window.L.circleMarker([item.lat, item.lng], {
    radius: 5,
    color: style.color,
    weight: 2,
    fillColor: style.fillColor,
    fillOpacity: 0.85,
  });
  marker.bindPopup(
    `<strong>${escapeHtml(item.name)}</strong><br/>${escapeHtml(getOverlayTypeLabel(item.type))}<br/>${escapeHtml(item.locationLabel)}<br/>Status: ${escapeHtml(formatStatusLabel(item.status))}`
  );
  peopleOverlayLayers.set(item.id, marker);
  return marker;
}

function applyPeopleOverlayFilters() {
  const shouldShow = Boolean(showPeopleOverlayInput?.checked);
  const filtered = getFilteredOverlayPeople();
  const visibleIds = new Set(filtered.map((item) => item.id));

  for (const [id, marker] of peopleOverlayLayers.entries()) {
    if (!shouldShow || !visibleIds.has(id)) {
      marker.remove();
    }
  }

  if (shouldShow) {
    for (const item of filtered) {
      const marker = ensureOverlayLayer(item);
      if (marker && !map.hasLayer(marker)) {
        marker.addTo(map);
      }
    }
  }

  const total = peopleOverlayData.length;
  const typeLabel = String(overlayTypeSelect?.value || "all");
  setPeopleOverlayStatus(
    overlayLoaded
      ? `Showing ${shouldShow ? filtered.length : 0} of ${total} people (${typeLabel}, ${showPeopleOverlayInput?.checked ? "visible" : "hidden"}).`
      : "Overlay not loaded."
  );
}

function buildOverlayItems(clientsPayload, carersPayload) {
  const clients = Array.isArray(clientsPayload?.clients) ? clientsPayload.clients : [];
  const carers = Array.isArray(carersPayload?.carers) ? carersPayload.carers : [];
  const items = [];
  let index = 0;

  for (const client of clients) {
    const query = buildClientQuery(client);
    if (!query) {
      continue;
    }
    items.push({
      id: `client-${normalizeText(client.id || client.name || index) || index}-${index}`,
      type: "client",
      name: normalizeLocation(client.name || "Unnamed client"),
      status: normalizeStatus(client.status),
      locationLabel: buildClientLocationLabel(client),
      geocodeQuery: query,
    });
    index += 1;
  }

  for (const carer of carers) {
    const query = buildCarerQuery(carer);
    if (!query) {
      continue;
    }
    items.push({
      id: `companion-${normalizeText(carer.id || carer.name || index) || index}-${index}`,
      type: "companion",
      name: normalizeLocation(carer.name || "Unnamed companion"),
      status: normalizeStatus(carer.status),
      locationLabel: buildCarerLocationLabel(carer),
      geocodeQuery: query,
    });
    index += 1;
  }
  return items;
}

async function geocodeOverlayItems(items) {
  const queries = [];
  const queryToKey = new Map();
  for (const item of items) {
    const query = normalizeLocation(item.geocodeQuery);
    if (!query) {
      continue;
    }
    const cacheKey = query.toLowerCase();
    queryToKey.set(query, cacheKey);
    if (!geocodeCache[cacheKey]) {
      queries.push(query);
    }
  }

  if (queries.length) {
    for (let offset = 0; offset < queries.length; offset += OVERLAY_BATCH_SIZE) {
      const batch = queries.slice(offset, offset + OVERLAY_BATCH_SIZE);
      setPeopleOverlayStatus(
        `Resolving map points (${Math.min(offset + batch.length, queries.length)}/${queries.length})...`
      );
      const response = await authedFetch(GEOCODE_BATCH_ENDPOINT, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          queries: batch.map((query) => ({ id: query, query })),
        }),
      });
      const data = await response.json().catch(() => ({}));
      if (!response.ok) {
        throw new Error(data?.error || "Could not geocode overlay points.");
      }
      const points = Array.isArray(data?.points) ? data.points : [];
      for (const point of points) {
        const query = normalizeLocation(point.query);
        if (!query) {
          continue;
        }
        geocodeCache[query.toLowerCase()] = {
          lat: Number(point.lat),
          lng: Number(point.lng),
          formattedAddress: normalizeLocation(point.formattedAddress),
        };
      }
      persistGeocodeCache();
    }
  }

  return items
    .map((item) => {
      const cacheKey = normalizeLocation(item.geocodeQuery).toLowerCase();
      const point = geocodeCache[cacheKey];
      if (!point || !Number.isFinite(point.lat) || !Number.isFinite(point.lng)) {
        return null;
      }
      return {
        ...item,
        lat: point.lat,
        lng: point.lng,
        resolvedAddress: point.formattedAddress || item.geocodeQuery,
      };
    })
    .filter(Boolean);
}

async function loadPeopleOverlay() {
  try {
    if (loadPeopleOverlayBtn) {
      loadPeopleOverlayBtn.disabled = true;
    }
    clearPeopleOverlayLayers();
    setPeopleOverlayStatus("Loading clients and companions...");

    const [clientsPayload, carersPayload] = await Promise.all([
      directoryApi.listOneTouchClients({ limit: 500 }),
      directoryApi.listCarers({ limit: 500 }),
    ]);
    const baseItems = buildOverlayItems(clientsPayload, carersPayload);
    if (!baseItems.length) {
      peopleOverlayData = [];
      overlayLoaded = true;
      renderOverlayStatusFilters();
      renderOverlayLocationOptions();
      applyPeopleOverlayFilters();
      setPeopleOverlayStatus("No mappable people records found.");
      return;
    }

    const resolvedItems = await geocodeOverlayItems(baseItems);
    peopleOverlayData = resolvedItems;
    overlayLoaded = true;
    renderOverlayStatusFilters();
    renderOverlayLocationOptions();
    applyPeopleOverlayFilters();
    setPeopleOverlayStatus(`Loaded ${resolvedItems.length} mapped people.`);
  } catch (error) {
    console.error(error);
    setPeopleOverlayStatus(error?.message || "Could not load people overlay.", true);
  } finally {
    if (loadPeopleOverlayBtn) {
      loadPeopleOverlayBtn.disabled = false;
    }
  }
}

function downloadJson(filename, payload) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function initMap() {
  if (!driveTimeMapRoot || !window.L) {
    return;
  }

  map = window.L.map(driveTimeMapRoot, {
    zoomControl: true,
    attributionControl: true,
  }).setView([51.2802, 1.0789], 11);

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
    draggable: true,
    title: payload?.formattedAddress || "Selected location",
  }).addTo(map);
  previewMarker.on("drag", (event) => {
    if (!currentResult || !Array.isArray(currentResult.polygon) || !currentResult.polygon.length) {
      return;
    }
    const next = event.target.getLatLng();
    const prevCenter = currentResult.center;
    const latDelta = next.lat - prevCenter.lat;
    const lngDelta = next.lng - prevCenter.lng;
    currentResult.center = { lat: next.lat, lng: next.lng };
    currentResult.polygon = currentResult.polygon.map((point) => ({
      lat: point.lat + latDelta,
      lng: point.lng + lngDelta,
    }));
    if (previewPolygon) {
      previewPolygon.setLatLngs(currentResult.polygon.map((point) => [point.lat, point.lng]));
    }
    syncPreviewVertexMarkers();
  });

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
  previewPolygon.on("click", (event) => {
    addPreviewPoint(event.latlng);
  });

  const polygonBounds = previewPolygon.getBounds();
  if (polygonBounds.isValid()) {
    map.fitBounds(polygonBounds.pad(0.15));
  } else {
    map.setView([center.lat, center.lng], 11);
  }
  syncPreviewVertexMarkers();
}

function clearPreviewVertexMarkers() {
  for (const marker of previewVertexMarkers) {
    marker.remove();
  }
  previewVertexMarkers = [];
}

function findClosestSegmentInsertIndex(points, latlng) {
  if (!Array.isArray(points) || points.length < 2) {
    return points.length;
  }
  const point = { lat: Number(latlng.lat), lng: Number(latlng.lng) };
  let bestIndex = points.length;
  let bestDistance = Number.POSITIVE_INFINITY;

  for (let i = 0; i < points.length; i += 1) {
    const a = points[i];
    const b = points[(i + 1) % points.length];
    const ax = a.lng;
    const ay = a.lat;
    const bx = b.lng;
    const by = b.lat;
    const px = point.lng;
    const py = point.lat;
    const abx = bx - ax;
    const aby = by - ay;
    const abLenSq = abx * abx + aby * aby || 1;
    const t = Math.max(0, Math.min(1, ((px - ax) * abx + (py - ay) * aby) / abLenSq));
    const cx = ax + abx * t;
    const cy = ay + aby * t;
    const dx = px - cx;
    const dy = py - cy;
    const distSq = dx * dx + dy * dy;
    if (distSq < bestDistance) {
      bestDistance = distSq;
      bestIndex = i + 1;
    }
  }
  return bestIndex;
}

function updatePreviewPolygonShape() {
  if (!currentResult || !previewPolygon) {
    return;
  }
  previewPolygon.setLatLngs(currentResult.polygon.map((point) => [point.lat, point.lng]));
}

function syncPreviewVertexMarkers() {
  clearPreviewVertexMarkers();
  if (!map || !window.L || !currentResult || !Array.isArray(currentResult.polygon)) {
    return;
  }

  currentResult.polygon.forEach((point, index) => {
    const marker = window.L.marker([point.lat, point.lng], {
      draggable: true,
      keyboard: false,
      icon: window.L.divIcon({
        className: "drive-time-vertex-icon",
        iconSize: [12, 12],
      }),
      title: "Drag to reshape. Right-click to delete.",
    }).addTo(map);

    marker.on("drag", (event) => {
      const latLng = event.target.getLatLng();
      currentResult.polygon[index] = { lat: latLng.lat, lng: latLng.lng };
      updatePreviewPolygonShape();
    });

    marker.on("contextmenu", () => {
      if (!currentResult || currentResult.polygon.length <= MIN_POLYGON_POINTS) {
        setStatus("Area needs at least 3 points.", true);
        return;
      }
      currentResult.polygon.splice(index, 1);
      updatePreviewPolygonShape();
      syncPreviewVertexMarkers();
      setStatus("Point deleted.");
    });

    previewVertexMarkers.push(marker);
  });
}

function addPreviewPoint(latlng) {
  if (!currentResult || !Array.isArray(currentResult.polygon) || !currentResult.polygon.length) {
    return;
  }
  const insertIndex = findClosestSegmentInsertIndex(currentResult.polygon, latlng);
  currentResult.polygon.splice(insertIndex, 0, { lat: latlng.lat, lng: latlng.lng });
  updatePreviewPolygonShape();
  syncPreviewVertexMarkers();
  setStatus("Point added.");
}

function beginEditingSavedSearch(search) {
  if (!search) {
    return;
  }
  editingSearchId = search.id;
  currentResult = {
    query: search.query || search.formattedAddress || "",
    formattedAddress: search.formattedAddress || search.query || search.name || "Saved area",
    minutes: clampMinutes(search.minutes),
    center: { lat: Number(search.center?.lat), lng: Number(search.center?.lng) },
    polygon: sanitizePolygon(search.polygon),
    quality: search.quality || null,
  };
  if (driveTimeMinutesInput) {
    driveTimeMinutesInput.value = String(currentResult.minutes);
  }
  if (searchNameInput) {
    searchNameInput.value = search.name || currentResult.formattedAddress;
  }
  renderPreview({
    center: currentResult.center,
    polygon: currentResult.polygon,
    formattedAddress: currentResult.formattedAddress,
  });
  if (saveSearchBtn) {
    saveSearchBtn.disabled = false;
  }
  updateSaveButtonLabel();
  setStatus(`Editing '${search.name}'. Adjust points then click Update area.`);
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

    const colourSelect = document.createElement("select");
    colourSelect.className = "drive-time-color-select";
    colourSelect.setAttribute("aria-label", `Colour for ${search.name}`);
    for (let i = 0; i < AREA_COLORS.length; i += 1) {
      const option = document.createElement("option");
      option.value = String(i);
      option.textContent = getColorLabel(i);
      colourSelect.appendChild(option);
    }
    colourSelect.value = String(Math.max(0, Number(search.colorIndex || 0) % AREA_COLORS.length));
    colourSelect.addEventListener("change", () => {
      search.colorIndex = Number(colourSelect.value);
      persistSavedSearches();
      syncSearchLayer(search);
      renderSavedSearches();
      setStatus(`Updated colour for '${search.name}'.`);
    });

    const editBtn = document.createElement("button");
    editBtn.type = "button";
    editBtn.className = "secondary";
    editBtn.textContent = "Edit";
    editBtn.addEventListener("click", () => {
      beginEditingSavedSearch(search);
    });

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
      const wasEditing = editingSearchId === search.id;
      removeAreaLayers(search.id);
      savedSearches = savedSearches.filter((item) => item.id !== search.id);
      persistSavedSearches();
      renderSavedSearches();
      if (wasEditing) {
        editingSearchId = null;
        currentResult = null;
        if (previewMarker) {
          previewMarker.remove();
          previewMarker = null;
        }
        if (previewPolygon) {
          previewPolygon.remove();
          previewPolygon = null;
        }
        clearPreviewVertexMarkers();
        if (saveSearchBtn) {
          saveSearchBtn.disabled = true;
        }
        if (searchNameInput) {
          searchNameInput.value = "";
        }
        updateSaveButtonLabel();
      }
      setStatus(`Removed '${search.name}'.`);
    });

    actions.append(colourSelect, editBtn, focusBtn, removeBtn);
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
  const editingTarget = editingSearchId ? savedSearches.find((item) => item.id === editingSearchId) : null;
  const area = sanitizeSavedSearch(
    {
      id: editingTarget?.id || createSearchId(),
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

  if (editingTarget) {
    const index = savedSearches.findIndex((item) => item.id === editingTarget.id);
    if (index >= 0) {
      savedSearches[index] = {
        ...savedSearches[index],
        ...area,
      };
    }
  } else {
    savedSearches.unshift(area);
  }
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
  clearPreviewVertexMarkers();

  if (searchNameInput) {
    searchNameInput.value = "";
  }
  currentResult = null;
  editingSearchId = null;
  updateSaveButtonLabel();
  if (saveSearchBtn) {
    saveSearchBtn.disabled = true;
  }
  setStatus(editingTarget ? `Updated '${area.name}'.` : `Saved '${area.name}'.`);
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

function exportSavedSearches() {
  if (!savedSearches.length) {
    setStatus("No saved searches to export.", true);
    return;
  }
  const stamp = new Date().toISOString().slice(0, 19).replaceAll(":", "-");
  downloadJson(`drive-time-searches-${stamp}.json`, {
    version: 1,
    exportedAt: new Date().toISOString(),
    searches: savedSearches,
  });
  setStatus(`Exported ${savedSearches.length} saved search(es).`);
}

function importSavedSearchesFromText(rawText) {
  let parsed = null;
  try {
    parsed = JSON.parse(String(rawText || ""));
  } catch {
    throw new Error("Import file is not valid JSON.");
  }

  const incoming = Array.isArray(parsed) ? parsed : Array.isArray(parsed?.searches) ? parsed.searches : null;
  if (!incoming) {
    throw new Error("Import JSON must be an array or an object with 'searches'.");
  }

  const existingIds = new Set(savedSearches.map((search) => search.id));
  const imported = [];
  for (let index = 0; index < incoming.length; index += 1) {
    const item = sanitizeSavedSearch(incoming[index], savedSearches.length + imported.length);
    if (!item) {
      continue;
    }
    if (existingIds.has(item.id)) {
      item.id = createSearchId();
    }
    existingIds.add(item.id);
    imported.push(item);
  }

  if (!imported.length) {
    throw new Error("No valid searches found in import file.");
  }

  savedSearches = [...imported, ...savedSearches].slice(0, 100);
  persistSavedSearches();
  renderSavedSearches();
  syncAllSearchLayers();
  setStatus(`Imported ${imported.length} search(es). Saved locally.`);
}

async function drawDriveTimeArea() {
  const location = String(locationInput?.value || "").trim();
  if (!location) {
    setStatus("Enter a location first.", true);
    return;
  }
  const minutes = resolveDriveTimeMinutes();
  let departure = null;
  try {
    departure = resolveDepartureTime();
  } catch (error) {
    setStatus(error?.message || "Invalid departure time.", true);
    return;
  }
  setBusy(true);
  setStatus(`Calculating ${minutes}-minute drive-time area (${departure.label})...`);
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
        minutes,
        departureTime: departure.iso,
      }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data?.error || "Could not calculate drive-time area.");
    }

    currentResult = {
      query: location,
      formattedAddress: String(data.formattedAddress || location),
      minutes: clampMinutes(data.minutes || minutes),
      center: data.center,
      polygon: sanitizePolygon(data.polygon),
      quality: data.quality || null,
    };
    editingSearchId = null;
    updateSaveButtonLabel();

    renderPreview({
      ...data,
      polygon: currentResult.polygon,
    });
    if (searchNameInput) {
      searchNameInput.value = currentResult.formattedAddress || currentResult.query;
    }
    setStatus(
      `Showing ${currentResult.minutes}-minute drive-time area for ${currentResult.formattedAddress || location} (${departure.label}).`
    );
    if (driveTimeMeta) {
      driveTimeMeta.textContent = formatQualityText(data.quality);
    }
    if (saveSearchBtn) {
      saveSearchBtn.disabled = false;
    }
  } catch (error) {
    console.error(error);
    currentResult = null;
    editingSearchId = null;
    updateSaveButtonLabel();
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
    loadGeocodeCache();
    loadSavedSearches();
    ensureDefaultDepartureInputs();
    setCustomDepartureVisibility(false);
    updateSaveButtonLabel();
    renderSavedSearches();
    syncAllSearchLayers();
    setPeopleOverlayStatus("Overlay not loaded.");
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

loadPeopleOverlayBtn?.addEventListener("click", () => {
  void loadPeopleOverlay();
});

showPeopleOverlayInput?.addEventListener("change", () => {
  applyPeopleOverlayFilters();
});

overlayTypeSelect?.addEventListener("change", () => {
  applyPeopleOverlayFilters();
});

overlayLocationSelect?.addEventListener("change", () => {
  applyPeopleOverlayFilters();
});

exportSearchesBtn?.addEventListener("click", () => {
  exportSavedSearches();
});

importSearchesBtn?.addEventListener("click", () => {
  importSearchesFileInput?.click();
});

importSearchesFileInput?.addEventListener("change", async () => {
  const file = importSearchesFileInput.files?.[0];
  if (!file) {
    return;
  }
  try {
    const text = await file.text();
    importSavedSearchesFromText(text);
  } catch (error) {
    setStatus(error?.message || "Could not import searches.", true);
  } finally {
    importSearchesFileInput.value = "";
  }
});

locationInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") {
    return;
  }
  event.preventDefault();
  void drawDriveTimeArea();
});

driveTimeMinutesInput?.addEventListener("change", () => {
  resolveDriveTimeMinutes();
});

customDepartureBtn?.addEventListener("click", () => {
  const nextVisible = !useCustomDepartureTime;
  setCustomDepartureVisibility(nextVisible);
  if (nextVisible && (!departureDateInput?.value || !departureTimeInput?.value)) {
    ensureDefaultDepartureInputs();
  }
});

useDefaultDepartureBtn?.addEventListener("click", () => {
  ensureDefaultDepartureInputs();
  setCustomDepartureVisibility(false);
  setStatus("Using default measurement time: Wednesday 10:00.");
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
