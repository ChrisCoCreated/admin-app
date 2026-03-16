import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260311";

const signOutBtn = document.getElementById("signOutBtn");
const locationInput = document.getElementById("locationInput");
const setOfficeBtn = document.getElementById("setOfficeBtn");
const clientOfficeSelect = document.getElementById("clientOfficeSelect");
const setOfficeFromClientBtn = document.getElementById("setOfficeFromClientBtn");
const driveTimeMinutesInput = document.getElementById("driveTimeMinutesInput");
const customDepartureBtn = document.getElementById("customDepartureBtn");
const customDepartureWrap = document.getElementById("customDepartureWrap");
const departureDateInput = document.getElementById("departureDateInput");
const departureTimeInput = document.getElementById("departureTimeInput");
const useDefaultDepartureBtn = document.getElementById("useDefaultDepartureBtn");
const searchNameInput = document.getElementById("searchNameInput");
const editAddPlacesBtn = document.getElementById("editAddPlacesBtn");
const saveSearchBtn = document.getElementById("saveSearchBtn");
const clearAcceptedPlacesBtn = document.getElementById("clearAcceptedPlacesBtn");
const acceptedPlacesList = document.getElementById("acceptedPlacesList");
const noAcceptedPlacesMessage = document.getElementById("noAcceptedPlacesMessage");
const toggleSavedSearchesBtn = document.getElementById("toggleSavedSearchesBtn");
const savedSearchesContent = document.getElementById("savedSearchesContent");
const savedSearchFilterSelect = document.getElementById("savedSearchFilterSelect");
const duplicateSelectedSearchesBtn = document.getElementById("duplicateSelectedSearchesBtn");
const mergeSelectedSearchesBtn = document.getElementById("mergeSelectedSearchesBtn");
const showAllSearchesBtn = document.getElementById("showAllSearchesBtn");
const hideAllSearchesBtn = document.getElementById("hideAllSearchesBtn");
const exportSearchesBtn = document.getElementById("exportSearchesBtn");
const importSearchesBtn = document.getElementById("importSearchesBtn");
const importSearchesFileInput = document.getElementById("importSearchesFileInput");
const loadPeopleOverlayBtn = document.getElementById("loadPeopleOverlayBtn");
const showPeopleOverlayInput = document.getElementById("showPeopleOverlayInput");
const overlayTypeSelect = document.getElementById("overlayTypeSelect");
const overlayLocationSelect = document.getElementById("overlayLocationSelect");
const overlayClientAreaFilters = document.getElementById("overlayClientAreaFilters");
const overlayCompanionAreaFilters = document.getElementById("overlayCompanionAreaFilters");
const overlayCompanionCareCompSelect = document.getElementById("overlayCompanionCareCompSelect");
const overlayStatusFilters = document.getElementById("overlayStatusFilters");
const peopleOverlayStatus = document.getElementById("peopleOverlayStatus");
const savedSearchesList = document.getElementById("savedSearchesList");
const noSavedSearchesMessage = document.getElementById("noSavedSearchesMessage");
const driveTimeStatus = document.getElementById("driveTimeStatus");
const driveTimeMeta = document.getElementById("driveTimeMeta");
const driveTimeMapRoot = document.getElementById("driveTimeMap");
const mapClickFeedback = document.getElementById("mapClickFeedback");

const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
const OFFICE_CATCHMENT_ENDPOINT = API_BASE_URL
  ? `${API_BASE_URL}/api/maps/office-catchment/check-click`
  : "/api/maps/office-catchment/check-click";
const GEOCODE_BATCH_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/maps/geocode-batch` : "/api/maps/geocode-batch";
const USE_OFFICE_CATCHMENT_MODE = FRONTEND_CONFIG.useOfficeCatchmentMode !== false;
const SAVED_SEARCHES_KEY = "thrive.drivetime.saved.v2";
const LEGACY_SAVED_SEARCHES_KEY = "thrive.drivetime.saved.v1";
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

const OFFICE = {
  name: String(FRONTEND_CONFIG.mapOffice?.name || "Canterbury Office"),
  postcode: String(FRONTEND_CONFIG.mapOffice?.postcode || "CT1"),
  lat: Number(FRONTEND_CONFIG.mapOffice?.lat ?? 51.2802),
  lng: Number(FRONTEND_CONFIG.mapOffice?.lng ?? 1.0789),
};
let activeOffice = null;

let map = null;
let officeMarker = null;
let previewPolygon = null;
let previewVertexMarkers = [];
let currentAcceptedMarkers = [];
let editingSearchId = null;
let savedSearches = [];
const areaLayers = new Map();
let peopleOverlayData = [];
const peopleOverlayLayers = new Map();
let overlayStatusSet = new Set();
let overlayClientAreaSet = new Set();
let overlayCompanionAreaSet = new Set();
let overlayLoaded = false;
let geocodeCache = {};
let useCustomDepartureTime = false;
let currentCatchment = createEmptyCatchment();
let clickFeedbackTimer = null;
let savedSearchFilter = "active";
let clientOfficeOptions = [];
let editAddPlacesMode = false;
let selectedSavedSearchIds = new Set();

const MAP_PANES = {
  searchArea: "search-area-pane",
  searchOffice: "search-office-pane",
  people: "people-pane",
  activeArea: "active-area-pane",
  activeOffice: "active-office-pane",
  activeVertex: "active-vertex-pane",
};

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

function createEmptyCatchment() {
  return {
    office: activeOffice ? { ...activeOffice } : null,
    thresholdMinutes: DEFAULT_DRIVE_TIME_MINUTES,
    acceptedPlaces: [],
    polygon: [],
    departureTimeMode: "default",
    departureTime: null,
  };
}

function setStatus(message, isError = false) {
  if (!driveTimeStatus) {
    return;
  }
  driveTimeStatus.textContent = message;
  driveTimeStatus.classList.toggle("error", isError);
}

function showClickFeedback(message, tone = "info") {
  if (!mapClickFeedback) {
    return;
  }
  if (clickFeedbackTimer) {
    window.clearTimeout(clickFeedbackTimer);
    clickFeedbackTimer = null;
  }
  mapClickFeedback.hidden = false;
  mapClickFeedback.textContent = String(message || "");
  mapClickFeedback.classList.remove("is-success", "is-error", "is-info", "is-visible");
  mapClickFeedback.classList.add(`is-${tone}`, "is-visible");
  clickFeedbackTimer = window.setTimeout(() => {
    mapClickFeedback.classList.remove("is-visible");
    window.setTimeout(() => {
      mapClickFeedback.hidden = true;
    }, 220);
    clickFeedbackTimer = null;
  }, 3000);
}

function setBusy(isBusy) {
  if (setOfficeBtn) {
    setOfficeBtn.disabled = isBusy;
  }
  if (setOfficeFromClientBtn) {
    setOfficeFromClientBtn.disabled = isBusy || !String(clientOfficeSelect?.value || "").trim();
  }
  if (saveSearchBtn) {
    saveSearchBtn.disabled = isBusy || !canSaveCurrentCatchment();
  }
}

function canSaveCurrentCatchment() {
  return Array.isArray(currentCatchment?.acceptedPlaces) && currentCatchment.acceptedPlaces.length > 0;
}

function updateSaveButtonState() {
  if (!saveSearchBtn) {
    return;
  }
  saveSearchBtn.disabled = !canSaveCurrentCatchment();
  saveSearchBtn.textContent = editingSearchId ? "Update area" : "Save search";
}

function updateEditAddPlacesButtonState() {
  if (!editAddPlacesBtn) {
    return;
  }
  editAddPlacesBtn.disabled = !editingSearchId;
  editAddPlacesBtn.textContent = editAddPlacesMode ? "Stop adding places" : "Add places";
  editAddPlacesBtn.classList.toggle("active", editAddPlacesMode);
}

function setSavedSearchesCollapsed(isCollapsed) {
  if (savedSearchesContent) {
    savedSearchesContent.hidden = isCollapsed;
  }
  if (toggleSavedSearchesBtn) {
    toggleSavedSearchesBtn.textContent = isCollapsed ? "Show saved searches" : "Hide saved searches";
    toggleSavedSearchesBtn.setAttribute("aria-expanded", String(!isCollapsed));
  }
}

function setClientOfficeButtonState() {
  if (!setOfficeFromClientBtn) {
    return;
  }
  setOfficeFromClientBtn.disabled = !String(clientOfficeSelect?.value || "").trim();
}

function updateClearAcceptedButtonState() {
  if (!clearAcceptedPlacesBtn) {
    return;
  }
  clearAcceptedPlacesBtn.disabled = !(currentCatchment?.acceptedPlaces?.length > 0);
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
  currentCatchment.thresholdMinutes = minutes;
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
      mode: "custom",
    };
  }

  const defaultDeparture = getDefaultDepartureDate();
  return {
    iso: defaultDeparture.toISOString(),
    label: formatDepartureLabel(defaultDeparture),
    mode: "default",
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

function sanitizeAcceptedPlaces(points) {
  if (!Array.isArray(points)) {
    return [];
  }
  return points
    .map((point) => ({
      key: normalizeSearchName(point?.key || ""),
      name: normalizeSearchName(point?.name || point?.formattedAddress || "Unknown location"),
      postcode: normalizeSearchName(point?.postcode || ""),
      formattedAddress: normalizeSearchName(point?.formattedAddress || point?.name || ""),
      lat: Number(point?.lat),
      lng: Number(point?.lng),
      durationMinutes: Number(point?.durationMinutes),
      distanceMiles: Number(point?.distanceMiles),
    }))
    .filter((point) => Number.isFinite(point.lat) && Number.isFinite(point.lng) && point.key);
}

function sanitizeOffice(rawOffice) {
  const lat = Number(rawOffice?.lat);
  const lng = Number(rawOffice?.lng);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
    return { ...OFFICE };
  }
  return {
    name: normalizeSearchName(rawOffice?.name || OFFICE.name),
    postcode: normalizeSearchName(rawOffice?.postcode || ""),
    lat,
    lng,
  };
}

function sanitizeSavedSearch(raw, fallbackIndex = 0) {
  const colorIndex = Number(raw?.colorIndex);
  const common = {
    id: String(raw?.id || createSearchId()),
    name: normalizeSearchName(raw?.name || "") || `Saved area ${fallbackIndex + 1}`,
    savedAt: String(raw?.savedAt || raw?.createdAt || new Date().toISOString()),
    visible: raw?.visible !== false,
    archived: raw?.archived === true,
    colorIndex: Number.isInteger(colorIndex) && colorIndex >= 0 ? colorIndex : fallbackIndex,
  };

  const office = raw?.office ? sanitizeOffice(raw.office) : null;
  const acceptedPlaces = sanitizeAcceptedPlaces(raw?.acceptedPlaces);
  const polygon = sanitizePolygon(raw?.polygon);
  const thresholdMinutes = clampMinutes(raw?.thresholdMinutes ?? raw?.minutes);

  if (office) {
    return {
      ...common,
      mode: "office",
      office,
      thresholdMinutes,
      departureTimeMode: raw?.departureTimeMode === "custom" ? "custom" : "default",
      departureTime: normalizeSearchName(raw?.departureTime || "") || null,
      acceptedPlaces,
      polygon,
      qualitySummary: raw?.qualitySummary || null,
      formattedAddress: normalizeSearchName(raw?.formattedAddress || office.name),
      query: normalizeSearchName(raw?.query || office.name),
    };
  }

  // Legacy radial shape support in read/edit mode.
  const legacyCenterLat = Number(raw?.center?.lat);
  const legacyCenterLng = Number(raw?.center?.lng);
  if (!polygon.length || !Number.isFinite(legacyCenterLat) || !Number.isFinite(legacyCenterLng)) {
    return null;
  }

  return {
    ...common,
    mode: "legacy",
    office: {
      name: normalizeSearchName(raw?.formattedAddress || raw?.query || "Legacy area center"),
      postcode: "",
      lat: legacyCenterLat,
      lng: legacyCenterLng,
    },
    thresholdMinutes,
    departureTimeMode: "default",
    departureTime: null,
    acceptedPlaces: [],
    polygon,
    qualitySummary: raw?.quality || null,
    formattedAddress: normalizeSearchName(raw?.formattedAddress || ""),
    query: normalizeSearchName(raw?.query || ""),
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

    if (!savedSearches.length) {
      const legacyRaw = localStorage.getItem(LEGACY_SAVED_SEARCHES_KEY);
      const legacyParsed = legacyRaw ? JSON.parse(legacyRaw) : [];
      const legacyItems = Array.isArray(legacyParsed) ? legacyParsed : [];
      savedSearches = legacyItems
        .map((entry, index) => sanitizeSavedSearch(entry, index))
        .filter(Boolean)
        .slice(0, 100);
      persistSavedSearches();
    }
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

function getFilteredSavedSearches() {
  if (savedSearchFilter === "archived") {
    return savedSearches.filter((search) => search.archived);
  }
  if (savedSearchFilter === "all") {
    return savedSearches;
  }
  return savedSearches.filter((search) => !search.archived);
}

function getSelectedSavedSearches() {
  return savedSearches.filter((search) => selectedSavedSearchIds.has(search.id));
}

function updateSavedSearchSelectionActions() {
  const selectedCount = getSelectedSavedSearches().length;
  if (duplicateSelectedSearchesBtn) {
    duplicateSelectedSearchesBtn.disabled = selectedCount < 1;
  }
  if (mergeSelectedSearchesBtn) {
    mergeSelectedSearchesBtn.disabled = selectedCount < 2;
  }
}

function getEditingSearch() {
  if (!editingSearchId) {
    return null;
  }
  return savedSearches.find((search) => search.id === editingSearchId) || null;
}

function persistEditingSearchChanges(statusMessage = "") {
  const editingSearch = getEditingSearch();
  if (!editingSearch) {
    return;
  }

  editingSearch.office = sanitizeOffice(currentCatchment.office);
  editingSearch.thresholdMinutes = clampMinutes(currentCatchment.thresholdMinutes);
  editingSearch.acceptedPlaces = sanitizeAcceptedPlaces(currentCatchment.acceptedPlaces);
  editingSearch.polygon = sanitizePolygon(currentCatchment.polygon);
  editingSearch.departureTimeMode = currentCatchment.departureTimeMode === "custom" ? "custom" : "default";
  editingSearch.departureTime = currentCatchment.departureTime || null;
  editingSearch.formattedAddress = normalizeSearchName(currentCatchment.office.name);
  editingSearch.query = normalizeSearchName(currentCatchment.office.name);
  editingSearch.savedAt = new Date().toISOString();
  editingSearch.qualitySummary = {
    ...(editingSearch.qualitySummary || {}),
    acceptedPlaces: editingSearch.acceptedPlaces.length,
  };

  persistSavedSearches();
  syncSearchLayer(editingSearch);
  renderSavedSearches();
  if (statusMessage) {
    setStatus(statusMessage);
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

function buildClientOfficeLabel(client, query) {
  const name = normalizeLocation(client?.name || "Unnamed client");
  const location = normalizeLocation(query || buildClientQuery(client));
  return location ? `${name} - ${location}` : name;
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

function deriveCompanionAreaOptions(items) {
  const values = new Set(
    items
      .filter((item) => item.type === "companion")
      .map((item) => normalizeLocation(item.areaLabel))
      .filter(Boolean)
  );
  return Array.from(values).sort((a, b) => a.localeCompare(b));
}

function deriveClientAreaOptions(items) {
  const values = new Set(
    items
      .filter((item) => item.type === "client")
      .map((item) => normalizeLocation(item.areaLabel))
      .filter(Boolean)
  );
  return Array.from(values).sort((a, b) => a.localeCompare(b));
}

function deriveCompanionCareCompOptions(items) {
  const values = new Set(
    items
      .filter((item) => item.type === "companion")
      .map((item) => normalizeLocation(item.careCompTag))
      .filter(Boolean)
  );
  return Array.from(values).sort((a, b) => a.localeCompare(b));
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
    const defaultOptions = options.filter((status) => status !== "archived");
    overlayStatusSet = new Set(defaultOptions.length ? defaultOptions : options);
  } else {
    overlayStatusSet = new Set(Array.from(overlayStatusSet).filter((status) => options.includes(status)));
    if (!overlayStatusSet.size) {
      const defaultOptions = options.filter((status) => status !== "archived");
      overlayStatusSet = new Set(defaultOptions.length ? defaultOptions : options);
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

function renderAreaFilterButtons(root, options, selectedSet, onChange, emptyLabel) {
  if (!root) {
    return;
  }
  if (!selectedSet.size) {
    for (const option of options) {
      selectedSet.add(option);
    }
  }
  selectedSet.forEach((value) => {
    if (!options.includes(value)) {
      selectedSet.delete(value);
    }
  });
  if (!selectedSet.size) {
    for (const option of options) {
      selectedSet.add(option);
    }
  }

  root.innerHTML = "";
  if (!options.length) {
    const empty = document.createElement("span");
    empty.className = "muted";
    empty.textContent = emptyLabel;
    root.appendChild(empty);
    return;
  }

  for (const optionValue of options) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `area-filter-btn${selectedSet.has(optionValue) ? " active" : ""}`;
    btn.textContent = optionValue;
    btn.addEventListener("click", () => {
      if (selectedSet.has(optionValue)) {
        selectedSet.delete(optionValue);
      } else {
        selectedSet.add(optionValue);
      }
      if (!selectedSet.size) {
        selectedSet.add(optionValue);
      }
      onChange();
    });
    root.appendChild(btn);
  }
}

function renderCompanionOverlayOptions() {
  if (!overlayCompanionCareCompSelect) {
    return;
  }

  const currentCareComp = String(overlayCompanionCareCompSelect.value || "all");
  const clientAreaOptions = deriveClientAreaOptions(peopleOverlayData);
  const companionAreaOptions = deriveCompanionAreaOptions(peopleOverlayData);
  const careCompOptions = deriveCompanionCareCompOptions(peopleOverlayData);

  renderAreaFilterButtons(
    overlayClientAreaFilters,
    clientAreaOptions,
    overlayClientAreaSet,
    () => {
      renderCompanionOverlayOptions();
      applyPeopleOverlayFilters();
    },
    "No client areas"
  );

  renderAreaFilterButtons(
    overlayCompanionAreaFilters,
    companionAreaOptions,
    overlayCompanionAreaSet,
    () => {
      renderCompanionOverlayOptions();
      applyPeopleOverlayFilters();
    },
    "No companion areas"
  );

  overlayCompanionCareCompSelect.innerHTML = '<option value="all">All tags</option>';
  for (const tag of careCompOptions) {
    const option = document.createElement("option");
    option.value = tag;
    option.textContent = tag;
    overlayCompanionCareCompSelect.appendChild(option);
  }

  overlayCompanionCareCompSelect.value = careCompOptions.includes(currentCareComp) ? currentCareComp : "all";
}

function getFilteredOverlayPeople() {
  const typeFilter = String(overlayTypeSelect?.value || "all");
  const locationFilter = String(overlayLocationSelect?.value || "all");
  const companionCareCompFilter = String(overlayCompanionCareCompSelect?.value || "all");
  return peopleOverlayData.filter((item) => {
    const matchesType = typeFilter === "all" || item.type === typeFilter;
    const matchesStatus = overlayStatusSet.has(normalizeStatus(item.status));
    const matchesLocation = locationFilter === "all" || normalizeLocation(item.locationLabel) === locationFilter;
    if (!(matchesType && matchesStatus && matchesLocation)) {
      return false;
    }

    const normalizedArea = normalizeLocation(item.areaLabel);
    if (item.type === "client") {
      return !overlayClientAreaSet.size || overlayClientAreaSet.has(normalizedArea);
    }

    if (item.type !== "companion") {
      return true;
    }

    if (overlayCompanionAreaSet.size && !overlayCompanionAreaSet.has(normalizedArea)) {
      return false;
    }
    if (companionCareCompFilter !== "all" && normalizeLocation(item.careCompTag) !== companionCareCompFilter) {
      return false;
    }

    return true;
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
  const normalizedArea = normalizeLocation(item.areaLabel || item.locationLabel || "");
  const isCentralArea = normalizedArea.toLowerCase().includes("central");
  const marker = window.L.marker([item.lat, item.lng], {
    icon: window.L.divIcon({
      className: `people-overlay-marker ${item.type === "client" ? "is-client" : "is-companion"} is-${normalizeStatus(
        item.status
      )}${isCentralArea ? " is-diamond" : " is-dot"}`,
      iconSize: [16, 16],
      iconAnchor: [8, 8],
      popupAnchor: [0, -8],
    }),
    pane: MAP_PANES.people,
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
      areaLabel: normalizeLocation(client.area || client.location || client.town || ""),
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
      areaLabel: normalizeLocation(carer.area || ""),
      careCompTag: normalizeLocation(carer.careCompanionshipTag || ""),
      geocodeQuery: query,
    });
    index += 1;
  }
  return items;
}

async function geocodeOverlayItems(items) {
  const queries = [];
  for (const item of items) {
    const query = normalizeLocation(item.geocodeQuery);
    if (!query) {
      continue;
    }
    const cacheKey = query.toLowerCase();
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
      renderCompanionOverlayOptions();
      applyPeopleOverlayFilters();
      setPeopleOverlayStatus("No mappable people records found.");
      return;
    }

    const resolvedItems = await geocodeOverlayItems(baseItems);
    peopleOverlayData = resolvedItems;
    overlayLoaded = true;
    renderOverlayStatusFilters();
    renderOverlayLocationOptions();
    renderCompanionOverlayOptions();
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

async function loadClientOfficeOptions() {
  if (!clientOfficeSelect) {
    return;
  }

  try {
    const payload = await directoryApi.listOneTouchClients({ limit: 500 });
    const clients = Array.isArray(payload?.clients) ? payload.clients : [];
    clientOfficeOptions = clients
      .map((client, index) => {
        const query = buildClientQuery(client);
        if (!query) {
          return null;
        }
        return {
          id: String(client?.id || `client-${index}`),
          label: buildClientOfficeLabel(client, query),
          query,
        };
      })
      .filter(Boolean)
      .sort((a, b) => a.label.localeCompare(b.label));

    clientOfficeSelect.innerHTML = '<option value="">Choose client location</option>';
    for (const optionData of clientOfficeOptions) {
      const option = document.createElement("option");
      option.value = optionData.id;
      option.textContent = optionData.label;
      clientOfficeSelect.appendChild(option);
    }
    setClientOfficeButtonState();
  } catch (error) {
    console.error(error);
    if (clientOfficeSelect) {
      clientOfficeSelect.innerHTML = '<option value="">Client locations unavailable</option>';
    }
    setClientOfficeButtonState();
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

function createMapPanes() {
  if (!map) {
    return;
  }

  const paneConfig = [
    { name: MAP_PANES.searchArea, zIndex: 410 },
    { name: MAP_PANES.searchOffice, zIndex: 420 },
    { name: MAP_PANES.people, zIndex: 450 },
    { name: MAP_PANES.activeArea, zIndex: 640 },
    { name: MAP_PANES.activeOffice, zIndex: 645 },
    { name: MAP_PANES.activeVertex, zIndex: 650 },
  ];

  for (const pane of paneConfig) {
    const layerPane = map.getPane(pane.name) || map.createPane(pane.name);
    layerPane.style.zIndex = String(pane.zIndex);
  }
}

function initMap() {
  if (!driveTimeMapRoot || !window.L) {
    return;
  }

  map = window.L.map(driveTimeMapRoot, {
    zoomControl: true,
    attributionControl: true,
  }).setView([OFFICE.lat, OFFICE.lng], 11);

  createMapPanes();

  window.L.tileLayer("https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png", {
    maxZoom: 19,
    attribution:
      '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> &copy; <a href="https://carto.com/attributions">CARTO</a>',
  }).addTo(map);

  map.on("click", (event) => {
    if (!USE_OFFICE_CATCHMENT_MODE) {
      return;
    }
    if (editingSearchId && !editAddPlacesMode) {
      return;
    }
    void handleMapClickForCatchment(event.latlng);
  });
}

function uniqueCoordinatePoints(points) {
  const seen = new Set();
  const result = [];
  for (const point of points) {
    const lat = Number(point?.lat);
    const lng = Number(point?.lng);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
      continue;
    }
    const key = `${lat.toFixed(6)},${lng.toFixed(6)}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push({ lat, lng });
  }
  return result;
}

function convexHull(points) {
  const unique = uniqueCoordinatePoints(points);
  if (unique.length < 3) {
    return unique;
  }

  const sorted = unique.sort((a, b) => (a.lng === b.lng ? a.lat - b.lat : a.lng - b.lng));
  const cross = (o, a, b) => (a.lng - o.lng) * (b.lat - o.lat) - (a.lat - o.lat) * (b.lng - o.lng);

  const lower = [];
  for (const point of sorted) {
    while (lower.length >= 2 && cross(lower[lower.length - 2], lower[lower.length - 1], point) <= 0) {
      lower.pop();
    }
    lower.push(point);
  }

  const upper = [];
  for (let i = sorted.length - 1; i >= 0; i -= 1) {
    const point = sorted[i];
    while (upper.length >= 2 && cross(upper[upper.length - 2], upper[upper.length - 1], point) <= 0) {
      upper.pop();
    }
    upper.push(point);
  }

  lower.pop();
  upper.pop();
  return [...lower, ...upper];
}

function recomputeCurrentHullFromAccepted() {
  if (!currentCatchment.office) {
    currentCatchment.polygon = [];
    return;
  }
  const source = [
    { lat: currentCatchment.office.lat, lng: currentCatchment.office.lng },
    ...currentCatchment.acceptedPlaces.map((place) => ({ lat: place.lat, lng: place.lng })),
  ];
  const hull = convexHull(source);
  currentCatchment.polygon = hull.length >= MIN_POLYGON_POINTS ? hull : [];
}

function clearPreviewVertexMarkers() {
  for (const marker of previewVertexMarkers) {
    marker.remove();
  }
  previewVertexMarkers = [];
}

function clearCurrentAcceptedMarkers() {
  for (const marker of currentAcceptedMarkers) {
    marker.remove();
  }
  currentAcceptedMarkers = [];
}

function updatePreviewPolygonShape() {
  if (!previewPolygon) {
    return;
  }
  previewPolygon.setLatLngs(currentCatchment.polygon.map((point) => [point.lat, point.lng]));
}

function syncPreviewVertexMarkers() {
  clearPreviewVertexMarkers();
  if (!map || !window.L || !Array.isArray(currentCatchment.polygon) || currentCatchment.polygon.length < MIN_POLYGON_POINTS) {
    return;
  }

  currentCatchment.polygon.forEach((point, index) => {
    const marker = window.L.marker([point.lat, point.lng], {
      draggable: true,
      keyboard: false,
      pane: MAP_PANES.activeVertex,
      icon: window.L.divIcon({
        className: "drive-time-vertex-icon",
        iconSize: [12, 12],
      }),
      title: "Drag to reshape. Right-click to delete.",
    }).addTo(map);

    marker.on("drag", (event) => {
      const latLng = event.target.getLatLng();
      currentCatchment.polygon[index] = { lat: latLng.lat, lng: latLng.lng };
      updatePreviewPolygonShape();
    });

    marker.on("dragend", () => {
      persistEditingSearchChanges("Point moved. Changes saved.");
    });

    marker.on("contextmenu", () => {
      if (currentCatchment.polygon.length <= MIN_POLYGON_POINTS) {
        setStatus("Area needs at least 3 points.", true);
        return;
      }
      currentCatchment.polygon.splice(index, 1);
      updatePreviewPolygonShape();
      syncPreviewVertexMarkers();
      persistEditingSearchChanges("Point deleted. Changes saved.");
    });

    previewVertexMarkers.push(marker);
  });
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

function addPreviewPoint(latlng) {
  if (!Array.isArray(currentCatchment.polygon) || currentCatchment.polygon.length < MIN_POLYGON_POINTS) {
    return;
  }
  const insertIndex = findClosestSegmentInsertIndex(currentCatchment.polygon, latlng);
  currentCatchment.polygon.splice(insertIndex, 0, { lat: latlng.lat, lng: latlng.lng });
  updatePreviewPolygonShape();
  syncPreviewVertexMarkers();
  persistEditingSearchChanges("Point added. Changes saved.");
}

function renderCurrentCatchmentShape() {
  clearCurrentAcceptedMarkers();
  if (!map || !window.L) {
    return;
  }

  for (const place of currentCatchment.acceptedPlaces) {
    const marker = window.L.circleMarker([place.lat, place.lng], {
      radius: 5,
      color: "#1f3c88",
      weight: 2,
      fillColor: "#31b7c8",
      fillOpacity: 0.9,
      pane: MAP_PANES.activeOffice,
    }).addTo(map);
    marker.bindPopup(
      `<strong>${escapeHtml(place.name)}</strong><br/>${escapeHtml(place.postcode || "No postcode")}<br/>${escapeHtml(
        place.durationMinutes
      )} mins`
    );
    currentAcceptedMarkers.push(marker);
  }

  if (previewPolygon) {
    previewPolygon.remove();
    previewPolygon = null;
  }

  if (Array.isArray(currentCatchment.polygon) && currentCatchment.polygon.length >= MIN_POLYGON_POINTS) {
    previewPolygon = window.L.polygon(
      currentCatchment.polygon.map((point) => [point.lat, point.lng]),
      {
        color: "#1d2a42",
        weight: 2,
        opacity: 0.85,
        dashArray: "5 5",
        fillColor: "#9db3df",
        fillOpacity: 0.2,
        pane: MAP_PANES.activeArea,
      }
    ).addTo(map);
    previewPolygon.on("click", (event) => {
      if (event.originalEvent) {
        window.L.DomEvent.stopPropagation(event.originalEvent);
      }
      addPreviewPoint(event.latlng);
    });
  }

  syncPreviewVertexMarkers();
  updateSaveButtonState();
  updateClearAcceptedButtonState();
}

function refreshCurrentCatchmentMapAndList() {
  renderCurrentCatchmentShape();
  renderAcceptedPlacesList();
}

function formatTravelMetaText() {
  const acceptedCount = Number(currentCatchment.acceptedPlaces?.length || 0);
  if (!acceptedCount) {
    return "No accepted places yet.";
  }
  const durations = currentCatchment.acceptedPlaces
    .map((place) => Number(place.durationMinutes))
    .filter((value) => Number.isFinite(value));
  if (!durations.length) {
    return `${acceptedCount} accepted place(s).`;
  }
  const min = Math.min(...durations);
  const max = Math.max(...durations);
  return `${acceptedCount} accepted place(s), travel-time range ${min}-${max} mins.`;
}

function removeAreaLayers(searchId) {
  const layers = areaLayers.get(searchId);
  if (!layers) {
    return;
  }
  if (layers.officeMarker) {
    layers.officeMarker.remove();
  }
  if (layers.polygon) {
    layers.polygon.remove();
  }
  areaLayers.delete(searchId);
}

function syncSearchLayer(search) {
  removeAreaLayers(search.id);
  if (!search.visible || search.archived || !map || !window.L) {
    return;
  }

  const color = getColorByIndex(search.colorIndex);
  const office = sanitizeOffice(search.office);
  const officeLayer = window.L.circleMarker([office.lat, office.lng], {
    radius: 6,
    color: color.stroke,
    weight: 2,
    fillColor: "#fff",
    fillOpacity: 1,
    pane: MAP_PANES.searchOffice,
  }).addTo(map);
  officeLayer.bindPopup(`${escapeHtml(search.name)}<br/>Office: ${escapeHtml(office.name)}`);

  let polygonLayer = null;
  if (Array.isArray(search.polygon) && search.polygon.length >= MIN_POLYGON_POINTS) {
    polygonLayer = window.L.polygon(
      search.polygon.map((point) => [point.lat, point.lng]),
      {
        color: color.stroke,
        weight: 2,
        opacity: 0.95,
        fillColor: color.fill,
        fillOpacity: 0.2,
        pane: MAP_PANES.searchArea,
      }
    ).addTo(map);
    polygonLayer.bindPopup(`${escapeHtml(search.name)} (${search.thresholdMinutes} mins)`);
  }

  areaLayers.set(search.id, { officeMarker: officeLayer, polygon: polygonLayer });
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
  map.setView([search.office.lat, search.office.lng], 11);
}

function normalizePlaceKey(value) {
  return normalizeSearchName(value).toLowerCase();
}

function renderAcceptedPlacesList() {
  if (!acceptedPlacesList || !noAcceptedPlacesMessage) {
    return;
  }

  acceptedPlacesList.innerHTML = "";
  for (const place of currentCatchment.acceptedPlaces) {
    const li = document.createElement("li");
    li.className = "drive-time-search-item";

    const info = document.createElement("div");
    info.className = "drive-time-search-title-wrap";

    const title = document.createElement("span");
    title.className = "drive-time-search-title";
    title.textContent = place.name;

    const subtitle = document.createElement("span");
    subtitle.className = "drive-time-search-subtitle";
    subtitle.textContent = `${place.postcode || "No postcode"} | ${place.durationMinutes} mins${
      Number.isFinite(place.distanceMiles) ? ` | ${place.distanceMiles} mi` : ""
    }`;

    info.append(title, subtitle);

    const actions = document.createElement("div");
    actions.className = "drive-time-search-actions";

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "secondary";
    removeBtn.textContent = "Remove";
    removeBtn.addEventListener("click", () => {
      currentCatchment.acceptedPlaces = currentCatchment.acceptedPlaces.filter((item) => item.key !== place.key);
      recomputeCurrentHullFromAccepted();
      refreshCurrentCatchmentMapAndList();
      persistEditingSearchChanges("Place removed. Changes saved.");
      if (driveTimeMeta) {
        driveTimeMeta.textContent = formatTravelMetaText();
      }
      if (!editingSearchId) {
        setStatus(`Removed ${place.name}.`);
      }
    });

    actions.append(removeBtn);
    li.append(info, actions);
    acceptedPlacesList.appendChild(li);
  }

  noAcceptedPlacesMessage.hidden = currentCatchment.acceptedPlaces.length > 0;
}

function beginEditingSavedSearch(search) {
  if (!search) {
    return;
  }

  editingSearchId = search.id;
  editAddPlacesMode = false;
  currentCatchment = {
    office: sanitizeOffice(search.office),
    thresholdMinutes: clampMinutes(search.thresholdMinutes),
    acceptedPlaces: sanitizeAcceptedPlaces(search.acceptedPlaces),
    polygon: sanitizePolygon(search.polygon),
    departureTimeMode: search.departureTimeMode === "custom" ? "custom" : "default",
    departureTime: search.departureTime || null,
  };
  activeOffice = { ...currentCatchment.office };

  if (!currentCatchment.polygon.length) {
    recomputeCurrentHullFromAccepted();
  }
  if (driveTimeMinutesInput) {
    driveTimeMinutesInput.value = String(currentCatchment.thresholdMinutes);
  }
  if (locationInput) {
    locationInput.value = `${currentCatchment.office.name}${currentCatchment.office.postcode ? ` (${currentCatchment.office.postcode})` : ""}`;
  }
  if (searchNameInput) {
    searchNameInput.value = search.name || currentCatchment.office.name;
  }

  if (currentCatchment.departureTimeMode === "custom" && currentCatchment.departureTime) {
    const custom = new Date(currentCatchment.departureTime);
    if (!Number.isNaN(custom.getTime())) {
      if (departureDateInput) {
        departureDateInput.value = toDateInputValue(custom);
      }
      if (departureTimeInput) {
        departureTimeInput.value = toTimeInputValue(custom);
      }
      setCustomDepartureVisibility(true);
    }
  }

  refreshCurrentCatchmentMapAndList();
  updateOfficeMarker();
  updateSaveButtonState();
  updateEditAddPlacesButtonState();
  if (map) {
    if (currentCatchment.polygon.length >= MIN_POLYGON_POINTS) {
      const bounds = window.L.latLngBounds(currentCatchment.polygon.map((point) => [point.lat, point.lng]));
      if (bounds.isValid()) {
        map.fitBounds(bounds.pad(0.15));
      }
    } else {
      map.setView([currentCatchment.office.lat, currentCatchment.office.lng], 11);
    }
  }

  setStatus(`Editing '${search.name}'. Drag points, click the area edge to add a point, or right-click a point to delete. Use 'Add places' to grow the catchment.`);
  if (driveTimeMeta) {
    driveTimeMeta.textContent = formatTravelMetaText();
  }
}

function renderSavedSearches() {
  if (!savedSearchesList || !noSavedSearchesMessage) {
    return;
  }

  savedSearchesList.innerHTML = "";
  const filteredSearches = getFilteredSavedSearches();
  for (const search of filteredSearches) {
    const color = getColorByIndex(search.colorIndex);
    const li = document.createElement("li");
    li.className = "drive-time-search-item";

    const visibilityLabel = document.createElement("label");
    visibilityLabel.className = "drive-time-search-visibility";

    const selectCheckbox = document.createElement("input");
    selectCheckbox.type = "checkbox";
    selectCheckbox.checked = selectedSavedSearchIds.has(search.id);
    selectCheckbox.setAttribute("aria-label", `Select ${search.name}`);
    selectCheckbox.addEventListener("change", () => {
      if (selectCheckbox.checked) {
        selectedSavedSearchIds.add(search.id);
      } else {
        selectedSavedSearchIds.delete(search.id);
      }
      updateSavedSearchSelectionActions();
    });

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = search.visible;
    checkbox.disabled = search.archived;
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
    subtitle.textContent = `${search.office.name} | ${search.thresholdMinutes} mins | ${search.acceptedPlaces.length} places${
      search.archived ? " | Archived" : ""
    }`;
    titleWrap.append(title, subtitle);

    visibilityLabel.append(selectCheckbox, checkbox, colorDot, titleWrap);

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
        resetCurrentCatchment();
      }
      setStatus(`Removed '${search.name}'.`);
    });

    const archiveBtn = document.createElement("button");
    archiveBtn.type = "button";
    archiveBtn.className = "secondary";
    archiveBtn.textContent = search.archived ? "Restore" : "Archive";
    archiveBtn.addEventListener("click", () => {
      search.archived = !search.archived;
      if (search.archived) {
        search.visible = false;
        if (editingSearchId === search.id) {
          resetCurrentCatchment();
        }
      }
      persistSavedSearches();
      syncSearchLayer(search);
      renderSavedSearches();
      setStatus(search.archived ? `Archived '${search.name}'.` : `Restored '${search.name}'.`);
    });

    actions.append(colourSelect, editBtn, focusBtn, archiveBtn, removeBtn);
    li.append(visibilityLabel, actions);
    savedSearchesList.appendChild(li);
  }

  noSavedSearchesMessage.hidden = filteredSearches.length > 0;
  noSavedSearchesMessage.textContent =
    savedSearchFilter === "archived"
      ? "No archived searches."
      : savedSearchFilter === "all"
        ? "No saved searches yet."
        : "No active searches.";
  updateSavedSearchSelectionActions();
}

function resetCurrentCatchment() {
  editingSearchId = null;
  editAddPlacesMode = false;
  currentCatchment = createEmptyCatchment();
  currentCatchment.thresholdMinutes = resolveDriveTimeMinutes();
  if (searchNameInput) {
    searchNameInput.value = activeOffice ? `${activeOffice.name} catchment` : "";
  }
  clearPreviewVertexMarkers();
  if (previewPolygon) {
    previewPolygon.remove();
    previewPolygon = null;
  }
  clearCurrentAcceptedMarkers();
  updateOfficeMarker();
  renderAcceptedPlacesList();
  updateSaveButtonState();
  updateEditAddPlacesButtonState();
  updateClearAcceptedButtonState();
}

function duplicateSelectedSearches() {
  const selected = getSelectedSavedSearches();
  if (!selected.length) {
    setStatus("Select at least one search to duplicate.", true);
    return;
  }

  const duplicates = selected.map((search, index) =>
    sanitizeSavedSearch(
      {
        ...search,
        id: createSearchId(),
        name: `${search.name} (Copy)`,
        savedAt: new Date().toISOString(),
        visible: false,
      },
      savedSearches.length + index
    )
  ).filter(Boolean);

  if (!duplicates.length) {
    setStatus("Could not duplicate selected searches.", true);
    return;
  }

  savedSearches = [...duplicates, ...savedSearches].slice(0, 100);
  persistSavedSearches();
  renderSavedSearches();
  syncAllSearchLayers();
  setStatus(`Duplicated ${duplicates.length} search${duplicates.length === 1 ? "" : "es"}.`);
}

function mergeSelectedSearches() {
  const selected = getSelectedSavedSearches();
  if (selected.length < 2) {
    setStatus("Select at least two searches to merge.", true);
    return;
  }

  const primary = selected[0];
  const acceptedByKey = new Map();
  const hullSourcePoints = [];

  for (const search of selected) {
    const office = sanitizeOffice(search.office);
    hullSourcePoints.push({ lat: office.lat, lng: office.lng });

    for (const point of sanitizePolygon(search.polygon)) {
      hullSourcePoints.push(point);
    }

    for (const place of sanitizeAcceptedPlaces(search.acceptedPlaces)) {
      if (!acceptedByKey.has(place.key)) {
        acceptedByKey.set(place.key, place);
      }
      hullSourcePoints.push({ lat: place.lat, lng: place.lng });
    }
  }

  const mergedPolygon = convexHull(hullSourcePoints);
  const mergedSearch = sanitizeSavedSearch(
    {
      id: createSearchId(),
      name: `Merged: ${selected.map((search) => search.name).join(" + ")}`,
      office: sanitizeOffice(primary.office),
      thresholdMinutes: Math.max(...selected.map((search) => clampMinutes(search.thresholdMinutes))),
      departureTimeMode: primary.departureTimeMode || "default",
      departureTime: primary.departureTime || null,
      acceptedPlaces: Array.from(acceptedByKey.values()),
      polygon: mergedPolygon,
      qualitySummary: {
        mergedFrom: selected.length,
        acceptedPlaces: acceptedByKey.size,
      },
      savedAt: new Date().toISOString(),
      visible: true,
      archived: false,
      colorIndex: savedSearches.length,
      formattedAddress: primary.formattedAddress || primary.office?.name || "Merged search",
      query: primary.query || primary.office?.name || "Merged search",
    },
    savedSearches.length
  );

  if (!mergedSearch) {
    setStatus("Could not merge selected searches.", true);
    return;
  }

  savedSearches.unshift(mergedSearch);
  savedSearches = savedSearches.slice(0, 100);
  selectedSavedSearchIds.clear();
  persistSavedSearches();
  renderSavedSearches();
  syncAllSearchLayers();
  fitToSearch(mergedSearch);
  setStatus(`Merged ${selected.length} searches into '${mergedSearch.name}'.`);
}

function saveCurrentSearch() {
  if (!currentCatchment.office) {
    setStatus("Set an office location before saving a catchment.", true);
    return;
  }
  if (!canSaveCurrentCatchment()) {
    setStatus("Add at least one accepted place before saving.", true);
    return;
  }

  const explicitName = normalizeSearchName(searchNameInput?.value || "");
  const fallbackName = `${currentCatchment.office.name} ${currentCatchment.thresholdMinutes}-min catchment`;
  const editingTarget = editingSearchId ? savedSearches.find((item) => item.id === editingSearchId) : null;

  const departure = resolveDepartureTime();
  currentCatchment.departureTimeMode = departure.mode;
  currentCatchment.departureTime = departure.mode === "custom" ? departure.iso : null;

  const saved = sanitizeSavedSearch(
    {
      id: editingTarget?.id || createSearchId(),
      name: explicitName || fallbackName,
      office: currentCatchment.office,
      thresholdMinutes: currentCatchment.thresholdMinutes,
      departureTimeMode: currentCatchment.departureTimeMode,
      departureTime: currentCatchment.departureTime,
      acceptedPlaces: currentCatchment.acceptedPlaces,
      polygon: currentCatchment.polygon,
      qualitySummary: {
        acceptedPlaces: currentCatchment.acceptedPlaces.length,
      },
      savedAt: new Date().toISOString(),
      visible: true,
      archived: editingTarget?.archived === true,
      colorIndex: editingTarget?.colorIndex ?? savedSearches.length,
      formattedAddress: currentCatchment.office.name,
      query: currentCatchment.office.name,
    },
    savedSearches.length
  );

  if (!saved) {
    setStatus("Current catchment cannot be saved.", true);
    return;
  }

  if (editingTarget) {
    const index = savedSearches.findIndex((item) => item.id === editingTarget.id);
    if (index >= 0) {
      savedSearches[index] = {
        ...savedSearches[index],
        ...saved,
      };
    }
  } else {
    savedSearches.unshift(saved);
  }

  persistSavedSearches();
  syncSearchLayer(saved);
  renderSavedSearches();
  resetCurrentCatchment();
  setStatus(editingTarget ? `Updated '${saved.name}'.` : `Saved '${saved.name}'.`);
  if (driveTimeMeta) {
    driveTimeMeta.textContent = "";
  }
}

function setAllSearchVisibility(isVisible) {
  const targetSearches = getFilteredSavedSearches().filter((search) => !search.archived);
  if (!targetSearches.length) {
    return;
  }
  for (const search of targetSearches) {
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
    version: 2,
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

function formatReasonMessage(reason, threshold) {
  if (reason === "exceeds_threshold") {
    return `Outside ${threshold}-minute threshold.`;
  }
  if (reason === "missing_route") {
    return "No drivable route found for that point.";
  }
  if (reason === "geocode_failed") {
    return "Could not resolve location details for the clicked point.";
  }
  return "Within threshold.";
}

async function handleMapClickForCatchment(latlng) {
  if (!map || !USE_OFFICE_CATCHMENT_MODE) {
    return;
  }
  if (!currentCatchment.office) {
    setStatus("Set an office location before testing places.", true);
    return;
  }

  const thresholdMinutes = resolveDriveTimeMinutes();
  let departure = null;
  try {
    departure = resolveDepartureTime();
  } catch (error) {
    setStatus(error?.message || "Invalid departure time.", true);
    return;
  }

  const existingKeys = currentCatchment.acceptedPlaces.map((place) => place.key);
  setBusy(true);
  setStatus(`Checking drive time (${thresholdMinutes} mins max, ${departure.label})...`);

  try {
    const response = await authedFetch(OFFICE_CATCHMENT_ENDPOINT, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        office: currentCatchment.office,
        clicked: {
          lat: Number(latlng.lat),
          lng: Number(latlng.lng),
        },
        thresholdMinutes,
        departureTime: departure.iso,
        existingKeys,
      }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data?.error || "Could not check clicked location.");
    }

    const resolved = data?.clickedResolved || {};
    const key = normalizeSearchName(resolved.key || "");
    if (!key) {
      throw new Error("Clicked location response was missing a key.");
    }

    const duplicateLocal = currentCatchment.acceptedPlaces.some((place) => normalizePlaceKey(place.key) === normalizePlaceKey(key));
    const travelMinutes = Number(data?.travel?.durationMinutes);
    const isAccepted = Boolean(data?.accepted) && !duplicateLocal && !Boolean(data?.duplicate);

    if (isAccepted) {
      currentCatchment.acceptedPlaces.push({
        key,
        name: normalizeSearchName(resolved.name || resolved.formattedAddress || "Unknown location"),
        postcode: normalizeSearchName(resolved.postcode || ""),
        formattedAddress: normalizeSearchName(resolved.formattedAddress || ""),
        lat: Number(resolved.lat),
        lng: Number(resolved.lng),
        durationMinutes: Number.isFinite(travelMinutes) ? travelMinutes : null,
        distanceMiles: Number(data?.travel?.distanceMiles),
      });
      recomputeCurrentHullFromAccepted();
      refreshCurrentCatchmentMapAndList();

      if (searchNameInput && !normalizeSearchName(searchNameInput.value)) {
        searchNameInput.value = `${currentCatchment.office.name} catchment`;
      }

      setStatus(
        `Accepted ${resolved.name || resolved.formattedAddress} (${Number.isFinite(travelMinutes) ? `${travelMinutes} mins` : "time unavailable"}).`
      );
      const travelDistanceMiles = Number(data?.travel?.distanceMiles);
      showClickFeedback(
        `Accepted: ${resolved.name || resolved.formattedAddress} (${Number.isFinite(travelMinutes) ? `${travelMinutes} mins` : "time unavailable"}${
          Number.isFinite(travelDistanceMiles) ? `, ${travelDistanceMiles} mi` : ""
        })`,
        "success"
      );
      if (driveTimeMeta) {
        driveTimeMeta.textContent = formatTravelMetaText();
      }
      if (editingSearchId) {
        persistEditingSearchChanges(`Accepted ${resolved.name || resolved.formattedAddress} and saved changes.`);
      }
    } else if (duplicateLocal || data?.duplicate) {
      setStatus("Already added.");
      showClickFeedback("Already added.", "info");
    } else {
      const reasonText = formatReasonMessage(data?.reason, thresholdMinutes);
      setStatus(
        `Rejected ${resolved.name || resolved.formattedAddress || "location"}${
          Number.isFinite(travelMinutes) ? ` (${travelMinutes} mins)` : ""
        }. ${reasonText}`,
        true
      );
      showClickFeedback(
        `Rejected: ${resolved.name || resolved.formattedAddress || "location"}${Number.isFinite(travelMinutes) ? ` (${travelMinutes} mins)` : ""}`,
        "error"
      );
    }
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not evaluate clicked location.", true);
    showClickFeedback("Could not evaluate clicked location.", "error");
  } finally {
    setBusy(false);
  }
}

function clearAcceptedPlaces() {
  currentCatchment.acceptedPlaces = [];
  currentCatchment.polygon = [];
  refreshCurrentCatchmentMapAndList();
  setStatus("Cleared accepted places.");
  if (driveTimeMeta) {
    driveTimeMeta.textContent = "";
  }
}

function updateOfficeMarker() {
  if (!map || !window.L) {
    return;
  }
  if (!currentCatchment.office) {
    if (officeMarker) {
      officeMarker.remove();
      officeMarker = null;
    }
    return;
  }
  if (!officeMarker) {
    officeMarker = window.L.marker([currentCatchment.office.lat, currentCatchment.office.lng], {
      title: `${currentCatchment.office.name}${currentCatchment.office.postcode ? ` (${currentCatchment.office.postcode})` : ""}`,
      keyboard: false,
      pane: MAP_PANES.activeOffice,
    }).addTo(map);
  } else {
    officeMarker.setLatLng([currentCatchment.office.lat, currentCatchment.office.lng]);
  }
  officeMarker.bindPopup(
    `<strong>${escapeHtml(currentCatchment.office.name)}</strong><br/>${escapeHtml(currentCatchment.office.postcode || "")}`
  );
}

async function setOfficeFromQuery(queryValue, sourceLabel = "") {
  const query = normalizeSearchName(queryValue || "");
  if (!query) {
    setStatus("Enter an office location first.", true);
    return;
  }

  setBusy(true);
  setStatus("Resolving office location...");
  try {
    const response = await authedFetch(GEOCODE_BATCH_ENDPOINT, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        queries: [{ id: "office", query }],
      }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(data?.error || "Could not set office location.");
    }

    const points = Array.isArray(data?.points) ? data.points : [];
    const officePoint = points.find((point) => Number.isFinite(Number(point?.lat)) && Number.isFinite(Number(point?.lng)));
    if (!officePoint) {
      throw new Error("Could not find that office location.");
    }

    const formatted = normalizeSearchName(officePoint.formattedAddress || query);
    currentCatchment.office = {
      name: formatted,
      postcode: "",
      lat: Number(officePoint.lat),
      lng: Number(officePoint.lng),
    };
    activeOffice = { ...currentCatchment.office };

    editingSearchId = null;
    clearAcceptedPlaces();
    updateOfficeMarker();
    map.setView([currentCatchment.office.lat, currentCatchment.office.lng], 11);

    if (locationInput) {
      locationInput.value = formatted;
    }
    if (searchNameInput) {
      searchNameInput.value = `${formatted} catchment`;
    }
    setStatus(`Office set to ${formatted}${sourceLabel ? ` from ${sourceLabel}` : ""}. Click map to build catchment.`);
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not set office location.", true);
  } finally {
    setBusy(false);
  }
}

async function setOfficeFromInput() {
  await setOfficeFromQuery(locationInput?.value || "");
}

async function setOfficeFromClientSelection() {
  const selectedId = String(clientOfficeSelect?.value || "").trim();
  if (!selectedId) {
    setStatus("Choose a client location first.", true);
    return;
  }

  const selected = clientOfficeOptions.find((option) => option.id === selectedId);
  if (!selected) {
    setStatus("That client location is no longer available.", true);
    return;
  }

  await setOfficeFromQuery(selected.query, selected.label);
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

    if (locationInput) {
      locationInput.value = "";
    }

    loadGeocodeCache();
    loadSavedSearches();
    await loadClientOfficeOptions();
    if (savedSearchFilterSelect) {
      savedSearchFilterSelect.value = savedSearchFilter;
    }
    ensureDefaultDepartureInputs();
    setCustomDepartureVisibility(false);

    currentCatchment = createEmptyCatchment();
    currentCatchment.thresholdMinutes = resolveDriveTimeMinutes();
    updateOfficeMarker();
    renderAcceptedPlacesList();
    renderSavedSearches();
    setSavedSearchesCollapsed(true);
    syncAllSearchLayers();

    if (searchNameInput) {
      searchNameInput.value = "";
    }

    updateSaveButtonState();
    updateEditAddPlacesButtonState();
    updateClearAcceptedButtonState();
    setPeopleOverlayStatus("Overlay not loaded.");
    setStatus("Set an office location to start building a catchment.");
    document.body.classList.toggle("office-click-mode", USE_OFFICE_CATCHMENT_MODE);
    if (!USE_OFFICE_CATCHMENT_MODE) {
      setStatus("Office catchment mode is disabled in config.", true);
    }
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

saveSearchBtn?.addEventListener("click", () => {
  saveCurrentSearch();
});

duplicateSelectedSearchesBtn?.addEventListener("click", () => {
  duplicateSelectedSearches();
});

mergeSelectedSearchesBtn?.addEventListener("click", () => {
  mergeSelectedSearches();
});

editAddPlacesBtn?.addEventListener("click", () => {
  if (!editingSearchId) {
    return;
  }
  editAddPlacesMode = !editAddPlacesMode;
  updateEditAddPlacesButtonState();
  setStatus(
    editAddPlacesMode
      ? "Add places mode enabled. Click the map to test and add places to this saved search."
      : "Add places mode stopped. Polygon editing remains active."
  );
});

setOfficeBtn?.addEventListener("click", () => {
  void setOfficeFromInput();
});

setOfficeFromClientBtn?.addEventListener("click", () => {
  void setOfficeFromClientSelection();
});

clientOfficeSelect?.addEventListener("change", () => {
  setClientOfficeButtonState();
});

clearAcceptedPlacesBtn?.addEventListener("click", () => {
  clearAcceptedPlaces();
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

overlayCompanionCareCompSelect?.addEventListener("change", () => {
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

toggleSavedSearchesBtn?.addEventListener("click", () => {
  const isCollapsed = Boolean(savedSearchesContent?.hidden);
  setSavedSearchesCollapsed(!isCollapsed);
});

savedSearchFilterSelect?.addEventListener("change", () => {
  savedSearchFilter = String(savedSearchFilterSelect.value || "active");
  renderSavedSearches();
  updateSavedSearchSelectionActions();
});

driveTimeMinutesInput?.addEventListener("change", () => {
  resolveDriveTimeMinutes();
  if (driveTimeMeta) {
    driveTimeMeta.textContent = formatTravelMetaText();
  }
});

locationInput?.addEventListener("keydown", (event) => {
  if (event.key !== "Enter") {
    return;
  }
  event.preventDefault();
  void setOfficeFromInput();
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
