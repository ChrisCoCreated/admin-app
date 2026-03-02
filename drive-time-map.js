import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const signOutBtn = document.getElementById("signOutBtn");
const locationInput = document.getElementById("locationInput");
const drawDriveTimeBtn = document.getElementById("drawDriveTimeBtn");
const driveTimeStatus = document.getElementById("driveTimeStatus");
const driveTimeMeta = document.getElementById("driveTimeMeta");
const driveTimeMapRoot = document.getElementById("driveTimeMap");

const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
const DRIVE_TIME_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/maps/drive-time` : "/api/maps/drive-time";

let map = null;
let centerMarker = null;
let driveTimePolygon = null;

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

function initMap() {
  if (!driveTimeMapRoot || !window.L) {
    return;
  }

  map = window.L.map(driveTimeMapRoot, {
    zoomControl: true,
    attributionControl: true,
  }).setView([51.5072, -0.1276], 10);

  window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
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

function renderDriveTimeOnMap(payload) {
  if (!map || !window.L) {
    throw new Error("Map is not available.");
  }

  const center = payload?.center;
  const polygonPoints = Array.isArray(payload?.polygon) ? payload.polygon : [];
  if (!center || typeof center.lat !== "number" || typeof center.lng !== "number" || !polygonPoints.length) {
    throw new Error("Drive-time response was missing map coordinates.");
  }

  if (centerMarker) {
    centerMarker.remove();
  }
  if (driveTimePolygon) {
    driveTimePolygon.remove();
  }

  centerMarker = window.L.marker([center.lat, center.lng], {
    title: payload?.formattedAddress || "Selected location",
  }).addTo(map);

  driveTimePolygon = window.L.polygon(
    polygonPoints.map((point) => [point.lat, point.lng]),
    {
      color: "#1f3c88",
      weight: 2,
      opacity: 0.95,
      fillColor: "#31b7c8",
      fillOpacity: 0.22,
    }
  ).addTo(map);

  const polygonBounds = driveTimePolygon.getBounds();
  if (polygonBounds.isValid()) {
    map.fitBounds(polygonBounds.pad(0.15));
  } else {
    map.setView([center.lat, center.lng], 11);
  }
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

    renderDriveTimeOnMap(data);
    setStatus(`Showing 20-minute drive-time area for ${data.formattedAddress || location}.`);
    if (driveTimeMeta) {
      driveTimeMeta.textContent = formatQualityText(data.quality);
    }
  } catch (error) {
    console.error(error);
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
