const { requireApiAuth } = require("../_lib/require-api-auth");

const GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json";
const GOOGLE_ROUTES_URL = "https://routes.googleapis.com/directions/v2:computeRoutes";
const DEPARTURE_LEAD_SECONDS = 120;
const DEFAULT_DEPARTURE_WEEKDAY = 3; // Wednesday
const DEFAULT_DEPARTURE_HOUR = 10;
const DEFAULT_THRESHOLD_MINUTES = 20;
const MIN_THRESHOLD_MINUTES = 1;
const MAX_THRESHOLD_MINUTES = 240;

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizePostcode(value) {
  return normalizeText(value).toUpperCase().replace(/\s+/g, "");
}

function getDefaultDepartureTimeIso() {
  const now = new Date();
  const nowMs = now.getTime();
  const weekday = now.getDay();
  const dayOffset = (DEFAULT_DEPARTURE_WEEKDAY - weekday + 7) % 7;
  let candidate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + dayOffset, DEFAULT_DEPARTURE_HOUR, 0, 0, 0);
  if (candidate.getTime() <= nowMs + DEPARTURE_LEAD_SECONDS * 1000) {
    candidate = new Date(
      candidate.getFullYear(),
      candidate.getMonth(),
      candidate.getDate() + 7,
      DEFAULT_DEPARTURE_HOUR,
      0,
      0,
      0
    );
  }
  return candidate.toISOString();
}

function resolveDepartureTime(requestedValue) {
  const requested = String(requestedValue || "").trim();
  if (!requested) {
    return getDefaultDepartureTimeIso();
  }
  const date = new Date(requested);
  if (Number.isNaN(date.getTime())) {
    throw new Error("Invalid departureTime.");
  }
  if (date.getTime() <= Date.now() + DEPARTURE_LEAD_SECONDS * 1000) {
    throw new Error("departureTime must be set in the future.");
  }
  return date.toISOString();
}

function parseLatLng(raw, fieldName) {
  const lat = Number(raw?.lat);
  const lng = Number(raw?.lng);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
    throw new Error(`${fieldName} lat/lng are required.`);
  }
  return { lat, lng };
}

function pickAddressPart(components, wantedType) {
  const match = Array.isArray(components)
    ? components.find((part) => Array.isArray(part?.types) && part.types.includes(wantedType))
    : null;
  return normalizeText(match?.long_name || "");
}

async function reverseGeocode(lat, lng, apiKey, region) {
  const url = new URL(GOOGLE_GEOCODE_URL);
  url.searchParams.set("latlng", `${lat},${lng}`);
  url.searchParams.set("language", "en-GB");
  url.searchParams.set("region", region);
  url.searchParams.set("key", apiKey);

  const response = await fetch(url, {
    headers: {
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Reverse geocoding request failed (${response.status}).`);
  }

  const data = await response.json();
  if (data.status !== "OK" || !Array.isArray(data.results) || !data.results.length) {
    return null;
  }

  const first = data.results[0];
  const components = Array.isArray(first.address_components) ? first.address_components : [];
  const locality =
    pickAddressPart(components, "postal_town") ||
    pickAddressPart(components, "locality") ||
    pickAddressPart(components, "administrative_area_level_3") ||
    pickAddressPart(components, "administrative_area_level_2");
  const postcode = pickAddressPart(components, "postal_code");

  return {
    formattedAddress: normalizeText(first.formatted_address || ""),
    name: locality || normalizeText(first.formatted_address || "Unknown location"),
    postcode,
    lat,
    lng,
  };
}

function parseDurationSeconds(durationValue) {
  const seconds = Number.parseInt(String(durationValue || "").replace("s", ""), 10);
  return Number.isFinite(seconds) && seconds > 0 ? seconds : 0;
}

async function computeRouteTravel(office, clicked, apiKey, departureTime) {
  const requestBody = {
    origin: {
      location: {
        latLng: {
          latitude: office.lat,
          longitude: office.lng,
        },
      },
    },
    destination: {
      location: {
        latLng: {
          latitude: clicked.lat,
          longitude: clicked.lng,
        },
      },
    },
    travelMode: "DRIVE",
    routingPreference: "TRAFFIC_AWARE",
    departureTime,
    languageCode: "en-GB",
    units: "IMPERIAL",
  };

  const response = await fetch(GOOGLE_ROUTES_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      "X-Goog-Api-Key": apiKey,
      "X-Goog-FieldMask": "routes.duration,routes.distanceMeters",
    },
    body: JSON.stringify(requestBody),
  });

  const rawBody = await response.text();
  if (!response.ok) {
    let detail = `Routes request failed (${response.status}).`;
    try {
      const parsed = JSON.parse(rawBody);
      detail = parsed?.error?.message || detail;
    } catch {
      // Keep default detail.
    }
    throw new Error(detail);
  }

  let data = {};
  try {
    data = JSON.parse(rawBody);
  } catch {
    data = {};
  }
  const route = Array.isArray(data?.routes) ? data.routes[0] : null;
  if (!route) {
    return null;
  }

  const durationSeconds = parseDurationSeconds(route.duration);
  const distanceMeters = Number(route.distanceMeters || 0);
  if (!durationSeconds) {
    return null;
  }

  const durationMinutes = Math.round((durationSeconds / 60) * 10) / 10;
  const distanceMiles = distanceMeters > 0 ? Math.round(distanceMeters * 0.000621371 * 10) / 10 : null;
  return {
    durationMinutes,
    durationText: `${durationMinutes} mins`,
    distanceMiles,
  };
}

function buildPlaceKey(clickedResolved) {
  const postcodeNorm = normalizePostcode(clickedResolved?.postcode);
  if (postcodeNorm) {
    return `pc:${postcodeNorm}`;
  }

  const locality = normalizeText(clickedResolved?.name || clickedResolved?.formattedAddress || "").toLowerCase();
  const roundedLat = Number(clickedResolved?.lat).toFixed(4);
  const roundedLng = Number(clickedResolved?.lng).toFixed(4);
  const localityPart = locality || "unknown";
  return `loc:${localityPart}:${roundedLat},${roundedLng}`;
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (
    !(await requireApiAuth(req, res, {
      allowedRoles: [
        "admin",
        "care_manager",
        "operations",
        "time_only",
        "time_clients",
        "time_hr",
        "time_hr_clients",
      ],
    }))
  ) {
    return;
  }

  const apiKey = String(process.env.GOOGLE_MAPS_API_KEY || "").trim();
  if (!apiKey) {
    res.status(500).json({ error: "Server missing GOOGLE_MAPS_API_KEY." });
    return;
  }

  const region = String(process.env.GOOGLE_MAPS_REGION || "gb").trim().toLowerCase() || "gb";

  let office = null;
  let clicked = null;
  try {
    office = parseLatLng(req.body?.office, "office");
    clicked = parseLatLng(req.body?.clicked, "clicked");
  } catch (error) {
    res.status(400).json({ error: error?.message || "Invalid coordinates." });
    return;
  }

  const thresholdMinutes = clamp(
    Math.round(Number(req.body?.thresholdMinutes) || DEFAULT_THRESHOLD_MINUTES),
    MIN_THRESHOLD_MINUTES,
    MAX_THRESHOLD_MINUTES
  );

  let departureTime = "";
  try {
    departureTime = resolveDepartureTime(req.body?.departureTime);
  } catch (error) {
    res.status(400).json({ error: error?.message || "Invalid departureTime." });
    return;
  }

  const existingKeys = Array.isArray(req.body?.existingKeys)
    ? new Set(req.body.existingKeys.map((key) => normalizeText(key)).filter(Boolean))
    : new Set();

  let clickedResolved = null;
  try {
    clickedResolved = await reverseGeocode(clicked.lat, clicked.lng, apiKey, region);
  } catch (error) {
    res.status(200).json({
      office: {
        name: normalizeText(req.body?.office?.name || "Office"),
        postcode: normalizeText(req.body?.office?.postcode || ""),
        lat: office.lat,
        lng: office.lng,
      },
      clickedResolved: {
        name: "Unknown location",
        postcode: "",
        formattedAddress: "",
        lat: clicked.lat,
        lng: clicked.lng,
        key: `loc:unknown:${clicked.lat.toFixed(4)},${clicked.lng.toFixed(4)}`,
      },
      travel: null,
      thresholdMinutes,
      accepted: false,
      duplicate: false,
      reason: "geocode_failed",
    });
    return;
  }

  if (!clickedResolved) {
    res.status(200).json({
      office: {
        name: normalizeText(req.body?.office?.name || "Office"),
        postcode: normalizeText(req.body?.office?.postcode || ""),
        lat: office.lat,
        lng: office.lng,
      },
      clickedResolved: {
        name: "Unknown location",
        postcode: "",
        formattedAddress: "",
        lat: clicked.lat,
        lng: clicked.lng,
        key: `loc:unknown:${clicked.lat.toFixed(4)},${clicked.lng.toFixed(4)}`,
      },
      travel: null,
      thresholdMinutes,
      accepted: false,
      duplicate: false,
      reason: "geocode_failed",
    });
    return;
  }

  const key = buildPlaceKey(clickedResolved);
  clickedResolved.key = key;

  const travel = await computeRouteTravel(office, clicked, apiKey, departureTime).catch(() => null);
  if (!travel) {
    res.status(200).json({
      office: {
        name: normalizeText(req.body?.office?.name || "Office"),
        postcode: normalizeText(req.body?.office?.postcode || ""),
        lat: office.lat,
        lng: office.lng,
      },
      clickedResolved,
      travel: null,
      thresholdMinutes,
      accepted: false,
      duplicate: false,
      reason: "missing_route",
    });
    return;
  }

  const duplicate = existingKeys.has(key);
  const accepted = !duplicate && travel.durationMinutes <= thresholdMinutes;
  const reason = accepted ? "within_threshold" : duplicate ? "within_threshold" : "exceeds_threshold";

  res.status(200).json({
    office: {
      name: normalizeText(req.body?.office?.name || "Office"),
      postcode: normalizeText(req.body?.office?.postcode || ""),
      lat: office.lat,
      lng: office.lng,
    },
    clickedResolved,
    travel,
    thresholdMinutes,
    accepted,
    duplicate,
    reason,
  });
};
