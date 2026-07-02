const { readCarersDirectoryData, readDirectoryData } = require("../_lib/directory-source");
const { buildGeocodeUrl, describeGeocodeFailure, normalizeRegion } = require("../_lib/google-geocode");
const { requireApiAuth } = require("../_lib/require-api-auth");

const GOOGLE_ROUTES_URL = "https://routes.googleapis.com/directions/v2:computeRoutes";
const ROUTE_CALL_CONCURRENCY = 3;
const GEOCODE_CALL_CONCURRENCY = 4;
const DEPARTURE_LEAD_SECONDS = 120;
const DEFAULT_DEPARTURE_WEEKDAY = 3; // Wednesday
const DEFAULT_DEPARTURE_HOUR = 10;
const ACTIVE_PENDING_STATUSES = new Set(["active", "pending"]);

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeComparable(value) {
  return normalizeText(value).toLowerCase();
}

function normalizeStatus(value) {
  return normalizeComparable(value);
}

function normalizeAudience(value) {
  const normalized = normalizeComparable(value);
  if (normalized === "associate" || normalized === "associates") {
    return "associates";
  }
  if (normalized === "client" || normalized === "clients") {
    return "clients";
  }
  return "clients";
}

function getDefaultDepartureTimeIso() {
  const now = new Date();
  const nowMs = now.getTime();
  const weekday = now.getDay();
  const dayOffset = (DEFAULT_DEPARTURE_WEEKDAY - weekday + 7) % 7;
  let candidate = new Date(now.getFullYear(), now.getMonth(), now.getDate() + dayOffset, DEFAULT_DEPARTURE_HOUR, 0, 0, 0, 0);
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

function parseDurationSeconds(durationValue) {
  const seconds = Number.parseInt(String(durationValue || "").replace("s", ""), 10);
  return Number.isFinite(seconds) && seconds > 0 ? seconds : 0;
}

function formatDurationSeconds(totalSeconds) {
  const safeSeconds = Number(totalSeconds || 0);
  if (!Number.isFinite(safeSeconds) || safeSeconds <= 0) {
    return "0 min";
  }

  const totalMinutes = Math.max(1, Math.round(safeSeconds / 60));
  if (totalMinutes < 60) {
    return `${totalMinutes} min`;
  }

  const hours = Math.floor(totalMinutes / 60);
  const mins = totalMinutes % 60;
  return mins ? `${hours}h ${mins}m` : `${hours}h`;
}

function metersToMiles(distanceMeters) {
  const miles = Number(distanceMeters || 0) * 0.000621371;
  return Number.isFinite(miles) ? miles : 0;
}

function getClientArea(client) {
  return normalizeText(client?.area || "Unassigned");
}

function getCarerArea(carer) {
  return normalizeText(carer?.area || "");
}

function buildClientAddress(client) {
  const parts = [client?.address, client?.town, client?.county, client?.postcode]
    .map((value) => normalizeText(value))
    .filter(Boolean);
  if (parts.length) {
    return parts.join(", ");
  }
  return normalizeText(client?.postcode || client?.location || "");
}

function buildDestinationRows(clients, carers, selectedArea, audience) {
  const selectedAreaKey = normalizeComparable(selectedArea);
  const rows = [];

  if (audience === "clients") {
    for (const client of clients) {
      const status = normalizeStatus(client?.status);
      if (!ACTIVE_PENDING_STATUSES.has(status)) {
        continue;
      }
      const area = getClientArea(client);
      if (normalizeComparable(area) !== selectedAreaKey) {
        continue;
      }
      const address = buildClientAddress(client);
      const postcode = normalizeText(client?.postcode);
      if (!address && !postcode) {
        continue;
      }
      rows.push({
        type: "client",
        name: normalizeText(client?.name || "Unnamed client"),
        id: normalizeText(client?.id),
        status,
        area,
        postcode,
        query: address || postcode,
      });
    }
  }

  if (audience === "associates") {
    for (const carer of carers) {
      const status = normalizeStatus(carer?.status);
      if (!ACTIVE_PENDING_STATUSES.has(status)) {
        continue;
      }
      const area = getCarerArea(carer);
      if (normalizeComparable(area) !== selectedAreaKey) {
        continue;
      }
      const postcode = normalizeText(carer?.postcode);
      if (!postcode) {
        continue;
      }
      rows.push({
        type: "associate",
        name: normalizeText(carer?.name || "Unnamed associate"),
        id: normalizeText(carer?.id),
        status,
        area,
        postcode,
        query: postcode,
      });
    }
  }

  return rows;
}

async function geocodeLocation(query, apiKey, region) {
  const url = buildGeocodeUrl(query, apiKey, region);

  const response = await fetch(url, {
    headers: {
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`Geocoding request failed (${response.status}).`);
  }

  const data = await response.json();
  if (data.status !== "OK" || !Array.isArray(data.results) || !data.results.length) {
    throw new Error(describeGeocodeFailure(data, query));
  }

  const result = data.results[0];
  const location = result.geometry?.location;
  if (!location || typeof location.lat !== "number" || typeof location.lng !== "number") {
    throw new Error(`Missing coordinates for location: ${query}`);
  }

  return {
    query,
    formattedAddress: normalizeText(result.formatted_address || query),
    latitude: location.lat,
    longitude: location.lng,
  };
}

async function geocodeUniqueDestinations(rows, apiKey, region) {
  const geocodeCache = new Map();
  const uniqueQueries = Array.from(new Set(rows.map((row) => normalizeComparable(row.query)).filter(Boolean)));
  let nextIndex = 0;

  async function worker() {
    while (nextIndex < uniqueQueries.length) {
      const index = nextIndex;
      nextIndex += 1;
      const key = uniqueQueries[index];
      const row = rows.find((item) => normalizeComparable(item.query) === key);
      try {
        geocodeCache.set(key, await geocodeLocation(row.query, apiKey, region));
      } catch (error) {
        geocodeCache.set(key, {
          error: error?.message || "Could not geocode this location.",
        });
      }
    }
  }

  await Promise.all(Array.from({ length: Math.min(GEOCODE_CALL_CONCURRENCY, uniqueQueries.length) }, worker));
  return geocodeCache;
}

async function computeRouteTravel(origin, destination, apiKey, departureTime) {
  const requestBody = {
    origin: {
      location: {
        latLng: {
          latitude: origin.latitude,
          longitude: origin.longitude,
        },
      },
    },
    destination: {
      location: {
        latLng: {
          latitude: destination.latitude,
          longitude: destination.longitude,
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
  const durationSeconds = parseDurationSeconds(route?.duration);
  if (!durationSeconds) {
    throw new Error("No route returned from Google Routes API.");
  }

  const distanceMeters = Number(route?.distanceMeters || 0);
  return {
    durationSeconds,
    durationMinutes: Math.round((durationSeconds / 60) * 10) / 10,
    durationText: formatDurationSeconds(durationSeconds),
    distanceMeters,
    distanceMiles: Number(metersToMiles(distanceMeters).toFixed(1)),
  };
}

async function attachTravelTimes(origin, rows, geocodeCache, apiKey, departureTime) {
  const results = new Array(rows.length);
  let nextIndex = 0;

  async function worker() {
    while (nextIndex < rows.length) {
      const index = nextIndex;
      nextIndex += 1;
      const row = rows[index];
      const geocoded = geocodeCache.get(normalizeComparable(row.query));
      if (!geocoded || geocoded.error) {
        results[index] = {
          ...row,
          travel: null,
          reason: geocoded?.error || "Could not geocode this location.",
        };
        continue;
      }

      try {
        results[index] = {
          ...row,
          formattedAddress: geocoded.formattedAddress,
          travel: await computeRouteTravel(origin, geocoded, apiKey, departureTime),
          reason: "",
        };
      } catch (error) {
        results[index] = {
          ...row,
          formattedAddress: geocoded.formattedAddress,
          travel: null,
          reason: error?.message || "Could not calculate route.",
        };
      }
    }
  }

  await Promise.all(Array.from({ length: Math.min(ROUTE_CALL_CONCURRENCY, rows.length) }, worker));
  return results;
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

  const originQuery = normalizeText(req.body?.origin || req.body?.postcode);
  const selectedArea = normalizeText(req.body?.area);
  const audience = normalizeAudience(req.body?.audience || req.body?.type);
  if (!originQuery) {
    res.status(400).json({ error: "Postcode is required." });
    return;
  }
  if (!selectedArea) {
    res.status(400).json({ error: "Area is required." });
    return;
  }

  let departureTime = "";
  try {
    departureTime = resolveDepartureTime(req.body?.departureTime);
  } catch (error) {
    res.status(400).json({ error: error?.message || "Invalid departureTime." });
    return;
  }

  const region = normalizeRegion(process.env.GOOGLE_MAPS_REGION);

  try {
    const [directory, carersDirectory, origin] = await Promise.all([
      readDirectoryData(),
      readCarersDirectoryData(),
      geocodeLocation(originQuery, apiKey, region),
    ]);
    const rows = buildDestinationRows(directory.clients || [], carersDirectory.carers || [], selectedArea, audience);
    const geocodeCache = await geocodeUniqueDestinations(rows, apiKey, region);
    const results = await attachTravelTimes(origin, rows, geocodeCache, apiKey, departureTime);
    const sorted = results.sort((a, b) => {
      const aDuration = Number(a.travel?.durationSeconds || Number.POSITIVE_INFINITY);
      const bDuration = Number(b.travel?.durationSeconds || Number.POSITIVE_INFINITY);
      if (aDuration !== bDuration) {
        return aDuration - bDuration;
      }
      return String(a.name || "").localeCompare(String(b.name || ""), undefined, { sensitivity: "base" });
    });

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      origin: {
        query: originQuery,
        formattedAddress: origin.formattedAddress,
        latitude: origin.latitude,
        longitude: origin.longitude,
      },
      area: selectedArea,
      audience,
      departureTime,
      counts: {
        total: sorted.length,
        withTravel: sorted.filter((item) => item.travel).length,
        clients: sorted.filter((item) => item.type === "client").length,
        associates: sorted.filter((item) => item.type === "associate").length,
      },
      results: sorted,
      warnings: [...(directory.warnings || []), ...(carersDirectory.warnings || [])],
    });
  } catch (error) {
    res.status(500).json({
      error: "Could not calculate journey times.",
      detail: error?.message || String(error),
    });
  }
};
