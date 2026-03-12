const { requireApiAuth } = require("../_lib/require-api-auth");

const GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json";
const GOOGLE_ROUTES_URL = "https://routes.googleapis.com/directions/v2:computeRoutes";
const EARTH_RADIUS_METERS = 6371000;
const TARGET_MINUTES = 20;
const ANGLE_STEP_DEGREES = 15;
const MIN_RADIUS_METERS = 2000;
const BASE_MAX_RADIUS_METERS = 50000;
const ROUTE_CALL_CONCURRENCY = 2;
const DEPARTURE_LEAD_SECONDS = 120;
const DEFAULT_DEPARTURE_WEEKDAY = 3; // Wednesday
const DEFAULT_DEPARTURE_HOUR = 10;
const RADIAL_MAX_ATTEMPTS = 3;
const RADIAL_SEED_DISTANCE_METERS = 16093.4; // 10 miles

function normalizeLocationQuery(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function toRadians(degrees) {
  return (degrees * Math.PI) / 180;
}

function toDegrees(radians) {
  return (radians * 180) / Math.PI;
}

function parseDurationSeconds(durationValue) {
  const seconds = Number.parseInt(String(durationValue || "").replace("s", ""), 10);
  return Number.isFinite(seconds) && seconds > 0 ? seconds : 0;
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

function buildBearings() {
  const bearings = [];
  for (let degrees = 0; degrees < 360; degrees += ANGLE_STEP_DEGREES) {
    bearings.push(degrees);
  }
  return bearings;
}

function destinationPoint(origin, bearingDegrees, distanceMeters) {
  const angularDistance = distanceMeters / EARTH_RADIUS_METERS;
  const bearing = toRadians(bearingDegrees);
  const lat1 = toRadians(origin.latitude);
  const lon1 = toRadians(origin.longitude);

  const sinLat1 = Math.sin(lat1);
  const cosLat1 = Math.cos(lat1);
  const sinAngular = Math.sin(angularDistance);
  const cosAngular = Math.cos(angularDistance);

  const lat2 = Math.asin(sinLat1 * cosAngular + cosLat1 * sinAngular * Math.cos(bearing));
  const lon2 = lon1 + Math.atan2(
    Math.sin(bearing) * sinAngular * cosLat1,
    cosAngular - sinLat1 * Math.sin(lat2)
  );

  return {
    latitude: toDegrees(lat2),
    longitude: toDegrees(lon2),
  };
}

async function geocodeLocation(query, apiKey, region) {
  const url = new URL(GOOGLE_GEOCODE_URL);
  url.searchParams.set("address", query);
  url.searchParams.set("region", region);
  url.searchParams.set("key", apiKey);

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
    throw new Error("Could not find that location.");
  }

  const firstResult = data.results[0];
  const location = firstResult.geometry?.location;
  if (typeof location?.lat !== "number" || typeof location?.lng !== "number") {
    throw new Error("Google geocoding did not return coordinates.");
  }

  return {
    query,
    formattedAddress: String(firstResult.formatted_address || query),
    latitude: location.lat,
    longitude: location.lng,
  };
}

async function computeDurationsForDestinations(origin, destinations, apiKey, departureTime) {
  if (!Array.isArray(destinations) || destinations.length === 0) {
    return { durations: [], failures: [] };
  }

  const durations = new Array(destinations.length).fill(0);
  const failures = [];

  const runRouteCall = async (destination, index) => {
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
      units: "METRIC",
    };

    const response = await fetch(GOOGLE_ROUTES_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        "X-Goog-Api-Key": apiKey,
        "X-Goog-FieldMask": "routes.duration",
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
    durations[index] = parseDurationSeconds(route?.duration);
  };

  for (let offset = 0; offset < destinations.length; offset += ROUTE_CALL_CONCURRENCY) {
    const batch = destinations.slice(offset, offset + ROUTE_CALL_CONCURRENCY);
    await Promise.all(
      batch.map((destination, localIndex) =>
        runRouteCall(destination, offset + localIndex).catch((error) => {
          durations[offset + localIndex] = 0;
          if (error?.message) {
            failures.push(error.message);
          }
        })
      )
    );
  }

  return { durations, failures };
}

async function solveDirectionalDistances({
  center,
  bearings,
  bandMinSeconds,
  bandMaxSeconds,
  bandTargetSeconds,
  minRadiusMeters,
  maxRadiusMeters,
  initialRadiusMeters,
  apiKey,
  departureTime,
}) {
  const distances = new Array(bearings.length).fill(0);
  const latestDurations = new Array(bearings.length).fill(0);
  const finalInBandByDirection = new Array(bearings.length).fill(false);
  const unresolvedByDirection = new Array(bearings.length).fill(true);
  const attemptsByDirection = new Array(bearings.length).fill(0);
  const statusByDirection = new Array(bearings.length).fill("unresolved");

  async function solveRadial(index) {
    const bearing = bearings[index];
    let distance = clamp(initialRadiusMeters, minRadiusMeters, maxRadiusMeters);
    let bestDuration = 0;
    let bestDistance = distance;
    let attemptCount = 0;

    for (let attempt = 0; attempt < RADIAL_MAX_ATTEMPTS; attempt += 1) {
      attemptCount = attempt + 1;
      const destination = destinationPoint(center, bearing, distance);
      const sample = await computeDurationsForDestinations(center, [destination], apiKey, departureTime);
      const durationSeconds = Number(sample.durations?.[0] || 0);
      if (!durationSeconds) {
        distance = clamp(distance * 1.2, minRadiusMeters, maxRadiusMeters);
        continue;
      }

      bestDuration = durationSeconds;
      bestDistance = distance;
      if (durationSeconds >= bandMinSeconds && durationSeconds <= bandMaxSeconds) {
        attemptsByDirection[index] = attemptCount;
        latestDurations[index] = durationSeconds;
        distances[index] = distance;
        finalInBandByDirection[index] = true;
        unresolvedByDirection[index] = false;
        statusByDirection[index] = "in_band";
        return;
      }

      const ratio = bandTargetSeconds / Math.max(durationSeconds, 1);
      const scaledRatio = 1 + (ratio - 1) * 0.8;
      distance = clamp(distance * scaledRatio, minRadiusMeters, maxRadiusMeters);
    }

    // Keep last valid sample as diagnostic, but mark as out-of-band.
    attemptsByDirection[index] = attemptCount;
    if (bestDuration > 0) {
      latestDurations[index] = bestDuration;
      distances[index] = bestDistance;
      unresolvedByDirection[index] = false;
      statusByDirection[index] = bestDuration < bandMinSeconds ? "below_band" : "above_band";
    } else {
      statusByDirection[index] = "unresolved";
    }
  }

  for (let offset = 0; offset < bearings.length; offset += ROUTE_CALL_CONCURRENCY) {
    const batch = bearings.slice(offset, offset + ROUTE_CALL_CONCURRENCY);
    await Promise.all(batch.map((_, localIndex) => solveRadial(offset + localIndex)));
  }

  return {
    distances,
    durations: latestDurations,
    inBandByDirection: finalInBandByDirection,
    unresolvedByDirection,
    attemptsByDirection,
    statusByDirection,
  };
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
  const locationQuery = normalizeLocationQuery(req.body?.location);
  if (!locationQuery) {
    res.status(400).json({ error: "Location is required." });
    return;
  }

  const requestedMinutes = Number(req.body?.minutes || TARGET_MINUTES);
  const minutes = Number.isFinite(requestedMinutes) && requestedMinutes > 0 ? Math.round(requestedMinutes) : TARGET_MINUTES;
  const bandMaxMinutes = minutes;
  const bandMinMinutes = Number((minutes * 0.9).toFixed(1));
  const bandTargetMinutes = (bandMinMinutes + bandMaxMinutes) / 2;
  const bandMinSeconds = Math.round(bandMinMinutes * 60);
  const bandMaxSeconds = Math.round(bandMaxMinutes * 60);
  const bandTargetSeconds = Math.round(bandTargetMinutes * 60);
  const fallbackSpeedMetersPerSecond = 16.7; // ~60 km/h fallback
  const dynamicMaxRadiusMeters = Math.max(
    BASE_MAX_RADIUS_METERS,
    Math.round(bandMaxSeconds * fallbackSpeedMetersPerSecond * 1.35)
  );
  let departureTime = "";
  try {
    departureTime = resolveDepartureTime(req.body?.departureTime);
  } catch (error) {
    res.status(400).json({ error: error?.message || "Invalid departureTime." });
    return;
  }
  const bearings = buildBearings();

  try {
    const center = await geocodeLocation(locationQuery, apiKey, region);
    const solved = await solveDirectionalDistances({
      center,
      bearings,
      bandMinSeconds,
      bandMaxSeconds,
      bandTargetSeconds,
      minRadiusMeters: MIN_RADIUS_METERS,
      maxRadiusMeters: dynamicMaxRadiusMeters,
      initialRadiusMeters: RADIAL_SEED_DISTANCE_METERS,
      apiKey,
      departureTime,
    });
    const smoothedFinalDistances = solved.distances;
    const polygon = bearings
      .map((bearing, index) => {
        if (!solved.inBandByDirection[index]) {
          return null;
        }
        const point = destinationPoint(center, bearing, smoothedFinalDistances[index]);
        return {
          lat: Number(point.latitude.toFixed(6)),
          lng: Number(point.longitude.toFixed(6)),
        };
      })
      .filter(Boolean);

    if (polygon.length < 3) {
      res.status(422).json({
        error: "Insufficient resolved directions in the requested drive-time band.",
      });
      return;
    }

    const successfulDurations = solved.durations.filter((duration) => Number(duration) > 0);
    const inBandDurations = solved.durations.filter((duration, index) => solved.inBandByDirection[index] && Number(duration) > 0);
    const unresolvedDirections = solved.unresolvedByDirection.filter(Boolean).length;
    const inBandDirections = solved.inBandByDirection.filter(Boolean).length;
    const attemptsHistogram = { "0": 0, "1": 0, "2": 0, "3": 0 };
    for (const attempts of solved.attemptsByDirection) {
      const key = String(Math.max(0, Math.min(RADIAL_MAX_ATTEMPTS, Number(attempts) || 0)));
      attemptsHistogram[key] = (attemptsHistogram[key] || 0) + 1;
    }
    const radialDiagnostics = bearings.map((bearing, index) => ({
      bearingDegrees: bearing,
      attempts: Number(solved.attemptsByDirection[index] || 0),
      status: solved.statusByDirection[index],
      durationMinutes:
        Number(solved.durations[index] || 0) > 0
          ? Math.round((Number(solved.durations[index]) / 60) * 10) / 10
          : null,
      distanceMiles:
        Number(solved.distances[index] || 0) > 0
          ? Math.round((Number(solved.distances[index]) * 0.000621371) * 100) / 100
          : null,
    }));

    const minMinutes = inBandDurations.length
      ? Math.round((Math.min(...inBandDurations) / 60) * 10) / 10
      : null;
    const maxMinutes = inBandDurations.length
      ? Math.round((Math.max(...inBandDurations) / 60) * 10) / 10
      : null;

    res.status(200).json({
      center: {
        lat: Number(center.latitude.toFixed(6)),
        lng: Number(center.longitude.toFixed(6)),
      },
      formattedAddress: center.formattedAddress,
      minutes,
      polygon,
      quality: {
        sampledDirections: bearings.length,
        successfulDirections: successfulDurations.length,
        inBandDirections,
        unresolvedDirections,
        minDurationMinutes: minMinutes,
        maxDurationMinutes: maxMinutes,
        bandMinMinutes,
        bandMaxMinutes,
        radialAttemptsHistogram: attemptsHistogram,
        radialDiagnostics,
      },
    });
  } catch (error) {
    res.status(500).json({
      error: error?.message || "Could not compute drive-time area.",
    });
  }
};
