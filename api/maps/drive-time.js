const { requireApiAuth } = require("../_lib/require-api-auth");

const GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json";
const GOOGLE_ROUTES_URL = "https://routes.googleapis.com/directions/v2:computeRoutes";
const EARTH_RADIUS_METERS = 6371000;
const TARGET_MINUTES = 20;
const ANGLE_STEP_DEGREES = 15;
const BASE_RADIUS_METERS = 30000;
const MIN_RADIUS_METERS = 2000;
const MAX_RADIUS_METERS = 50000;
const ROUTE_CALL_CONCURRENCY = 2;
const DEPARTURE_LEAD_SECONDS = 120;

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

function smoothDistances(values) {
  if (!Array.isArray(values) || values.length < 3) {
    return values;
  }

  return values.map((value, index) => {
    const prev = values[(index - 1 + values.length) % values.length];
    const next = values[(index + 1) % values.length];
    return clamp((prev + value + next) / 3, MIN_RADIUS_METERS, MAX_RADIUS_METERS);
  });
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

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "care_manager", "operations"] }))) {
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
  const targetSeconds = minutes * 60;
  const fallbackSpeedMetersPerSecond = 16.7; // ~60 km/h fallback
  const fallbackDistanceMeters = clamp(targetSeconds * fallbackSpeedMetersPerSecond, MIN_RADIUS_METERS, MAX_RADIUS_METERS);
  const departureTime = new Date(Date.now() + DEPARTURE_LEAD_SECONDS * 1000).toISOString();
  const bearings = buildBearings();

  try {
    const center = await geocodeLocation(locationQuery, apiKey, region);

    const firstPassDestinations = bearings.map((bearing) => destinationPoint(center, bearing, BASE_RADIUS_METERS));
    const firstPass = await computeDurationsForDestinations(
      center,
      firstPassDestinations,
      apiKey,
      departureTime
    );
    const firstPassDurations = firstPass.durations;
    const firstPassDistances = bearings.map((_, index) => {
      const durationSeconds = Number(firstPassDurations[index] || 0);
      if (!durationSeconds) {
        return fallbackDistanceMeters;
      }
      return clamp(BASE_RADIUS_METERS * (targetSeconds / durationSeconds), MIN_RADIUS_METERS, MAX_RADIUS_METERS);
    });

    const smoothedFirstPassDistances = smoothDistances(firstPassDistances);
    const secondPassDestinations = bearings.map((bearing, index) =>
      destinationPoint(center, bearing, smoothedFirstPassDistances[index])
    );
    const secondPass = await computeDurationsForDestinations(
      center,
      secondPassDestinations,
      apiKey,
      departureTime
    );
    const secondPassDurations = secondPass.durations;

    const finalDistances = bearings.map((_, index) => {
      const durationSeconds = Number(secondPassDurations[index] || 0);
      const seedDistance = smoothedFirstPassDistances[index];
      if (!durationSeconds) {
        return seedDistance;
      }
      return clamp(seedDistance * (targetSeconds / durationSeconds), MIN_RADIUS_METERS, MAX_RADIUS_METERS);
    });
    const smoothedFinalDistances = smoothDistances(finalDistances);
    const polygon = bearings.map((bearing, index) => {
      const point = destinationPoint(center, bearing, smoothedFinalDistances[index]);
      return {
        lat: Number(point.latitude.toFixed(6)),
        lng: Number(point.longitude.toFixed(6)),
      };
    });

    const successfulDurations = secondPassDurations.filter((duration) => Number(duration) > 0);
    const failedMessages = [...firstPass.failures, ...secondPass.failures];
    if (successfulDurations.length === 0) {
      const firstFailure = failedMessages[0] || "No valid routes returned by Google Routes API.";
      res.status(502).json({
        error: `Drive-time calculation failed: ${firstFailure}`,
      });
      return;
    }

    const minMinutes = successfulDurations.length
      ? Math.round((Math.min(...successfulDurations) / 60) * 10) / 10
      : null;
    const maxMinutes = successfulDurations.length
      ? Math.round((Math.max(...successfulDurations) / 60) * 10) / 10
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
        minDurationMinutes: minMinutes,
        maxDurationMinutes: maxMinutes,
        failedDirectionCalls: failedMessages.length,
      },
    });
  } catch (error) {
    res.status(500).json({
      error: error?.message || "Could not compute drive-time area.",
    });
  }
};
