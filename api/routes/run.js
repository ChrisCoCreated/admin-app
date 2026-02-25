const { requireApiAuth } = require("../_lib/require-api-auth");

const GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json";
const GOOGLE_ROUTES_URL = "https://routes.googleapis.com/directions/v2:computeRoutes";

function normalizePostcode(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function compactPostcode(value) {
  return normalizePostcode(value).replace(/\s+/g, "");
}

function formatDuration(durationValue) {
  const seconds = Number.parseInt(String(durationValue || "").replace("s", ""), 10);
  if (!Number.isFinite(seconds) || seconds <= 0) {
    return "0 min";
  }

  const hours = Math.floor(seconds / 3600);
  const mins = Math.round((seconds % 3600) / 60);

  if (hours <= 0) {
    return `${Math.max(mins, 1)} min`;
  }

  if (mins <= 0) {
    return `${hours}h`;
  }

  return `${hours}h ${mins}m`;
}

function metersToMiles(distanceMeters) {
  const miles = Number(distanceMeters || 0) * 0.000621371;
  return Number.isFinite(miles) ? miles : 0;
}

async function geocodePostcode(postcode, apiKey, region) {
  const url = new URL(GOOGLE_GEOCODE_URL);
  url.searchParams.set("address", postcode);
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
    throw new Error(`Could not geocode postcode: ${postcode}`);
  }

  const result = data.results[0];
  const location = result.geometry?.location;
  if (!location || typeof location.lat !== "number" || typeof location.lng !== "number") {
    throw new Error(`Missing coordinates for postcode: ${postcode}`);
  }

  return {
    postcode,
    normalizedPostcode: normalizePostcode(postcode),
    formattedAddress: String(result.formatted_address || ""),
    latitude: location.lat,
    longitude: location.lng,
  };
}

async function computeRun(staff, clients, apiKey) {
  const requestBody = {
    origin: {
      location: {
        latLng: {
          latitude: staff.latitude,
          longitude: staff.longitude,
        },
      },
    },
    destination: {
      location: {
        latLng: {
          latitude: staff.latitude,
          longitude: staff.longitude,
        },
      },
    },
    intermediates: clients.map((client) => ({
      location: {
        latLng: {
          latitude: client.latitude,
          longitude: client.longitude,
        },
      },
    })),
    travelMode: "DRIVE",
    routingPreference: "TRAFFIC_UNAWARE",
    optimizeWaypointOrder: true,
    units: "METRIC",
    languageCode: "en-GB",
  };

  const response = await fetch(GOOGLE_ROUTES_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      "X-Goog-Api-Key": apiKey,
      "X-Goog-FieldMask": [
        "routes.distanceMeters",
        "routes.duration",
        "routes.legs.distanceMeters",
        "routes.legs.duration",
        "routes.optimizedIntermediateWaypointIndex",
      ].join(","),
    },
    body: JSON.stringify(requestBody),
  });

  const data = await response.json();
  if (!response.ok) {
    const detail = data?.error?.message || `Routes request failed (${response.status}).`;
    throw new Error(detail);
  }

  const route = Array.isArray(data?.routes) ? data.routes[0] : null;
  if (!route || !Array.isArray(route.legs) || !route.legs.length) {
    throw new Error("No route returned from Google Routes API.");
  }

  return route;
}

function buildResponse(staff, clients, route) {
  const optimizedIndex = Array.isArray(route.optimizedIntermediateWaypointIndex)
    ? route.optimizedIntermediateWaypointIndex
    : clients.map((_, index) => index);

  const orderedClients = optimizedIndex.map((index) => clients[index]).filter(Boolean);
  const stopNames = [staff.normalizedPostcode, ...orderedClients.map((client) => client.normalizedPostcode), staff.normalizedPostcode];

  const legs = route.legs.map((leg, index) => {
    const from = stopNames[index] || "";
    const to = stopNames[index + 1] || "";
    const distanceMeters = Number(leg.distanceMeters || 0);
    const miles = metersToMiles(distanceMeters);

    return {
      legNumber: index + 1,
      from,
      to,
      distanceMeters,
      distanceMiles: Number(miles.toFixed(2)),
      duration: String(leg.duration || "0s"),
      durationText: formatDuration(leg.duration),
    };
  });

  const totalDistanceMeters = Number(route.distanceMeters || legs.reduce((sum, leg) => sum + leg.distanceMeters, 0));
  const totalMiles = metersToMiles(totalDistanceMeters);

  return {
    run: {
      staffStart: {
        postcode: staff.normalizedPostcode,
        formattedAddress: staff.formattedAddress,
        latitude: staff.latitude,
        longitude: staff.longitude,
      },
      orderedClients: orderedClients.map((client) => ({
        postcode: client.normalizedPostcode,
        formattedAddress: client.formattedAddress,
        latitude: client.latitude,
        longitude: client.longitude,
      })),
      legs,
      totalDistanceMeters,
      totalDistanceMiles: Number(totalMiles.toFixed(2)),
      totalDuration: String(route.duration || "0s"),
      totalDurationText: formatDuration(route.duration),
    },
  };
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res))) {
    return;
  }

  const apiKey = String(process.env.GOOGLE_MAPS_API_KEY || "").trim();
  if (!apiKey) {
    res.status(500).json({ error: "Server missing GOOGLE_MAPS_API_KEY." });
    return;
  }

  const region = String(process.env.GOOGLE_MAPS_REGION || "gb").trim().toLowerCase() || "gb";
  const staffPostcode = normalizePostcode(req.body?.staffPostcode);
  const rawClientPostcodes = Array.isArray(req.body?.clientPostcodes) ? req.body.clientPostcodes : [];

  if (!staffPostcode) {
    res.status(400).json({ error: "Staff postcode is required." });
    return;
  }

  const dedupedClients = [];
  const seen = new Set();
  for (const input of rawClientPostcodes) {
    const normalized = normalizePostcode(input);
    const key = compactPostcode(normalized);
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    dedupedClients.push(normalized);
  }

  if (!dedupedClients.length) {
    res.status(400).json({ error: "Add at least one client postcode." });
    return;
  }

  if (dedupedClients.length > 23) {
    res.status(400).json({ error: "Maximum 23 client postcodes per run." });
    return;
  }

  try {
    const [staff, ...clients] = await Promise.all([
      geocodePostcode(staffPostcode, apiKey, region),
      ...dedupedClients.map((postcode) => geocodePostcode(postcode, apiKey, region)),
    ]);

    const route = await computeRun(staff, clients, apiKey);
    const payload = buildResponse(staff, clients, route);

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json(payload);
  } catch (error) {
    res.status(400).json({
      error: "Could not build run.",
      detail: error && error.message ? error.message : String(error),
    });
  }
};
