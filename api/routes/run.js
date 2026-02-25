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

function normalizeLocationQuery(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
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

function parseDurationSeconds(durationValue) {
  const seconds = Number.parseInt(String(durationValue || "").replace("s", ""), 10);
  return Number.isFinite(seconds) && seconds > 0 ? seconds : 0;
}

function parseRate(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : 0;
}

function parseThreshold(value) {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) {
    return null;
  }
  const parsed = Number(trimmed);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : null;
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
    throw new Error(`Could not geocode location: ${query}`);
  }

  const result = data.results[0];
  const location = result.geometry?.location;
  if (!location || typeof location.lat !== "number" || typeof location.lng !== "number") {
    throw new Error(`Missing coordinates for location: ${query}`);
  }

  return {
    query,
    stopLabel: query,
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
  const stopNames = [staff.normalizedPostcode, ...orderedClients.map((client) => client.stopLabel), staff.normalizedPostcode];

  const legs = route.legs.map((leg, index) => {
    const from = stopNames[index] || "";
    const to = stopNames[index + 1] || "";
    const distanceMeters = Number(leg.distanceMeters || 0);
    const durationSeconds = parseDurationSeconds(leg.duration);
    const miles = metersToMiles(distanceMeters);

    return {
      legNumber: index + 1,
      from,
      to,
      distanceMeters,
      distanceMiles: Number(miles.toFixed(2)),
      duration: String(leg.duration || "0s"),
      durationSeconds,
      durationText: formatDuration(leg.duration),
    };
  });

  const totalDistanceMeters = Number(route.distanceMeters || legs.reduce((sum, leg) => sum + leg.distanceMeters, 0));
  const totalMiles = metersToMiles(totalDistanceMeters);
  const totalDurationSeconds = parseDurationSeconds(route.duration);
  const maxDistanceMiles = parseThreshold(process.env.MAX_DISTANCE);
  const maxTimeMinutes = parseThreshold(process.env.MAX_TIME);
  const travelPayRate = parseRate(process.env.TRAVEL_PAY);
  const perMileRate = parseRate(process.env.PER_MILE);
  const costMode = maxDistanceMiles !== null ? "distance" : maxTimeMinutes !== null ? "time" : "distance";
  const distanceThresholdMiles = maxDistanceMiles ?? 0;
  const timeThresholdSeconds = Math.round((maxTimeMinutes ?? 0) * 60);

  const homeLegIndexes = [0, legs.length - 1].filter((index, idx, arr) => index >= 0 && arr.indexOf(index) === idx);
  const runLegIndexes = legs
    .map((_, index) => index)
    .filter((index) => !homeLegIndexes.includes(index));
  const payableLegs = [];
  let paidDistanceMeters = 0;
  let paidDurationSeconds = 0;

  for (const index of homeLegIndexes) {
    const leg = legs[index];
    if (!leg) {
      continue;
    }

    let legPaidDistanceMeters = 0;
    let legPaidDurationSeconds = 0;

    if (costMode === "distance") {
      const thresholdMeters = distanceThresholdMiles / 0.000621371;
      legPaidDistanceMeters = Math.max(0, leg.distanceMeters - thresholdMeters);
      if (leg.distanceMeters > 0) {
        legPaidDurationSeconds = leg.durationSeconds * (legPaidDistanceMeters / leg.distanceMeters);
      }
    } else {
      legPaidDurationSeconds = Math.max(0, leg.durationSeconds - timeThresholdSeconds);
      if (leg.durationSeconds > 0) {
        legPaidDistanceMeters = leg.distanceMeters * (legPaidDurationSeconds / leg.durationSeconds);
      }
    }

    paidDistanceMeters += legPaidDistanceMeters;
    paidDurationSeconds += legPaidDurationSeconds;
    payableLegs.push({
      legNumber: leg.legNumber,
      paidDistanceMiles: Number(metersToMiles(legPaidDistanceMeters).toFixed(2)),
      paidDurationMinutes: Number((legPaidDurationSeconds / 60).toFixed(1)),
    });
  }

  const paidDistanceMiles = metersToMiles(paidDistanceMeters);
  const paidDurationHours = paidDurationSeconds / 3600;
  const homeTimeCost = paidDurationHours * travelPayRate;
  const homeMileageCost = paidDistanceMiles * perMileRate;
  const exceptionalHomeTotal = homeTimeCost + homeMileageCost;

  let runDistanceMeters = 0;
  let runDurationSeconds = 0;
  for (const index of runLegIndexes) {
    const leg = legs[index];
    if (!leg) {
      continue;
    }
    runDistanceMeters += leg.distanceMeters;
    runDurationSeconds += leg.durationSeconds;
  }

  const runDistanceMiles = metersToMiles(runDistanceMeters);
  const runDurationHours = runDurationSeconds / 3600;
  const runTimeCost = runDurationHours * travelPayRate;
  const runMileageCost = runDistanceMiles * perMileRate;
  const runTravelTotal = runTimeCost + runMileageCost;
  const grandTotal = exceptionalHomeTotal + runTravelTotal;

  return {
    run: {
      staffStart: {
        postcode: staff.normalizedPostcode,
        formattedAddress: staff.formattedAddress,
        latitude: staff.latitude,
        longitude: staff.longitude,
      },
      orderedClients: orderedClients.map((client) => ({
        query: client.query,
        stopLabel: client.stopLabel,
        formattedAddress: client.formattedAddress,
        latitude: client.latitude,
        longitude: client.longitude,
      })),
      legs,
      totalDistanceMeters,
      totalDistanceMiles: Number(totalMiles.toFixed(2)),
      totalDuration: String(route.duration || "0s"),
      totalDurationText: formatDuration(route.duration),
      totalDurationSeconds,
      cost: {
        mode: costMode,
        thresholds: {
          maxDistanceMiles,
          maxTimeMinutes,
        },
        rates: {
          travelPayPerHour: travelPayRate,
          perMile: perMileRate,
        },
        homeTravel: {
          paidDistanceMiles: Number(paidDistanceMiles.toFixed(2)),
          paidDurationHours: Number(paidDurationHours.toFixed(2)),
          paidDurationSeconds: Math.round(paidDurationSeconds),
        },
        runTravel: {
          distanceMiles: Number(runDistanceMiles.toFixed(2)),
          durationHours: Number(runDurationHours.toFixed(2)),
          durationSeconds: Math.round(runDurationSeconds),
        },
        components: {
          homeTimeCost: Number(homeTimeCost.toFixed(2)),
          homeMileageCost: Number(homeMileageCost.toFixed(2)),
          runTimeCost: Number(runTimeCost.toFixed(2)),
          runMileageCost: Number(runMileageCost.toFixed(2)),
        },
        totals: {
          exceptionalHomeTotal: Number(exceptionalHomeTotal.toFixed(2)),
          runTravelTotal: Number(runTravelTotal.toFixed(2)),
          grandTotal: Number(grandTotal.toFixed(2)),
        },
        payableLegs,
        totalCost: Number(grandTotal.toFixed(2)),
      },
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
  const rawClientLocations = Array.isArray(req.body?.clientLocations)
    ? req.body.clientLocations
    : Array.isArray(req.body?.clientPostcodes)
      ? req.body.clientPostcodes
      : [];

  if (!staffPostcode) {
    res.status(400).json({ error: "Staff postcode is required." });
    return;
  }

  const dedupedClients = [];
  const seen = new Set();
  for (const input of rawClientLocations) {
    const query = normalizeLocationQuery(typeof input === "string" ? input : input?.query);
    const stopLabel = normalizeLocationQuery(typeof input === "string" ? input : input?.label || input?.query);
    const key = query.toLowerCase();
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    dedupedClients.push({
      query,
      stopLabel: stopLabel || query,
    });
  }

  if (!dedupedClients.length) {
    res.status(400).json({ error: "Add at least one client stop." });
    return;
  }

  if (dedupedClients.length > 23) {
    res.status(400).json({ error: "Maximum 23 client postcodes per run." });
    return;
  }

  try {
    const [staff, ...clients] = await Promise.all([
      geocodeLocation(staffPostcode, apiKey, region).then((item) => ({
        ...item,
        normalizedPostcode: normalizePostcode(staffPostcode),
      })),
      ...dedupedClients.map((client) =>
        geocodeLocation(client.query, apiKey, region).then((item) => ({
          ...item,
          stopLabel: client.stopLabel,
        }))
      ),
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
