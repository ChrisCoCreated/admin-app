const { requireApiAuth } = require("../_lib/require-api-auth");

const GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json";
const MAX_QUERIES = 400;
const CONCURRENCY = 4;
const geocodeCache = new Map();

function normalizeQuery(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

async function geocodeQuery(query, apiKey, region) {
  const normalized = normalizeQuery(query);
  if (!normalized) {
    return null;
  }
  const cacheKey = normalized.toLowerCase();
  if (geocodeCache.has(cacheKey)) {
    return geocodeCache.get(cacheKey);
  }

  const url = new URL(GOOGLE_GEOCODE_URL);
  url.searchParams.set("address", normalized);
  url.searchParams.set("region", region);
  url.searchParams.set("key", apiKey);

  const response = await fetch(url, {
    headers: {
      Accept: "application/json",
    },
  });
  if (!response.ok) {
    throw new Error(`Geocode request failed (${response.status}).`);
  }

  const data = await response.json();
  if (data.status !== "OK" || !Array.isArray(data.results) || !data.results.length) {
    geocodeCache.set(cacheKey, null);
    return null;
  }

  const first = data.results[0];
  const lat = Number(first?.geometry?.location?.lat);
  const lng = Number(first?.geometry?.location?.lng);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
    geocodeCache.set(cacheKey, null);
    return null;
  }

  const result = {
    lat: Number(lat.toFixed(6)),
    lng: Number(lng.toFixed(6)),
    formattedAddress: String(first.formatted_address || normalized),
  };
  geocodeCache.set(cacheKey, result);
  return result;
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
  const queriesRaw = Array.isArray(req.body?.queries) ? req.body.queries : [];
  const queries = queriesRaw.slice(0, MAX_QUERIES);

  const results = new Array(queries.length).fill(null);
  const errors = [];

  const runOne = async (entry, index) => {
    const query = normalizeQuery(entry?.query);
    if (!query) {
      return;
    }
    try {
      const point = await geocodeQuery(query, apiKey, region);
      if (!point) {
        return;
      }
      results[index] = {
        id: String(entry?.id || String(index)),
        query,
        ...point,
      };
    } catch (error) {
      errors.push(error?.message || "Geocode lookup failed.");
    }
  };

  for (let offset = 0; offset < queries.length; offset += CONCURRENCY) {
    const batch = queries.slice(offset, offset + CONCURRENCY);
    await Promise.all(batch.map((entry, localIndex) => runOne(entry, offset + localIndex)));
  }

  res.status(200).json({
    points: results.filter(Boolean),
    totalRequested: queries.length,
    totalResolved: results.filter(Boolean).length,
    totalErrors: errors.length,
    firstError: errors[0] || null,
  });
};
