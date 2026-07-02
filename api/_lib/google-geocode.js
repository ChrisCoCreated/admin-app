const crypto = require("node:crypto");

const GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json";

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeRegion(value) {
  return String(value || "gb").trim().toLowerCase() || "gb";
}

function countryComponentForRegion(region) {
  const normalized = normalizeRegion(region);
  if (normalized === "gb" || normalized === "uk") {
    return "country:GB";
  }
  return "";
}

function formatUkPostcode(value) {
  const compact = normalizeText(value).toUpperCase().replace(/\s+/g, "");
  if (!/^[A-Z]{1,2}\d[A-Z\d]?\d[A-Z]{2}$/.test(compact)) {
    return "";
  }
  return `${compact.slice(0, -3)} ${compact.slice(-3)}`;
}

function normalizeGeocodeAddress(value, region) {
  const normalized = normalizeText(value);
  if (!normalized) {
    return "";
  }

  const postcode = countryComponentForRegion(region) ? formatUkPostcode(normalized) : "";
  return postcode || normalized;
}

function buildGeocodeUrl(query, apiKey, region) {
  const normalizedRegion = normalizeRegion(region);
  const url = new URL(GOOGLE_GEOCODE_URL);
  url.searchParams.set("address", normalizeGeocodeAddress(query, normalizedRegion));
  url.searchParams.set("language", "en-GB");
  url.searchParams.set("region", normalizedRegion);
  const components = countryComponentForRegion(normalizedRegion);
  if (components) {
    url.searchParams.set("components", components);
  }
  url.searchParams.set("key", apiKey);
  return url;
}

function getApiKeyFingerprint(apiKey) {
  const normalized = normalizeText(apiKey);
  if (!normalized) {
    return "missing";
  }
  return crypto.createHash("sha256").update(normalized).digest("hex").slice(0, 12);
}

function getSafeGeocodeDiagnostics(query, apiKey, region) {
  const normalizedRegion = normalizeRegion(region);
  return {
    address: normalizeGeocodeAddress(query, normalizedRegion),
    region: normalizedRegion,
    components: countryComponentForRegion(normalizedRegion) || "",
    keyFingerprint: getApiKeyFingerprint(apiKey),
  };
}

function describeGeocodeFailure(data, query, diagnostics = null) {
  const status = normalizeText(data?.status) || "UNKNOWN";
  const apiMessage = normalizeText(data?.error_message);
  const suffix = apiMessage ? `: ${apiMessage}` : "";
  const diagnosticText = diagnostics
    ? `; address=${diagnostics.address || "unknown"}; region=${diagnostics.region || "unknown"}; components=${diagnostics.components || "none"}; key=${diagnostics.keyFingerprint || "unknown"}`
    : "";
  return `Could not geocode location: ${query} (${status}${suffix}${diagnosticText})`;
}

module.exports = {
  buildGeocodeUrl,
  describeGeocodeFailure,
  getApiKeyFingerprint,
  getSafeGeocodeDiagnostics,
  normalizeGeocodeAddress,
  normalizeRegion,
};
