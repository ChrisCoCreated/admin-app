const assert = require("node:assert/strict");
const test = require("node:test");

const {
  buildGeocodeUrl,
  describeGeocodeFailure,
  getApiKeyFingerprint,
  getSafeGeocodeDiagnostics,
  normalizeGeocodeAddress,
} = require("../api/_lib/google-geocode");

test("normalizes compact UK postcodes for Google geocoding", () => {
  assert.equal(normalizeGeocodeAddress("CT11QX", "gb"), "CT1 1QX");
  assert.equal(normalizeGeocodeAddress("ct20 3qp", "gb"), "CT20 3QP");
});

test("constrains UK geocode lookups to Great Britain", () => {
  const url = buildGeocodeUrl("CT11QX", "test-key", "gb");

  assert.equal(url.searchParams.get("address"), "CT1 1QX");
  assert.equal(url.searchParams.get("components"), "country:GB");
  assert.equal(url.searchParams.get("region"), "gb");
  assert.equal(url.searchParams.get("language"), "en-GB");
});

test("includes Google geocode status in failure details", () => {
  const message = describeGeocodeFailure(
    {
      status: "REQUEST_DENIED",
      error_message: "API key not valid.",
    },
    "CT1 1QX"
  );

  assert.equal(message, "Could not geocode location: CT1 1QX (REQUEST_DENIED: API key not valid.)");
});

test("adds safe diagnostics without exposing the Google API key", () => {
  const apiKey = "secret-google-key";
  const diagnostics = getSafeGeocodeDiagnostics("CT11QX", apiKey, "gb");
  const message = describeGeocodeFailure({ status: "REQUEST_DENIED" }, "CT11QX", diagnostics);

  assert.equal(diagnostics.keyFingerprint, getApiKeyFingerprint(apiKey));
  assert.match(message, /address=CT1 1QX/);
  assert.match(message, /components=country:GB/);
  assert.match(message, /key=[a-f0-9]{12}/);
  assert.doesNotMatch(message, /secret-google-key/);
});
