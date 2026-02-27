const { readMarketingPhotos } = require("../_lib/photos-source");
const { requireApiAuth } = require("../_lib/require-api-auth");

let cachedPayload = null;
let cacheExpiresAt = 0;
let inFlightRead = null;

function getCacheTtlMs() {
  const configured = Number(process.env.MARKETING_PHOTOS_CACHE_TTL_MS || 120000);
  if (!Number.isFinite(configured) || configured < 0) {
    return 120000;
  }
  return Math.floor(configured);
}

async function getPhotosWithCache() {
  const now = Date.now();
  if (cachedPayload && now < cacheExpiresAt) {
    return cachedPayload;
  }

  if (inFlightRead) {
    return inFlightRead;
  }

  inFlightRead = (async () => {
    const photos = await readMarketingPhotos();
    const payload = {
      photos,
      total: photos.length,
      defaultView: "all",
    };
    cachedPayload = payload;
    cacheExpiresAt = Date.now() + getCacheTtlMs();
    return payload;
  })();

  try {
    return await inFlightRead;
  } finally {
    inFlightRead = null;
  }
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "marketing"] }))) {
    return;
  }

  try {
    const payload = await getPhotosWithCache();
    res.setHeader("Cache-Control", "private, max-age=30");
    res.status(200).json(payload);
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error && error.message ? error.message : String(error),
    });
  }
};
