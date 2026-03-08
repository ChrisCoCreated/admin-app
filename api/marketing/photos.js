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

function normalizeClientName(value) {
  return String(value || "").trim() || "Unassigned";
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
    const clientsOnly = String(req.query.clientsOnly || "").trim().toLowerCase();
    const clientFilter = String(req.query.client || "").trim();

    if (clientsOnly === "1" || clientsOnly === "true") {
      const counts = new Map();
      for (const photo of payload.photos) {
        const clientName = normalizeClientName(photo?.client);
        counts.set(clientName, (counts.get(clientName) || 0) + 1);
      }
      const clients = Array.from(counts.entries())
        .map(([name, count]) => ({ name, count }))
        .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
      res.setHeader("Cache-Control", "private, max-age=30");
      res.status(200).json({
        clients,
        totalClients: clients.length,
        defaultView: "clients",
      });
      return;
    }

    if (clientFilter) {
      const normalizedClient = normalizeClientName(clientFilter);
      const photos = payload.photos.filter((photo) => normalizeClientName(photo?.client) === normalizedClient);
      res.setHeader("Cache-Control", "private, max-age=30");
      res.status(200).json({
        photos,
        total: photos.length,
        client: normalizedClient,
        defaultView: "client",
      });
      return;
    }

    res.setHeader("Cache-Control", "private, max-age=30");
    res.status(200).json(payload);
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error && error.message ? error.message : String(error),
    });
  }
};
