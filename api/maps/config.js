const { requireApiAuth } = require("../_lib/require-api-auth");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "care_manager", "operations"] }))) {
    return;
  }

  const apiKey = String(process.env.GOOGLE_MAPS_BROWSER_API_KEY || process.env.GOOGLE_MAPS_API_KEY || "").trim();
  if (!apiKey) {
    res.status(500).json({ error: "Server missing GOOGLE_MAPS_BROWSER_API_KEY or GOOGLE_MAPS_API_KEY." });
    return;
  }

  const region = String(process.env.GOOGLE_MAPS_REGION || "gb").trim().toLowerCase() || "gb";
  res.status(200).json({
    mapsJavaScriptApiKey: apiKey,
    region,
  });
};
