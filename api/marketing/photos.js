const { readMarketingPhotos } = require("../_lib/photos-source");
const { requireApiAuth } = require("../_lib/require-api-auth");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "marketing"] }))) {
    return;
  }

  try {
    const photos = await readMarketingPhotos();
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      photos,
      total: photos.length,
      defaultView: "all",
    });
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error && error.message ? error.message : String(error),
    });
  }
};
