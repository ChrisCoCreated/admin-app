const fs = require("fs/promises");
const path = require("path");
const { requireApiAuth } = require("../_lib/require-api-auth");

const COMPANY_ASSETS_DIR = path.join(process.cwd(), "assets", "company_marketing_assets");
const ALLOWED_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".svg", ".webp", ".gif", ".avif"]);

function toTitle(fileName) {
  const base = String(fileName || "").replace(/\.[^.]+$/, "");
  return base
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function toAssetId(fileName) {
  return `__company_asset__${String(fileName || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")}`;
}

function getMediaType(fileName) {
  const ext = path.extname(String(fileName || "")).toLowerCase();
  if ([".mp4", ".mov", ".webm"].includes(ext)) {
    return "video";
  }
  return "image";
}

async function readCompanyAssets() {
  const entries = await fs.readdir(COMPANY_ASSETS_DIR, { withFileTypes: true });
  return entries
    .filter((entry) => entry.isFile())
    .map((entry) => entry.name)
    .filter((fileName) => ALLOWED_EXTENSIONS.has(path.extname(fileName).toLowerCase()))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }))
    .map((fileName) => {
      const assetPath = `./assets/company_marketing_assets/${encodeURIComponent(fileName).replace(/%2F/gi, "/")}`;
      return {
        id: toAssetId(fileName),
        title: toTitle(fileName),
        client: "Company Assets",
        imageUrl: assetPath,
        mediaUrl: assetPath,
        attachmentUrl: assetPath,
        mediaType: getMediaType(fileName),
        fileName,
      };
    });
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "marketing", "photo_layout"] }))) {
    return;
  }

  try {
    const assets = await readCompanyAssets();
    res.setHeader("Cache-Control", "private, max-age=60");
    res.status(200).json({
      assets,
      total: assets.length,
    });
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error && error.message ? error.message : String(error),
    });
  }
};
