const { readCarersDirectoryData } = require("../_lib/directory-source");
const { requireApiAuth } = require("../_lib/require-api-auth");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "care_manager", "operations"] }))) {
    return;
  }

  try {
    const directory = await readCarersDirectoryData();
    const q = String(req.query.q || "").trim().toLowerCase();
    const limitRaw = Number(req.query.limit || "250");
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(limitRaw, 1000)) : 250;

    const filtered = q
      ? directory.carers.filter((carer) => {
          return (
            String(carer.name || "").toLowerCase().includes(q) ||
            String(carer.id || "").toLowerCase().includes(q) ||
            String(carer.postcode || "").toLowerCase().includes(q)
          );
        })
      : directory.carers;

    res.setHeader("Cache-Control", "private, max-age=30");
    res.setHeader("X-Client-Source", directory.source);

    res.status(200).json({
      carers: filtered.slice(0, limit),
      total: filtered.length,
      warnings: directory.warnings,
    });
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error?.message || String(error),
    });
  }
};
