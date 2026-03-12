const { readCarersDirectoryData } = require("../_lib/directory-source");
const { requireApiAuth } = require("../_lib/require-api-auth");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (
    !(await requireApiAuth(req, res, {
      allowedRoles: [
        "admin",
        "care_manager",
        "operations",
        "hr_only",
        "hr_clients",
        "time_hr",
        "time_hr_clients",
      ],
    }))
  ) {
    return;
  }

  try {
    const directory = await readCarersDirectoryData();
    const q = String(req.query.q || "").trim().toLowerCase();
    const area = String(req.query.area || "").trim().toLowerCase();
    const careComp = String(req.query.care_comp || req.query.careComp || "").trim().toLowerCase();
    const limitRaw = Number(req.query.limit || "250");
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(limitRaw, 1000)) : 250;

    const filtered = directory.carers.filter((carer) => {
      if (q) {
        const matchesQuery =
          String(carer.name || "").toLowerCase().includes(q) ||
          String(carer.id || "").toLowerCase().includes(q) ||
          String(carer.postcode || "").toLowerCase().includes(q) ||
          String(carer.area || "").toLowerCase().includes(q);
        if (!matchesQuery) {
          return false;
        }
      }

      if (area) {
        if (String(carer.area || "").toLowerCase() !== area) {
          return false;
        }
      }

      if (careComp) {
        if (String(carer.careCompanionshipTag || "").toLowerCase() !== careComp) {
          return false;
        }
      }

      return true;
    });

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
