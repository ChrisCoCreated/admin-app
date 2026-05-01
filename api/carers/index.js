const { readCarersDirectoryData } = require("../_lib/directory-source");
const { listSharePointColleagues } = require("../_lib/colleagues-source");
const { requireApiAuth } = require("../_lib/require-api-auth");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res))) {
    return;
  }

  try {
    const directory = await readCarersDirectoryData();
    let colleaguesByOneTouchId = new Map();
    let colleaguesAvailable = true;
    const warnings = Array.isArray(directory.warnings) ? [...directory.warnings] : [];
    try {
      const colleagues = await listSharePointColleagues();
      colleaguesByOneTouchId = new Map(
        colleagues
          .filter((colleague) => String(colleague?.oneTouchId || "").trim())
          .map((colleague) => [String(colleague.oneTouchId).trim(), colleague])
      );
    } catch (error) {
      colleaguesAvailable = false;
      warnings.push(`Could not load Colleagues list. ${error?.message || String(error)}`);
    }

    const carers = directory.carers.map((carer) => {
      const oneTouchId = String(carer?.id || "").trim();
      const colleague = oneTouchId ? colleaguesByOneTouchId.get(oneTouchId) : null;
      return {
        ...carer,
        inSharePoint: colleaguesAvailable ? Boolean(colleague) : false,
        sharePointItemId: colleague?.itemId || "",
        sharePointState: colleaguesAvailable ? (colleague ? "present" : "missing") : "unknown",
      };
    });

    const q = String(req.query.q || "").trim().toLowerCase();
    const area = String(req.query.area || "").trim().toLowerCase();
    const careComp = String(req.query.care_comp || req.query.careComp || "").trim().toLowerCase();
    const limitRaw = Number(req.query.limit || "250");
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(limitRaw, 1000)) : 250;

    const filtered = carers.filter((carer) => {
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
      warnings,
    });
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error?.message || String(error),
    });
  }
};
