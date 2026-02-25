const { readDirectoryData } = require("../_lib/directory-source");
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
    const directory = await readDirectoryData();
    const q = String(req.query.q || "").trim().toLowerCase();
    const limitRaw = Number(req.query.limit || "250");
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(limitRaw, 1000)) : 250;

    const filtered = q
      ? directory.clients.filter((client) => {
          return (
            String(client.name || "").toLowerCase().includes(q) ||
            String(client.id || "").toLowerCase().includes(q) ||
            String(client.postcode || "").toLowerCase().includes(q)
          );
        })
      : directory.clients;

    res.setHeader("Cache-Control", "no-store");
    res.setHeader("X-Client-Source", directory.source);

    res.status(200).json({
      clients: filtered.slice(0, limit),
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
