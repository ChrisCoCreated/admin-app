const { readClients } = require("../_lib/clients-source");
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
    const { clients, source, graphError } = await readClients();
    const q = String(req.query.q || "").trim().toLowerCase();
    const limitRaw = Number(req.query.limit || "100");
    const limit = Number.isFinite(limitRaw) ? Math.max(1, Math.min(limitRaw, 500)) : 100;

    const filtered = q
      ? clients.filter((client) => {
          return client.name.toLowerCase().includes(q) || client.id.toLowerCase().includes(q);
        })
      : clients;

    res.setHeader("Cache-Control", "no-store");
    res.setHeader("X-Client-Source", source);
    if (graphError) {
      res.setHeader("X-Graph-Error", graphError.slice(0, 180));
    }

    res.status(200).json({
      clients: filtered.slice(0, limit),
      total: filtered.length,
    });
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error && error.message ? error.message : String(error),
    });
  }
};
