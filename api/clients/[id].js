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
    const targetId = String(req.query.id || "").trim();
    const client = clients.find((item) => item.id === targetId);

    res.setHeader("Cache-Control", "no-store");
    res.setHeader("X-Client-Source", source);
    if (graphError) {
      res.setHeader("X-Graph-Error", graphError.slice(0, 180));
    }

    if (!client) {
      res.status(404).json({ error: "Client not found" });
      return;
    }

    res.status(200).json({ client });
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error && error.message ? error.message : String(error),
    });
  }
};
