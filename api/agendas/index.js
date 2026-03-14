const { requireApiAuth } = require("../_lib/require-api-auth");
const {
  createAgendaForUser,
  listAgendasForUser,
  listAgendaSummariesForUser,
  mapAgendaError,
  updateAgendaForUser,
} = require("../_lib/agendas/service");

module.exports = async (req, res) => {
  if (!(await requireApiAuth(req, res))) {
    return;
  }

  try {
    if (req.method === "GET") {
      const summaryOnly = String(req.query?.summaryOnly || "").trim().toLowerCase() === "true";
      const payload = summaryOnly
        ? await listAgendaSummariesForUser(req.authUser?.email || "")
        : await listAgendasForUser(req.authUser?.email || "");
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json(payload);
      return;
    }

    if (req.method === "POST") {
      const payload = await createAgendaForUser(req.authUser?.email || "", req.body || {});
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json(payload);
      return;
    }

    if (req.method === "PATCH") {
      const payload = await updateAgendaForUser(req.authUser?.email || "", req.body || {});
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json(payload);
      return;
    }

    res.status(405).json({
      error: { code: "METHOD_NOT_ALLOWED", message: "Method Not Allowed" },
    });
  } catch (error) {
    const mapped = mapAgendaError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
