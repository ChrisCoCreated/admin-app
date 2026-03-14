const { requireApiAuth } = require("../_lib/require-api-auth");
const { createAgendaItemForUser, mapAgendaError, updateAgendaItemForUser } = require("../_lib/agendas/service");

module.exports = async (req, res) => {
  if (!(await requireApiAuth(req, res))) {
    return;
  }

  try {
    if (req.method === "POST") {
      const payload = await createAgendaItemForUser(req.authUser?.email || "", req.body || {});
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json(payload);
      return;
    }

    if (req.method === "PATCH") {
      const payload = await updateAgendaItemForUser(req.authUser?.email || "", req.body || {});
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
