const { requireApiAuth } = require("../_lib/require-api-auth");
const { getAgendaDetailForUser, mapAgendaError } = require("../_lib/agendas/service");

module.exports = async (req, res) => {
  if (!(await requireApiAuth(req, res))) {
    return;
  }

  try {
    if (req.method !== "GET") {
      res.status(405).json({
        error: { code: "METHOD_NOT_ALLOWED", message: "Method Not Allowed" },
      });
      return;
    }

    const payload = await getAgendaDetailForUser(req.authUser?.email || "", req.query?.agendaId || "");
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json(payload);
  } catch (error) {
    const mapped = mapAgendaError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
