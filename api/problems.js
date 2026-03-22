const { requireApiAuth } = require("./_lib/require-api-auth");
const { createProblemForUser, listProblemsForUser, mapProblemError, updateProblemForUser } = require("./_lib/problems/service");

module.exports = async (req, res) => {
  if (!(await requireApiAuth(req, res))) {
    return;
  }

  try {
    if (req.method === "GET") {
      const payload = await listProblemsForUser(req.authUser?.email || "");
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json(payload);
      return;
    }

    if (req.method === "POST") {
      const payload = await createProblemForUser(req.authUser?.email || "", req.body || {});
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json(payload);
      return;
    }

    if (req.method === "PATCH") {
      const payload = await updateProblemForUser(req.authUser?.email || "", req.body || {});
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json(payload);
      return;
    }

    res.status(405).json({
      error: { code: "METHOD_NOT_ALLOWED", message: "Method Not Allowed" },
    });
  } catch (error) {
    const mapped = mapProblemError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
