const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { getUnifiedTasksCached, mapGraphError } = require("../_lib/tasks/tasks-service");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({
      error: {
        code: "METHOD_NOT_ALLOWED",
        message: "Method Not Allowed",
      },
    });
    return;
  }

  if (!(await requireGraphAuth(req, res))) {
    return;
  }

  try {
    const payload = await getUnifiedTasksCached({
      graphAccessToken: req.authUser?.graphAccessToken,
      claims: req.authUser?.claims,
    });

    res.setHeader("Cache-Control", "private, max-age=20");
    res.status(200).json(payload);
  } catch (error) {
    const mapped = mapGraphError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
