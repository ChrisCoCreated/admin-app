const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { mapGraphError, syncWhiteboardTasks } = require("../_lib/tasks/tasks-service");

module.exports = async (req, res) => {
  if (req.method !== "POST") {
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
    const payload = await syncWhiteboardTasks({
      graphAccessToken: req.authUser?.graphAccessToken,
      claims: req.authUser?.claims,
      body: req.body && typeof req.body === "object" ? req.body : {},
    });

    res.setHeader("Cache-Control", "no-store");
    res.status(Number(payload?.statusCode) || 200).json({
      meta: payload?.meta || {},
    });
  } catch (error) {
    const mapped = mapGraphError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
