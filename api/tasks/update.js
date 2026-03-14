const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { mapGraphError, updateWhiteboardTask } = require("../_lib/tasks/tasks-service");

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
    const payload = await updateWhiteboardTask({
      graphAccessToken: req.authUser?.graphAccessToken,
      body: req.body && typeof req.body === "object" ? req.body : {},
    });

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json(payload);
  } catch (error) {
    const mapped = mapGraphError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
