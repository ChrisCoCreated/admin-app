const { requireApiAuth } = require("../_lib/require-api-auth");
const { listTaskSetTemplates } = require("../_lib/tasks/task-set-source");

const ALLOWED_ROLES = ["admin"];

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

  if (!(await requireApiAuth(req, res, { allowedRoles: ALLOWED_ROLES }))) {
    return;
  }

  try {
    const payload = await listTaskSetTemplates({
      taskSet: req.query?.taskSet,
      area: req.query?.area,
    });
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json(payload);
  } catch (error) {
    res.status(Number(error?.status) || 500).json({
      error: {
        code: String(error?.code || "TASK_SET_LIST_FAILED"),
        message: error?.message || "Could not load task set templates.",
      },
    });
  }
};
