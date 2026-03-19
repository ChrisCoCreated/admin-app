const { requireApiAuth } = require("../_lib/require-api-auth");
const { createPlannerBatch } = require("../_lib/tasks/planner-batch-service");

const ALLOWED_ROLES = ["admin"];

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

  if (!(await requireApiAuth(req, res, { allowedRoles: ALLOWED_ROLES }))) {
    return;
  }

  try {
    const payload = await createPlannerBatch(req.body && typeof req.body === "object" ? req.body : {});
    res.setHeader("Cache-Control", "no-store");
    res.status(payload?.dryRun ? 200 : 201).json(payload);
  } catch (error) {
    res.status(Number(error?.status) || 500).json({
      error: {
        code: String(error?.code || "TASK_BATCH_CREATE_FAILED"),
        message: error?.message || "Could not create Planner task batch.",
      },
    });
  }
};
