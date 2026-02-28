const { requireApiAuth } = require("../../_lib/require-api-auth");
const { buildReconciliationPreview } = require("../../_lib/clients-reconcile-service");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "care_manager"] }))) {
    return;
  }

  try {
    const preview = await buildReconciliationPreview();
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json(preview);
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error?.message || String(error),
    });
  }
};
