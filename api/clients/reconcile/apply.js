const { requireApiAuth } = require("../../_lib/require-api-auth");
const { applyReconciliationAction } = require("../../_lib/clients-reconcile-service");

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "care_manager"] }))) {
    return;
  }

  try {
    const result = await applyReconciliationAction(req.body || {});
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({ result });
  } catch (error) {
    res.status(error?.status || 500).json({
      error: error?.status ? "Request failed" : "Server error",
      detail: error?.message || String(error),
    });
  }
};
