const { requireApiAuth } = require("./_lib/require-api-auth");
const { listAssignments, saveAssignment } = require("./_lib/role-assignments");

module.exports = async (req, res) => {
  res.setHeader("Cache-Control", "no-store");
  if (!["GET", "PUT"].includes(req.method)) {
    res.setHeader("Allow", "GET, PUT");
    return res.status(405).json({ error: "Method Not Allowed" });
  }
  if (!(await requireApiAuth(req, res))) return;
  // Exact capability check: page-equivalent roles must never grant management access.
  if (req.authUser?.isSuperadmin !== true) return res.status(403).json({ error: "Superadmin access required." });
  try {
    if (req.method === "PUT") {
      const assignment = await saveAssignment(req.body, req.authUser.email);
      return res.status(200).json({ assignment });
    }
    return res.status(200).json(await listAssignments());
  } catch (error) {
    console.error("Role management failed:", error.message);
    return res.status(error.status || 503).json({ error: error.status ? error.message : "Shared role storage is temporarily unavailable." });
  }
};
