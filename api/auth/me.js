const { requireApiAuth } = require("../_lib/require-api-auth");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  const claims = await requireApiAuth(req, res);
  if (!claims) {
    return;
  }

  res.setHeader("Cache-Control", "no-store");
  res.status(200).json({
    email: req.authUser?.email || "",
    role: req.authUser?.role || "",
  });
};
