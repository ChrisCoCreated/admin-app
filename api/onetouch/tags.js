const { listOneTouchTags } = require("../_lib/onetouch-client");
const { requireApiAuth } = require("../_lib/require-api-auth");

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (
    !(await requireApiAuth(req, res, {
      allowedRoles: [
        "admin",
        "care_manager",
        "operations",
        "clients_only",
        "hr_only",
        "hr_clients",
        "time_only",
        "time_clients",
        "time_hr",
        "time_hr_clients",
        "consultant",
      ],
    }))
  ) {
    return;
  }

  try {
    const tags = await listOneTouchTags();
    res.setHeader("Cache-Control", "private, max-age=60");
    res.status(200).json({
      tags,
      total: tags.length,
    });
  } catch (error) {
    res.status(500).json({
      error: "Server error",
      detail: error?.message || String(error),
    });
  }
};
