const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { getOneTouchAreaOptions } = require("../_lib/onetouch-client");

const ALLOWED_ROLES = [
  "admin",
  "care_manager",
  "operations",
  "hr_only",
  "hr_clients",
  "time_hr",
  "time_hr_clients",
];

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

  if (!(await requireGraphAuth(req, res, { allowedRoles: ALLOWED_ROLES }))) {
    return;
  }

  try {
    const areas = await getOneTouchAreaOptions();
    res.setHeader("Cache-Control", "private, max-age=300");
    res.status(200).json({ areas });
  } catch (error) {
    res.status(502).json({
      error: {
        code: "ONETOUCH_OPTIONS_FAILED",
        message: error?.message || "Could not load OneTouch location/area options.",
      },
    });
  }
};
