const { requireGraphAuth } = require("../_lib/require-graph-auth");
const {
  getOneTouchAreaOptions,
  getOneTouchRecruitmentSourceOptions,
  getOneTouchPositionOptions,
  getOneTouchStatusOptions,
} = require("../_lib/onetouch-client");
const { RECRUITMENT_ALLOWED_ROLES } = require("../_lib/recruitment-access");

const ALLOWED_ROLES = RECRUITMENT_ALLOWED_ROLES;

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
    const [areas, recruitmentSources, positions, statuses] = await Promise.all([
      getOneTouchAreaOptions(),
      getOneTouchRecruitmentSourceOptions(),
      getOneTouchPositionOptions(),
      getOneTouchStatusOptions(),
    ]);
    res.setHeader("Cache-Control", "private, max-age=300");
    res.status(200).json({ areas, recruitmentSources, positions, statuses });
  } catch (error) {
    res.status(502).json({
      error: {
        code: "ONETOUCH_OPTIONS_FAILED",
        message: error?.message || "Could not load OneTouch location/area options.",
      },
    });
  }
};
