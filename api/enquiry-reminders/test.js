const { requireApiAuth } = require("../_lib/require-api-auth");
const { runEnquiryReminderJob } = require("../_lib/enquiry-reminders");

const ALLOWED_ROLES = ["admin"];

module.exports = async (req, res) => {
  if (req.method !== "GET" && req.method !== "POST") {
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
    const dryRun = req.method === "GET" || req.body?.dryRun !== false;
    const payload = await runEnquiryReminderJob({ dryRun });
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json(payload);
  } catch (error) {
    res.status(Number(error?.status) || 500).json({
      error: {
        code: String(error?.code || "ENQUIRY_REMINDER_TEST_FAILED"),
        message: error?.message || "Could not test enquiry reminders.",
      },
    });
  }
};
