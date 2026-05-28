const { runEnquiryReminderJob } = require("../_lib/enquiry-reminders");

function getAuthorizationHeader(req) {
  return String(req?.headers?.authorization || req?.headers?.Authorization || "").trim();
}

function isCronAuthorized(req) {
  const secret = String(process.env.CRON_SECRET || "").trim();
  return Boolean(secret) && getAuthorizationHeader(req) === `Bearer ${secret}`;
}

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

  if (!isCronAuthorized(req)) {
    res.status(401).json({
      error: {
        code: "UNAUTHORIZED",
        message: "Unauthorized",
      },
    });
    return;
  }

  try {
    const dryRun = String(req.query?.dryRun || "").trim() === "1";
    const payload = await runEnquiryReminderJob({ dryRun });
    res.setHeader("Cache-Control", "no-store");
    res.status(200).json(payload);
  } catch (error) {
    res.status(Number(error?.status) || 500).json({
      error: {
        code: String(error?.code || "ENQUIRY_REMINDER_FAILED"),
        message: error?.message || "Could not run enquiry reminders.",
      },
    });
  }
};

module.exports.isCronAuthorized = isCronAuthorized;
