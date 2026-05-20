const { requireApiAuth } = require("../_lib/require-api-auth");
const { generateStructuredReport } = require("../_lib/reports/service");

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

  if (!(await requireApiAuth(req, res))) {
    return;
  }

  const body = req.body && typeof req.body === "object" ? req.body : {};
  try {
    const report = await generateStructuredReport({
      reportType: body.reportType || body.report_type,
      notes: body.notes,
      provider: body.provider,
      model: body.model,
      thinking: body.thinking,
      reasoningEffort: body.reasoningEffort || body.reasoning_effort,
      previousReport: body.previousReport || body.previous_report,
      revisionRequest: body.revisionRequest || body.revision_request,
    });

    res.status(200).json({
      success: true,
      report,
    });
  } catch (error) {
    res.status(Number(error?.status) || 500).json({
      error: {
        code: "REPORT_GENERATION_FAILED",
        message: error?.message || "Could not generate report.",
      },
    });
  }
};
