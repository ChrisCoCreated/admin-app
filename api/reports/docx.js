const { requireApiAuth } = require("../_lib/require-api-auth");
const { buildReportDocx } = require("../_lib/reports/docx");

function safeFilename(value) {
  return String(value || "report")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80) || "report";
}

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
    const report = body.report && typeof body.report === "object" ? body.report : null;
    if (!report) {
      res.status(400).json({
        error: {
          code: "REPORT_REQUIRED",
          message: "report is required.",
        },
      });
      return;
    }

    const output = await buildReportDocx({
      reportType: body.reportType || body.report_type || report.report_type,
      report,
    });

    const filename = `${safeFilename(report.report_title || "wellbeing-assurance-visit")}.docx`;
    res.status(200);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Cache-Control", "no-store");
    res.send(output);
  } catch (error) {
    res.status(Number(error?.status) || 500).json({
      error: {
        code: "REPORT_DOCX_FAILED",
        message: error?.message || "Could not generate report document.",
      },
    });
  }
};
