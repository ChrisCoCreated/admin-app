const { requireApiAuth } = require("../_lib/require-api-auth");
const { buildReportDocx } = require("../_lib/reports/docx");

function safeFilenamePart(value, fallback) {
  return String(value || fallback || "")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "")
    .replace(/\s+/g, " ")
    .slice(0, 80) || fallback;
}

function parseReportDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value;
  }
  const raw = String(value || "").trim();
  if (!raw) {
    return null;
  }

  const isoMatch = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    return new Date(Number(year), Number(month) - 1, Number(day));
  }

  const ukMatch = raw.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (ukMatch) {
    const [, day, month, year] = ukMatch;
    const fullYear = year.length === 2 ? Number(`20${year}`) : Number(year);
    return new Date(fullYear, Number(month) - 1, Number(day));
  }

  const parsed = new Date(raw);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function formatReportDate(value, fallbackDate = new Date()) {
  const date = parseReportDate(value) || fallbackDate;
  return new Intl.DateTimeFormat("en-GB", {
    day: "numeric",
    month: "long",
    year: "numeric",
  }).format(date);
}

function titleClientName(title) {
  const match = String(title || "").match(/\bfor\s+(.+)$/i);
  return match ? match[1].trim() : "";
}

function cleanReportTitle(title, clientName = "") {
  const raw = String(title || "Report").trim();
  if (!clientName) {
    return raw;
  }
  return raw.replace(new RegExp(`\\s+for\\s+${clientName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*$`, "i"), "").trim() || raw;
}

function buildReportFilename(report, fallbackDate = new Date()) {
  const clientDetails = report?.client_details && typeof report.client_details === "object" ? report.client_details : {};
  const rawTitle = report?.report_title || "Report";
  const clientName =
    clientDetails.client_name ||
    clientDetails.name ||
    clientDetails.full_name ||
    clientDetails.preferred_name ||
    titleClientName(rawTitle) ||
    "Client";
  const reportDate =
    clientDetails.report_date ||
    clientDetails.visit_date ||
    clientDetails.assessment_date ||
    report?.report_date ||
    report?.date;
  const title = cleanReportTitle(rawTitle, clientName);
  const filename = [
    safeFilenamePart(clientName, "Client"),
    safeFilenamePart(title, "Report"),
    safeFilenamePart(formatReportDate(reportDate, fallbackDate), "Date"),
  ].join(" - ");
  return `${filename}.docx`;
}

function asciiHeaderFilename(filename) {
  const fallback = String(filename || "report.docx")
    .normalize("NFKD")
    .replace(/[^\x20-\x7E]+/g, "")
    .replace(/["\\]/g, "")
    .replace(/[;,]/g, "")
    .replace(/\s+/g, " ")
    .trim();
  return fallback || "report.docx";
}

function contentDispositionForFilename(filename) {
  const clean = String(filename || "report.docx").replace(/[\r\n]+/g, " ").trim() || "report.docx";
  return `attachment; filename="${asciiHeaderFilename(clean)}"; filename*=UTF-8''${encodeURIComponent(clean)}`;
}

async function reportDocxHandler(req, res) {
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

    const filename = buildReportFilename(report);
    res.status(200);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", contentDispositionForFilename(filename));
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
}

module.exports = reportDocxHandler;
module.exports.buildReportFilename = buildReportFilename;
module.exports.contentDispositionForFilename = contentDispositionForFilename;
module.exports.formatReportDate = formatReportDate;
