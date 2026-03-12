const fs = require("fs/promises");
const path = require("path");
const { requireApiAuth } = require("../_lib/require-api-auth");
const { buildConsultantDocx } = require("../_lib/consultant-docx");

function cleanText(value, max = 5000) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, max);
}

function cleanHtml(value, max = 120000) {
  return String(value || "").trim().slice(0, max);
}

function deriveNameFromEmail(email) {
  const local = String(email || "")
    .trim()
    .split("@")[0]
    .replace(/[._-]+/g, " ")
    .trim();
  if (!local) {
    return "";
  }
  return local
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.slice(0, 1).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "consultant"] }))) {
    return;
  }

  try {
    const consultantNameInput = cleanText(req.body?.consultantName, 200);
    const fallbackConsultantName =
      cleanText(req.authUser?.claims?.name || "", 200) ||
      deriveNameFromEmail(req.authUser?.email || "") ||
      "Consultant";
    const consultantName = consultantNameInput || fallbackConsultantName;
    const clientName = cleanText(req.body?.clientName, 200);
    const clientAddress = cleanText(req.body?.clientAddress, 600);
    const reportHtml = cleanHtml(req.body?.reportHtml, 120000);

    if (!consultantName || !clientName || !clientAddress || !reportHtml) {
      res.status(400).json({
        error: "consultantName, clientName, clientAddress, and reportHtml are required.",
      });
      return;
    }

    const templatePath = path.join(process.cwd(), "assets", "consultant-report-template.docx");
    const templateBuffer = await fs.readFile(templatePath);

    const now = new Date();
    const createdDate = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;

    const output = await buildConsultantDocx({
      templateBuffer,
      consultantName,
      clientName,
      clientAddress,
      reportHtml,
      createdDate,
    });

    res.status(200);
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", `attachment; filename="consultant-report-${createdDate}.docx"`);
    res.setHeader("Cache-Control", "no-store");
    res.send(output);
  } catch (error) {
    res.status(500).json({
      error: "Could not generate consultant report.",
      detail: error?.message || String(error),
    });
  }
};
