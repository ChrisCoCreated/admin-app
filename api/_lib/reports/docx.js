const fs = require("fs/promises");
const JSZip = require("jszip");
const { normalizeText } = require("../deepseek-client");
const { getReportType } = require("./registry");

function escapeXml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function stripXmlTags(value) {
  return String(value || "").replace(/<[^>]+>/g, "");
}

function escapeRegExp(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function valueAtPath(source, path) {
  return String(path || "")
    .split(".")
    .filter(Boolean)
    .reduce((current, part) => (current && typeof current === "object" ? current[part] : undefined), source);
}

function asText(value) {
  if (Array.isArray(value)) {
    return value.map(asText).filter(Boolean).join("\n");
  }
  if (value && typeof value === "object") {
    return Object.values(value).map(asText).filter(Boolean).join(" ");
  }
  return normalizeText(value);
}

function goalField(goal, keys) {
  for (const key of keys) {
    const value = asText(goal?.[key]);
    if (value) {
      return value;
    }
  }
  return "";
}

function buildGoalRows(report) {
  const goals = Array.isArray(report?.suggested_smart_goals) ? report.suggested_smart_goals.slice(0, 3) : [];
  while (goals.length < 3) {
    goals.push({});
  }
  return goals.map((goal) => ({
    goal: goalField(goal, ["goal", "specific", "title"]),
    measure: goalField(goal, ["measure", "measurable"]),
    owner: goalField(goal, ["owner", "responsible_person"]),
    timescale: goalField(goal, ["timescale", "time_bound", "timeframe"]),
    status_notes:
      goalField(goal, ["status_notes", "notes", "source_note_reference"]) ||
      (goalField(goal, ["goal", "specific", "title"]) ? "Suggested goal - review or amend as needed." : ""),
  }));
}

function placeholderPattern(placeholder, flags = "g") {
  return new RegExp(`\\{\\{\\s*${escapeRegExp(placeholder)}\\s*\\}\\}`, flags);
}

function replacePlaceholderTokens(xml, placeholder, value) {
  return xml.replace(placeholderPattern(placeholder), escapeXml(value));
}

function replaceFirstPlaceholderToken(xml, placeholder, value) {
  return xml.replace(placeholderPattern(placeholder, ""), escapeXml(value));
}

function replaceRepeatedGoalTokens(documentXml, report) {
  let output = documentXml;
  const rows = buildGoalRows(report);
  for (const row of rows) {
    output = replaceFirstPlaceholderToken(output, "goal", row.goal);
    output = replaceFirstPlaceholderToken(output, "measure", row.measure);
    output = replaceFirstPlaceholderToken(output, "owner", row.owner);
    output = replaceFirstPlaceholderToken(output, "timescale", row.timescale);
    output = replaceFirstPlaceholderToken(output, "status_notes", row.status_notes);
  }
  return output;
}

function normalizeSplitPlaceholder(documentXml, placeholder) {
  const escaped = escapeRegExp(placeholder);
  const pattern = new RegExp(
    `<w:t>\\{\\{<\\/w:t><\\/w:r>(?:(?!<\\/w:p>)[\\s\\S])*?<w:t>\\s*${escaped}\\s*<\\/w:t><\\/w:r>(?:(?!<\\/w:p>)[\\s\\S])*?<w:t>\\}\\}<\\/w:t>`,
    "g"
  );
  return documentXml.replace(pattern, `<w:t>{{${placeholder}}}</w:t>`);
}

function normalizeSplitPlaceholders(documentXml, placeholders) {
  return placeholders.reduce((output, placeholder) => normalizeSplitPlaceholder(output, placeholder), documentXml);
}

function replaceTextPlaceholders(documentXml, reportType, report) {
  const placeholders = [
    ...Object.keys(reportType.template_mapping),
    "goal",
    "measure",
    "owner",
    "timescale",
    "status_notes",
    "images",
    "assessor_name",
    "assessor_role",
    "organisation",
    "email",
    "phone",
  ];
  let output = normalizeSplitPlaceholders(documentXml, placeholders);
  for (const [placeholder, sourcePath] of Object.entries(reportType.template_mapping)) {
    if (placeholder === "smart_goals") {
      continue;
    }
    const value = asText(valueAtPath(report, sourcePath));
    output = replacePlaceholderTokens(output, placeholder, value);
  }

  for (const placeholder of ["images", "assessor_name", "assessor_role", "organisation", "email", "phone"]) {
    output = replacePlaceholderTokens(output, placeholder, "");
  }
  return replaceRepeatedGoalTokens(output, report);
}

function removeParagraphRangeByText(documentXml, startText, endText) {
  const paragraphs = documentXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  if (!paragraphs.length) {
    return documentXml;
  }

  let removing = false;
  const remove = new Set();
  for (const paragraph of paragraphs) {
    const text = stripXmlTags(paragraph);
    if (text.includes(startText)) {
      removing = true;
    }
    if (removing) {
      remove.add(paragraph);
    }
    if (endText && text.includes(endText)) {
      removing = false;
      remove.delete(paragraph);
    }
  }

  let output = documentXml;
  for (const paragraph of remove) {
    output = output.replace(paragraph, "");
  }
  return output;
}

function removeImplementationNotes(documentXml) {
  const marker = "Implementation Notes for Reporting Flow";
  const paragraphs = documentXml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  let found = false;
  const remove = new Set();
  for (const paragraph of paragraphs) {
    const text = stripXmlTags(paragraph);
    if (text.includes(marker)) {
      found = true;
    }
    if (found) {
      remove.add(paragraph);
    }
  }
  let output = documentXml;
  for (const paragraph of remove) {
    output = output.replace(paragraph, "");
  }
  return output;
}

function shouldOmitWellbeingHighlights(report) {
  const omitted = Array.isArray(report?.omitted_sections) ? report.omitted_sections : [];
  return omitted.some((item) => /wellbeing highlights/i.test(String(item || "")));
}

async function buildReportDocx({ reportType: reportTypeKey, report }) {
  const reportType = getReportType(reportTypeKey || report?.report_type);
  const templateBuffer = await fs.readFile(reportType.template_document);
  const zip = await JSZip.loadAsync(templateBuffer);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) {
    throw new Error("Template missing word/document.xml.");
  }

  let documentXml = await documentFile.async("string");
  if (shouldOmitWellbeingHighlights(report)) {
    documentXml = removeParagraphRangeByText(documentXml, "Wellbeing Highlights", "Goals");
  }
  documentXml = removeImplementationNotes(documentXml);
  documentXml = replaceTextPlaceholders(documentXml, reportType, report || {});

  zip.file("word/document.xml", documentXml);
  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}

module.exports = {
  buildReportDocx,
  shouldOmitWellbeingHighlights,
};
