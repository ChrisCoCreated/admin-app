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

function replaceFirstToken(xml, token, value) {
  return xml.replace(token, escapeXml(value));
}

function replaceRepeatedGoalTokens(documentXml, report) {
  let output = documentXml;
  const rows = buildGoalRows(report);
  for (const row of rows) {
    output = replaceFirstToken(output, "{{goal}}", row.goal);
    output = replaceFirstToken(output, "{{measure}}", row.measure);
    output = replaceFirstToken(output, "{{owner}}", row.owner);
    output = replaceFirstToken(output, "{{timescale}}", row.timescale);
    output = replaceFirstToken(output, "{{status_notes}}", row.status_notes);
  }
  return output;
}

function normalizeSplitPlaceholders(documentXml) {
  return documentXml.replace(
    /<w:t>\{\{<\/w:t><\/w:r>(?:(?!<\/w:p>)[\s\S])*?<w:t>assessment_type<\/w:t><\/w:r>(?:(?!<\/w:p>)[\s\S])*?<w:t>\}\}<\/w:t>/g,
    "<w:t>{{assessment_type}}</w:t>"
  );
}

function replaceTextPlaceholders(documentXml, reportType, report) {
  let output = normalizeSplitPlaceholders(documentXml);
  for (const [placeholder, sourcePath] of Object.entries(reportType.template_mapping)) {
    if (placeholder === "smart_goals") {
      continue;
    }
    const token = `{{${placeholder}}}`;
    const value = asText(valueAtPath(report, sourcePath));
    output = output.split(token).join(escapeXml(value));
  }

  for (const placeholder of ["assessor_name", "assessor_role", "organisation", "email", "phone"]) {
    output = output.split(`{{${placeholder}}}`).join("");
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
    documentXml = removeParagraphRangeByText(documentXml, "5. Wellbeing Highlights", "6. SMART Goals");
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
