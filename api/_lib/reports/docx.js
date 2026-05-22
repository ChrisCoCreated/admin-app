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

function normalizeImagePayload(images) {
  const collage = images?.collage && typeof images.collage === "object" ? images.collage : null;
  const dataBase64 = String(collage?.dataBase64 || "").replace(/^data:image\/[a-z0-9.+-]+;base64,/i, "");
  if (!dataBase64) {
    return null;
  }
  const mimeType = String(collage?.mimeType || "image/png").toLowerCase();
  if (!/^image\/(?:png|jpeg|jpg)$/.test(mimeType)) {
    return null;
  }
  const width = Math.max(1, Number(collage?.width) || 2000);
  const height = Math.max(1, Number(collage?.height) || 1200);
  return {
    buffer: Buffer.from(dataBase64, "base64"),
    mimeType: mimeType === "image/jpg" ? "image/jpeg" : mimeType,
    width,
    height,
  };
}

function nextRelationshipId(relsXml) {
  const ids = [...String(relsXml || "").matchAll(/\bId="rId(\d+)"/g)].map((match) => Number(match[1]));
  return `rId${Math.max(0, ...ids) + 1}`;
}

function addImageRelationship(zip, image, imageIndex = 1) {
  const extension = image.mimeType === "image/jpeg" ? "jpg" : "png";
  const mediaPath = `word/media/report-collage-${imageIndex}.${extension}`;
  zip.file(mediaPath, image.buffer);

  const relsPath = "word/_rels/document.xml.rels";
  const relsFile = zip.file(relsPath);
  const relsPromise = relsFile
    ? relsFile.async("string")
    : Promise.resolve('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>');

  return relsPromise.then((relsXml) => {
    const relationshipId = nextRelationshipId(relsXml);
    const relationshipXml = `<Relationship Id="${relationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/report-collage-${imageIndex}.${extension}"/>`;
    const nextRelsXml = relsXml.replace("</Relationships>", `${relationshipXml}</Relationships>`);
    zip.file(relsPath, nextRelsXml);
    ensureImageContentType(zip, extension, image.mimeType);
    return relationshipId;
  });
}

function ensureImageContentType(zip, extension, mimeType) {
  const contentTypesFile = zip.file("[Content_Types].xml");
  if (!contentTypesFile) {
    return;
  }
  zip.file("[Content_Types].xml", contentTypesFile.async("string").then((xml) => {
    if (xml.includes(`Extension="${extension}"`)) {
      return xml;
    }
    return xml.replace("</Types>", `<Default Extension="${extension}" ContentType="${mimeType}"/></Types>`);
  }));
}

function imageDrawingXml(relationshipId, image) {
  const maxWidthEmu = 5486400; // 6 inches.
  const maxHeightEmu = 3657600; // 4 inches.
  const sourceWidth = Math.max(1, Number(image.width) || 1);
  const sourceHeight = Math.max(1, Number(image.height) || 1);
  const scale = Math.min(maxWidthEmu / sourceWidth, maxHeightEmu / sourceHeight);
  const cx = Math.round(sourceWidth * scale);
  const cy = Math.round(sourceHeight * scale);
  return `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1001" name="Report photo collage"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1001" name="Report photo collage"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relationshipId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
}

function replaceImagesPlaceholderWithDrawing(documentXml, relationshipId, image) {
  const normalized = normalizeSplitPlaceholder(documentXml, "images");
  const drawingXml = imageDrawingXml(relationshipId, image);
  const paragraphs = normalized.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  const placeholderParagraph = paragraphs.find((paragraph) => placeholderPattern("images").test(paragraph));
  if (placeholderParagraph) {
    return normalized.replace(placeholderParagraph, drawingXml);
  }
  return replacePlaceholderTokens(normalized, "images", "");
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

async function buildReportDocx({ reportType: reportTypeKey, report, images = {} }) {
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
  const image = normalizeImagePayload(images);
  if (image) {
    const relationshipId = await addImageRelationship(zip, image);
    documentXml = replaceImagesPlaceholderWithDrawing(documentXml, relationshipId, image);
  }
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
