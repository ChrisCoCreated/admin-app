const JSZip = require("jszip");

const PLACEHOLDER_BODY = "{{REPORT_BODY}}";

function escapeXml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function decodeEntities(value) {
  return String(value || "")
    .replaceAll("&nbsp;", " ")
    .replaceAll("&amp;", "&")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&#39;", "'");
}

function tokenizeHtml(html) {
  const tokens = [];
  const regex = /<[^>]+>|[^<]+/g;
  let match;
  while ((match = regex.exec(String(html || ""))) !== null) {
    tokens.push(match[0]);
  }
  return tokens;
}

function isAllowedTag(tag) {
  return new Set(["p", "br", "strong", "b", "em", "i", "ul", "ol", "li", "h1", "h2", "h3"]).has(tag);
}

function parseTag(token) {
  const clean = token.trim();
  const closeMatch = /^<\s*\/\s*([a-z0-9]+)[^>]*>$/i.exec(clean);
  if (closeMatch) {
    const tag = closeMatch[1].toLowerCase();
    return isAllowedTag(tag) ? { type: "close", tag } : null;
  }

  const openMatch = /^<\s*([a-z0-9]+)(\s[^>]*)?>$/i.exec(clean);
  if (!openMatch) {
    return null;
  }

  const tag = openMatch[1].toLowerCase();
  if (!isAllowedTag(tag)) {
    return null;
  }

  const selfClosing = /\/>$/.test(clean) || tag === "br";
  return { type: "open", tag, selfClosing };
}

function createParagraph({ style = "", listType = "", level = 0 } = {}) {
  return {
    style,
    listType,
    level,
    runs: [],
  };
}

function createRun(text, bold, italic) {
  return {
    type: "text",
    text,
    bold,
    italic,
  };
}

function createBreakRun() {
  return { type: "break" };
}

function paragraphToXml(paragraph) {
  const pPrParts = [];
  if (paragraph.style) {
    pPrParts.push(`<w:pStyle w:val="${escapeXml(paragraph.style)}"/>`);
  }

  if (paragraph.listType) {
    const numId = paragraph.listType === "ol" ? 2 : 1;
    const ilvl = Number.isFinite(paragraph.level) ? Math.max(0, Math.min(paragraph.level, 8)) : 0;
    pPrParts.push(`<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`);
  }

  const pPrXml = pPrParts.length ? `<w:pPr>${pPrParts.join("")}</w:pPr>` : "";
  const runsXml = paragraph.runs
    .map((run) => {
      if (run.type === "break") {
        return "<w:r><w:br/></w:r>";
      }
      const rPr = [];
      if (run.bold) {
        rPr.push("<w:b/>");
      }
      if (run.italic) {
        rPr.push("<w:i/>");
      }
      const rPrXml = rPr.length ? `<w:rPr>${rPr.join("")}</w:rPr>` : "";
      const text = escapeXml(run.text);
      return `<w:r>${rPrXml}<w:t xml:space="preserve">${text}</w:t></w:r>`;
    })
    .join("");

  const fallbackRun = runsXml || "<w:r><w:t></w:t></w:r>";
  return `<w:p>${pPrXml}${fallbackRun}</w:p>`;
}

function htmlToWordParagraphs(html) {
  const tokens = tokenizeHtml(html);
  const paragraphs = [];

  let current = null;
  let boldDepth = 0;
  let italicDepth = 0;
  const listStack = [];

  function ensureParagraph(style = "") {
    if (!current) {
      const listType = listStack.length ? listStack[listStack.length - 1] : "";
      current = createParagraph({
        style,
        listType,
        level: Math.max(0, listStack.length - 1),
      });
    }
    if (style && !current.style) {
      current.style = style;
    }
  }

  function commitParagraph() {
    if (current) {
      paragraphs.push(current);
      current = null;
    }
  }

  for (const token of tokens) {
    if (token.startsWith("<")) {
      const tag = parseTag(token);
      if (!tag) {
        continue;
      }

      if (tag.type === "open") {
        if (tag.tag === "strong" || tag.tag === "b") {
          boldDepth += 1;
          continue;
        }
        if (tag.tag === "em" || tag.tag === "i") {
          italicDepth += 1;
          continue;
        }

        if (tag.tag === "ul" || tag.tag === "ol") {
          listStack.push(tag.tag);
          continue;
        }

        if (tag.tag === "p") {
          commitParagraph();
          ensureParagraph("");
          continue;
        }

        if (tag.tag === "h1" || tag.tag === "h2" || tag.tag === "h3") {
          commitParagraph();
          const style = tag.tag === "h1" ? "Heading1" : tag.tag === "h2" ? "Heading2" : "Heading3";
          ensureParagraph(style);
          continue;
        }

        if (tag.tag === "li") {
          commitParagraph();
          const listType = listStack.length ? listStack[listStack.length - 1] : "ul";
          current = createParagraph({
            listType,
            level: Math.max(0, listStack.length - 1),
          });
          continue;
        }

        if (tag.tag === "br") {
          ensureParagraph("");
          current.runs.push(createBreakRun());
          continue;
        }

        if (tag.selfClosing) {
          continue;
        }
      }

      if (tag.type === "close") {
        if (tag.tag === "strong" || tag.tag === "b") {
          boldDepth = Math.max(0, boldDepth - 1);
          continue;
        }
        if (tag.tag === "em" || tag.tag === "i") {
          italicDepth = Math.max(0, italicDepth - 1);
          continue;
        }

        if (tag.tag === "ul" || tag.tag === "ol") {
          listStack.pop();
          commitParagraph();
          continue;
        }

        if (tag.tag === "p" || tag.tag === "li" || tag.tag === "h1" || tag.tag === "h2" || tag.tag === "h3") {
          commitParagraph();
          continue;
        }
      }

      continue;
    }

    const text = decodeEntities(token).replace(/\s+/g, " ");
    if (!text.trim()) {
      if (current && current.runs.length) {
        current.runs.push(createRun(" ", boldDepth > 0, italicDepth > 0));
      }
      continue;
    }

    ensureParagraph("");
    current.runs.push(createRun(text, boldDepth > 0, italicDepth > 0));
  }

  commitParagraph();

  if (!paragraphs.length) {
    paragraphs.push(createParagraph({}));
  }

  return paragraphs.map(paragraphToXml).join("");
}

function replaceBodyPlaceholder(documentXml, bodyXml) {
  const bodyParagraphPattern = /<w:p\b[\s\S]*?\{\{REPORT_BODY\}\}[\s\S]*?<\/w:p>/;
  if (bodyParagraphPattern.test(documentXml)) {
    return documentXml.replace(bodyParagraphPattern, bodyXml);
  }
  return documentXml.replace(PLACEHOLDER_BODY, bodyXml);
}

function replaceTextPlaceholders(documentXml, replacements) {
  let output = documentXml;
  for (const [placeholder, value] of Object.entries(replacements)) {
    output = output.split(`{{${placeholder}}}`).join(escapeXml(value));
  }
  return output;
}

async function buildConsultantDocx({ templateBuffer, consultantName, clientName, clientAddress, reportHtml, createdDate }) {
  const zip = await JSZip.loadAsync(templateBuffer);
  const documentFile = zip.file("word/document.xml");
  if (!documentFile) {
    throw new Error("Template missing word/document.xml.");
  }

  let documentXml = await documentFile.async("string");
  documentXml = replaceTextPlaceholders(documentXml, {
    CONSULTANT_NAME: consultantName,
    CLIENT_NAME: clientName,
    CLIENT_ADDRESS: clientAddress,
    CREATED_DATE: createdDate,
  });

  const bodyXml = htmlToWordParagraphs(reportHtml);
  documentXml = replaceBodyPlaceholder(documentXml, bodyXml);
  zip.file("word/document.xml", documentXml);

  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}

module.exports = {
  buildConsultantDocx,
};
