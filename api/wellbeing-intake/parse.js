const { requireApiAuth } = require("../_lib/require-api-auth");

const ALLOWED_TYPES = new Set(["Text", "Note", "Choice", "MultiChoice", "Boolean", "DateTime", "Number", "Currency", "URL"]);
const DEFAULT_MODEL = "deepseek-chat";

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeField(field) {
  const internalName = normalizeText(field?.internalName);
  const title = normalizeText(field?.title);
  const type = normalizeText(field?.type);
  const choices = Array.isArray(field?.choices) ? field.choices.map(normalizeText).filter(Boolean) : [];
  return {
    internalName,
    title,
    type,
    required: field?.required === true,
    description: normalizeText(field?.description),
    choices,
  };
}

function sanitizeFields(fields) {
  return (Array.isArray(fields) ? fields : [])
    .map(normalizeField)
    .filter((field) => field.internalName && field.title && ALLOWED_TYPES.has(field.type));
}

function buildFieldGuide(fields) {
  return fields
    .map((field) => {
      const parts = [`${field.title} [${field.internalName}]`, `type=${field.type}`];
      if (field.required) {
        parts.push("required");
      }
      if (field.description) {
        parts.push(`description=${field.description}`);
      }
      if (field.choices.length) {
        parts.push(`choices=${field.choices.join(" | ")}`);
      }
      return `- ${parts.join("; ")}`;
    })
    .join("\n");
}

function buildSchema(fields) {
  const properties = {};
  const required = [];

  for (const field of fields) {
    properties[field.internalName] = {
      type: ["string", "null"],
      description: `${field.title} (${field.type})`,
    };
    required.push(field.internalName);
  }

  return {
    type: "json_schema",
    name: "wellbeing_intake_entry",
    strict: true,
    schema: {
      type: "object",
      additionalProperties: false,
      properties,
      required,
    },
  };
}

function buildPrompt({ sourceText, fields }) {
  return [
    "Return valid json only.",
    "Turn the pasted text into a draft SharePoint list entry.",
    "This is for a suppliers and experiences SharePoint list.",
    "Only use facts present in the text.",
    "If a field is not clearly present, return null.",
    "Do not copy email headers or quoted thread text into Notes unless truly needed.",
    "Prefer concise cleaned values over raw pasted blocks.",
    "For Title, prefer the supplier or contact name.",
    "For Contact Details, include the clean email address and any phone number, not the full email header.",
    "For Notes, summarise the service offered, coverage area, home-visit availability, and pricing.",
    "Only choose a Supplier Type when one of the available choices is clearly supported by the text; otherwise return null.",
    "Only choose Town when one of the exact available choices is clearly mentioned.",
    "If places mentioned are in Kent and County has Kent as an option, choose Kent.",
    "For choice fields, use the closest exact choice string when possible.",
    "For booleans, return 'true' or 'false' as strings.",
    "For dates, prefer YYYY-MM-DD when the date is clear.",
    "For URLs, return the full URL.",
    "",
    "Fields:",
    buildFieldGuide(fields),
    "",
    "Pasted text:",
    sourceText,
  ].join("\n");
}

function extractResponseText(payload) {
  if (typeof payload?.output_text === "string" && payload.output_text.trim()) {
    return payload.output_text.trim();
  }

  const outputs = Array.isArray(payload?.output) ? payload.output : [];
  for (const output of outputs) {
    const content = Array.isArray(output?.content) ? output.content : [];
    for (const item of content) {
      if (typeof item?.text === "string" && item.text.trim()) {
        return item.text.trim();
      }
    }
  }

  return "";
}

function parseJsonContent(content) {
  if (!content) {
    throw new Error("DeepSeek returned an empty extraction response.");
  }

  const parsed = JSON.parse(content);
  return parsed && typeof parsed === "object" ? parsed : {};
}

function fieldByTitle(fields, candidates) {
  for (const candidate of candidates) {
    const matched = fields.find((field) => normalizeText(field.title).toLowerCase() === normalizeText(candidate).toLowerCase());
    if (matched) {
      return matched;
    }
  }
  return null;
}

function extractEmail(sourceText) {
  const matched = String(sourceText || "").match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return matched ? normalizeText(matched[0]) : "";
}

function extractUrl(sourceText) {
  const matched = String(sourceText || "").match(/https?:\/\/\S+/i);
  return matched ? normalizeText(matched[0].replace(/[),.;]+$/, "")) : "";
}

function extractName(sourceText) {
  const contactMatch = String(sourceText || "").match(/Contact:\s*([^\n\r]+)/i);
  if (contactMatch?.[1]) {
    return normalizeText(contactMatch[1]);
  }

  const fromMatch = String(sourceText || "").match(/From:\s*([^<\n\r]+?)\s*</i);
  if (fromMatch?.[1]) {
    return normalizeText(fromMatch[1]);
  }

  const signoffMatch = String(sourceText || "").match(/thank you\s+([A-Z][A-Za-z' -]{1,80})/i);
  if (signoffMatch?.[1]) {
    return normalizeText(signoffMatch[1]);
  }

  return "";
}

function extractPhone(sourceText) {
  const matched = String(sourceText || "").match(/(?:\+44\s?7\d{3}|\(?07\d{3}\)?)\s?\d{3}\s?\d{3,4}/i);
  return matched ? normalizeText(matched[0].replace(/\s+/g, " ")) : "";
}

function extractBodyText(sourceText) {
  const lines = String(sourceText || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const kept = [];
  for (const line of lines) {
    if (/^(from|to|subject|date):/i.test(line)) {
      continue;
    }
    if (/^sent from my iphone$/i.test(line)) {
      continue;
    }
    if (/^on .* wrote:$/i.test(line)) {
      break;
    }
    if (/^www\./i.test(line) || /^thrive homecare$/i.test(line)) {
      continue;
    }
    if (/^<image/i.test(line)) {
      continue;
    }
    kept.push(line);
  }

  return kept.join(" ").replace(/\s+/g, " ").trim();
}

function splitLines(sourceText) {
  return String(sourceText || "")
    .split(/\r?\n/)
    .map((line) => line.trim());
}

function isLikelyHeading(line) {
  const value = normalizeText(line);
  if (!value || value.length > 80) {
    return false;
  }
  return /^[A-Z][A-Za-z0-9 /&()+-]+$/.test(value);
}

function extractStructuredSections(sourceText) {
  const lines = splitLines(sourceText);
  const sections = [];
  let current = null;

  for (const rawLine of lines) {
    const line = normalizeText(rawLine);
    if (!line) {
      continue;
    }
    if (/^(from|to|subject|date):/i.test(line)) {
      continue;
    }
    if (/^on .* wrote:$/i.test(line) || /^sent from my iphone$/i.test(line) || /^<image/i.test(line)) {
      continue;
    }

    if (isLikelyHeading(line)) {
      current = { heading: line, values: [] };
      sections.push(current);
      continue;
    }

    if (current) {
      current.values.push(line.replace(/^[•*-]\s*/, ""));
    }
  }

  return sections.filter((section) => section.values.length > 0);
}

function titleCaseWords(values) {
  return values.map((value) =>
    value
      .split(/\s+/)
      .filter(Boolean)
      .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
      .join(" ")
  );
}

function extractLocations(sourceText) {
  const text = ` ${String(sourceText || "").toLowerCase()} `;
  const placeMap = new Map([
    ["faversham", "Faversham"],
    ["tenyham", "Tenyham"],
    ["teynham", "Teynham"],
    ["hernebay", "Hernebay"],
    ["herne bay", "Herne Bay"],
    ["ashford", "Ashford"],
    ["lenham", "Lenham"],
    ["kent", "Kent"],
  ]);

  const found = [];
  for (const [needle, label] of placeMap.entries()) {
    if (text.includes(` ${needle} `)) {
      found.push(label);
    }
  }

  return Array.from(new Set(found));
}

function inferCounty(sourceText, countyField) {
  if (!countyField?.choices?.length) {
    return "";
  }

  const locations = extractLocations(sourceText).map((value) => value.toLowerCase());
  if (locations.some((value) => ["faversham", "teynham", "tenyham", "herne bay", "hernebay", "ashford", "lenham", "kent"].includes(value))) {
    const kentChoice = countyField.choices.find((choice) => normalizeText(choice).toLowerCase() === "kent");
    return kentChoice || "";
  }

  return "";
}

function inferTown(sourceText, townField) {
  if (!townField?.choices?.length) {
    return "";
  }
  const lowerText = String(sourceText || "").toLowerCase();
  return townField.choices.find((choice) => lowerText.includes(normalizeText(choice).toLowerCase())) || "";
}

function inferSupplierType(sourceText, supplierTypeField) {
  if (!supplierTypeField?.choices?.length) {
    const lowerText = String(sourceText || "").toLowerCase();
    if (lowerText.includes("namaste")) {
      return "Namaste";
    }
    if (lowerText.includes("therapy")) {
      return "Therapy";
    }
    return "";
  }

  const text = String(sourceText || "").toLowerCase();
  const exact = supplierTypeField.choices.find((choice) => text.includes(normalizeText(choice).toLowerCase()));
  if (exact) {
    return exact;
  }

  if (/\bhome visits?\b/.test(text)) {
    const domCare = supplierTypeField.choices.find((choice) => normalizeText(choice).toLowerCase() === "dom care");
    if (domCare) {
      return domCare;
    }
  }

  if (text.includes("namaste")) {
    return "Namaste";
  }

  return "";
}

function inferTags(sourceText, tagsField) {
  const text = String(sourceText || "").toLowerCase();
  const matched = [];
  const candidates = Array.isArray(tagsField?.choices) ? tagsField.choices : [];

  for (const choice of candidates) {
    const normalizedChoice = normalizeText(choice).toLowerCase();
    if (normalizedChoice && text.includes(normalizedChoice)) {
      matched.push(choice);
    }
  }

  if (!matched.length && text.includes("namaste")) {
    matched.push("namaste");
  }

  return Array.from(new Set(matched)).join(", ");
}

function extractPricePhrase(sourceText) {
  const matched = String(sourceText || "").match(/£\s*\d+(?:\.\d{1,2})?(?:\s*(?:plus|per|\/)\s*[A-Za-z ]+)?/i);
  return matched ? normalizeText(matched[0].replace(/\s+/g, " ")) : "";
}

function buildNotesSummary(sourceText) {
  const structuredSections = extractStructuredSections(sourceText);
  if (structuredSections.length) {
    return structuredSections
      .map((section) => `${section.heading}: ${section.values.join("; ")}`)
      .join("\n");
  }

  const bodyText = extractBodyText(sourceText);
  const sentences = [];

  if (/namaste/i.test(bodyText)) {
    sentences.push("Offers namaste services.");
  }
  if (/\bhome visits?\b/i.test(bodyText)) {
    sentences.push("Provides home visits.");
  }

  const locations = extractLocations(bodyText).filter((value) => value.toLowerCase() !== "kent");
  if (locations.length) {
    sentences.push(`Coverage area includes ${titleCaseWords(locations).join(", ")}.`);
  }

  const price = extractPricePhrase(bodyText);
  if (price) {
    sentences.push(`Charges ${price}.`);
  }

  if (!sentences.length && bodyText) {
    sentences.push(bodyText);
  }

  return sentences.join(" ");
}

function buildContactDetails(sourceText) {
  const parts = [];
  const name = extractName(sourceText);
  const phone = extractPhone(sourceText);
  const email = extractEmail(sourceText);

  if (name) {
    parts.push(name);
  }
  if (phone) {
    parts.push(phone);
  }
  if (email) {
    parts.push(email);
  }

  return parts.join(" - ");
}

function mergeTextBlocks(primary, secondary) {
  const left = normalizeText(primary);
  const right = normalizeText(secondary);
  if (left && right) {
    if (left.toLowerCase() === right.toLowerCase()) {
      return left;
    }
    if (left.toLowerCase().includes(right.toLowerCase())) {
      return left;
    }
    if (right.toLowerCase().includes(left.toLowerCase())) {
      return right;
    }
    return `${left}\n${right}`;
  }
  return left || right || "";
}

function applyHeuristics({ values, fields, sourceText }) {
  const next = values && typeof values === "object" ? { ...values } : {};
  const titleField = fieldByTitle(fields, ["Title"]);
  const contactField = fieldByTitle(fields, ["Contact Details"]);
  const countyField = fieldByTitle(fields, ["County"]);
  const notesField = fieldByTitle(fields, ["Notes"]);
  const townField = fieldByTitle(fields, ["Town"]);
  const websiteField = fieldByTitle(fields, ["Website"]);
  const supplierTypeField = fieldByTitle(fields, ["Supplier Type"]);
  const tagsField = fieldByTitle(fields, ["Tags"]);
  const structuredSections = extractStructuredSections(sourceText);
  const serviceName = normalizeText(
    structuredSections.find((section) => normalizeText(section.heading).toLowerCase() === "service name")?.values?.[0]
  );

  if (titleField?.internalName) {
    const title = normalizeText(next[titleField.internalName]) || serviceName || extractName(sourceText);
    if (title) {
      next[titleField.internalName] = title;
    }
  }

  if (contactField?.internalName) {
    const contactDetails = buildContactDetails(sourceText);
    if (contactDetails) {
      next[contactField.internalName] = mergeTextBlocks(next[contactField.internalName], contactDetails);
    }
  }

  if (countyField?.internalName && !normalizeText(next[countyField.internalName])) {
    const county = inferCounty(sourceText, countyField);
    if (county) {
      next[countyField.internalName] = county;
    }
  }

  if (townField?.internalName && !normalizeText(next[townField.internalName])) {
    const town = inferTown(sourceText, townField);
    if (town) {
      next[townField.internalName] = town;
    }
  }

  if (supplierTypeField?.internalName && !normalizeText(next[supplierTypeField.internalName])) {
    const supplierType = inferSupplierType(sourceText, supplierTypeField);
    if (supplierType) {
      next[supplierTypeField.internalName] = supplierType;
    }
  }

  if (tagsField?.internalName && !normalizeText(next[tagsField.internalName])) {
    const tags = inferTags(sourceText, tagsField);
    if (tags) {
      next[tagsField.internalName] = tags;
    }
  }

  if (websiteField?.internalName) {
    const url = normalizeText(next[websiteField.internalName]) || extractUrl(sourceText);
    if (url && /^https?:\/\//i.test(url)) {
      next[websiteField.internalName] = url;
    } else if (!normalizeText(next[websiteField.internalName])) {
      next[websiteField.internalName] = null;
    }
  }

  if (notesField?.internalName) {
    const notes = buildNotesSummary(sourceText);
    if (notes) {
      next[notesField.internalName] = mergeTextBlocks(next[notesField.internalName], notes);
    }
  }

  return next;
}

async function callDeepSeek({ fields, sourceText }) {
  const apiKey = normalizeText(process.env.DEEPSEEK_API_KEY);
  if (!apiKey) {
    const error = new Error("Server missing DEEPSEEK_API_KEY for AI extraction.");
    error.status = 500;
    throw error;
  }

  const model = normalizeText(process.env.DEEPSEEK_MODEL || DEFAULT_MODEL);
  const response = await fetch("https://api.deepseek.com/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      messages: [
        {
          role: "system",
          content:
            "You extract structured list-entry drafts from free text. Return valid json only and match the requested object shape.",
        },
        {
          role: "user",
          content: buildPrompt({ sourceText, fields }),
        },
      ],
      response_format: {
        type: "json_object",
      },
      max_tokens: 1200,
    }),
  });

  const payload = await response.json().catch(() => null);
  if (!response.ok) {
    const detail = payload?.error?.message || `DeepSeek request failed (${response.status}).`;
    const error = new Error(detail);
    error.status = response.status;
    throw error;
  }

  const content = normalizeText(payload?.choices?.[0]?.message?.content) || extractResponseText(payload);
  return parseJsonContent(content);
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

  const sourceText = normalizeText(req.body?.sourceText);
  const fields = sanitizeFields(req.body?.fields);

  if (!sourceText) {
    res.status(400).json({
      error: {
        code: "BAD_REQUEST",
        message: "Pasted text is required.",
      },
    });
    return;
  }

  if (!fields.length) {
    res.status(400).json({
      error: {
        code: "BAD_REQUEST",
        message: "At least one supported field is required.",
      },
    });
    return;
  }

  try {
    const aiValues = await callDeepSeek({ fields, sourceText });
    const values = applyHeuristics({ values: aiValues, fields, sourceText });
    res.status(200).json({
      success: true,
      values,
    });
  } catch (error) {
    res.status(Number(error?.status) || 500).json({
      error: {
        code: "WELLBEING_INTAKE_PARSE_FAILED",
        message: error?.message || "Could not extract a SharePoint entry from the pasted text.",
      },
    });
  }
};
