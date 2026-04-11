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
    "Only use facts present in the text.",
    "If a field is not clearly present, return null.",
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
    const values = await callDeepSeek({ fields, sourceText });
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
