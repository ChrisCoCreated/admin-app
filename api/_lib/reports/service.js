const fs = require("fs/promises");
const { createChatCompletion } = require("../ai-client");
const { normalizeText } = require("../deepseek-client");
const { REPORT_STATUS_VALUES, getReportType } = require("./registry");

function safeJson(value) {
  return JSON.stringify(value, null, 2);
}

function extractJsonObject(text) {
  const raw = normalizeText(text);
  if (!raw) {
    const error = new Error("AI returned an empty report response.");
    error.status = 502;
    throw error;
  }
  try {
    return JSON.parse(raw);
  } catch (_error) {
    const start = raw.indexOf("{");
    const end = raw.lastIndexOf("}");
    if (start >= 0 && end > start) {
      try {
        return JSON.parse(raw.slice(start, end + 1));
      } catch {
        const error = new Error("AI report response was not valid JSON.");
        error.status = 502;
        throw error;
      }
    }
    const error = new Error("AI report response was not valid JSON.");
    error.status = 502;
    throw error;
  }
}

function normalizeStringArray(value) {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.map((item) => normalizeText(item)).filter(Boolean);
}

function normalizeReportJson(payload, reportType) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    const error = new Error("AI report response was not a JSON object.");
    error.status = 502;
    throw error;
  }

  const status = normalizeText(payload.status);
  if (!REPORT_STATUS_VALUES.has(status)) {
    const error = new Error("AI report response used an invalid status.");
    error.status = 502;
    throw error;
  }

  const sections = payload.report_sections && typeof payload.report_sections === "object" ? payload.report_sections : {};
  return {
    status,
    report_type: reportType.report_type,
    report_title: normalizeText(payload.report_title) || reportType.display_name,
    client_details: payload.client_details && typeof payload.client_details === "object" ? payload.client_details : {},
    inferred_context:
      payload.inferred_context && typeof payload.inferred_context === "object"
        ? payload.inferred_context
        : { assessment_type: "unclear", evidence: [] },
    suggested_smart_goals: Array.isArray(payload.suggested_smart_goals) ? payload.suggested_smart_goals : [],
    report_sections: {
      executive_summary: normalizeText(sections.executive_summary),
      current_situation: normalizeText(sections.current_situation),
      physical_wellbeing: normalizeText(sections.physical_wellbeing),
      emotional_wellbeing: normalizeText(sections.emotional_wellbeing),
      environmental_wellbeing: normalizeText(sections.environmental_wellbeing),
      wellbeing_highlights: normalizeText(sections.wellbeing_highlights),
      recommendations: normalizeStringArray(sections.recommendations),
      next_steps: normalizeStringArray(sections.next_steps),
    },
    omitted_sections: normalizeStringArray(payload.omitted_sections),
    assumptions_avoided: normalizeStringArray(payload.assumptions_avoided),
    clarification_notes: normalizeStringArray(payload.clarification_notes),
    source_notes_used: normalizeStringArray(payload.source_notes_used),
    warnings: normalizeStringArray(payload.warnings),
    tone_check: payload.tone_check && typeof payload.tone_check === "object" ? payload.tone_check : {},
    revision_prompt: normalizeText(payload.revision_prompt),
  };
}

async function readExampleDocumentText(reportType) {
  try {
    const buffer = await fs.readFile(reportType.example_document);
    return `Example document available at ${reportType.example_document}. Size: ${buffer.length} bytes.`;
  } catch {
    return `Example document path: ${reportType.example_document}.`;
  }
}

async function buildReportMessages({ reportType, notes, previousReport, revisionRequest }) {
  const exampleSummary = await readExampleDocumentText(reportType);
  const isRevision = previousReport && typeof previousReport === "object" && normalizeText(revisionRequest);

  return [
    {
      role: "system",
      content: [
        ...reportType.instructions,
        "",
        "JSON schema:",
        safeJson(reportType.json_schema),
        "",
        "Template mapping:",
        safeJson(reportType.template_mapping),
        "",
        exampleSummary,
      ].join("\n"),
    },
    {
      role: "user",
      content: [
        `Report type: ${reportType.report_type}`,
        isRevision ? "Task: Revise the previous structured report." : "Task: Generate the structured report.",
        "",
        "Original pseudonymised notes:",
        normalizeText(notes) || "(none)",
        "",
        isRevision ? `User requested changes:\n${normalizeText(revisionRequest)}` : "",
        isRevision ? `Previous structured JSON:\n${safeJson(previousReport)}` : "",
      ]
        .filter((part) => part !== "")
        .join("\n"),
    },
  ];
}

async function generateStructuredReport({
  reportType: reportTypeKey,
  notes,
  provider,
  model,
  thinking,
  reasoningEffort,
  previousReport,
  revisionRequest,
  createCompletion = createChatCompletion,
}) {
  const reportType = getReportType(reportTypeKey);
  const cleanNotes = normalizeText(notes);
  if (!cleanNotes) {
    return normalizeReportJson(
      {
        status: "needs_notes",
        report_type: reportType.report_type,
        report_title: reportType.display_name,
        client_details: {},
        inferred_context: { assessment_type: "unclear", evidence: [] },
        suggested_smart_goals: [],
        report_sections: {},
        omitted_sections: [],
        assumptions_avoided: [],
        clarification_notes: ["Please provide the notes before generating the report."],
        source_notes_used: [],
        warnings: [],
        tone_check: {},
        revision_prompt: "Provide notes to generate the report.",
      },
      reportType
    );
  }

  const completion = await createCompletion({
    provider,
    model,
    thinking,
    reasoningEffort,
    maxTokens: 2500,
    temperature: 0.2,
    responseFormat: { type: "json_object" },
    messages: await buildReportMessages({
      reportType,
      notes: cleanNotes,
      previousReport,
      revisionRequest,
    }),
  });

  const parsed = extractJsonObject(completion.content || completion.text);
  return normalizeReportJson(parsed, reportType);
}

module.exports = {
  buildReportMessages,
  generateStructuredReport,
  normalizeReportJson,
};
