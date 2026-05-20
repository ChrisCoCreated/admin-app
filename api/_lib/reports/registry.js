const path = require("path");

const REPORT_STATUS_VALUES = new Set(["ready_for_render", "needs_notes", "needs_clarification"]);

const WELLBEING_ASSURANCE_VISIT_SCHEMA = {
  type: "object",
  required: [
    "status",
    "report_type",
    "report_title",
    "client_details",
    "inferred_context",
    "suggested_smart_goals",
    "report_sections",
    "omitted_sections",
    "assumptions_avoided",
    "clarification_notes",
    "source_notes_used",
    "warnings",
    "tone_check",
    "revision_prompt",
  ],
  properties: {
    status: { enum: Array.from(REPORT_STATUS_VALUES) },
    report_type: { const: "wellbeing_assurance_visit" },
    report_title: { type: "string" },
    client_details: { type: "object" },
    inferred_context: {
      type: "object",
      properties: {
        assessment_type: { enum: ["initial_assessment", "review", "unclear"] },
        evidence: { type: "array", items: { type: "string" } },
      },
      additionalProperties: true,
    },
    suggested_smart_goals: { type: "array", items: { type: "object" } },
    report_sections: {
      type: "object",
      properties: {
        executive_summary: { type: "string" },
        current_situation: { type: "string" },
        physical_wellbeing: { type: "string" },
        emotional_wellbeing: { type: "string" },
        environmental_wellbeing: { type: "string" },
        wellbeing_highlights: { type: "string" },
        recommendations: { type: "array", items: { type: "string" } },
        next_steps: { type: "array", items: { type: "string" } },
      },
      additionalProperties: true,
    },
    omitted_sections: { type: "array", items: { type: "string" } },
    assumptions_avoided: { type: "array", items: { type: "string" } },
    clarification_notes: { type: "array", items: { type: "string" } },
    source_notes_used: { type: "array", items: { type: "string" } },
    warnings: { type: "array", items: { type: "string" } },
    tone_check: { type: "object" },
    revision_prompt: { type: "string" },
  },
  additionalProperties: true,
};

const WELLBEING_ASSURANCE_VISIT = {
  report_type: "wellbeing_assurance_visit",
  display_name: "Wellbeing Assurance Visit Summary",
  example_document: path.join(process.cwd(), "assets", "documents", "Wellbeing_Assurance_Visit_Example.docx"),
  template_document: path.join(process.cwd(), "assets", "documents", "Wellbeing_Assurance_Report_Template.docx"),
  json_schema: WELLBEING_ASSURANCE_VISIT_SCHEMA,
  template_mapping: {
    client_name: "client_details.client_name",
    report_date: "client_details.report_date",
    conducted_by: "client_details.conducted_by",
    assessment_venue: "client_details.assessment_venue",
    date_of_birth: "client_details.date_of_birth",
    prepared_for: "client_details.prepared_for",
    assessment_type: "inferred_context.assessment_type",
    executive_summary: "report_sections.executive_summary",
    current_situation: "report_sections.current_situation",
    physical_wellbeing: "report_sections.physical_wellbeing",
    emotional_wellbeing: "report_sections.emotional_wellbeing",
    environmental_wellbeing: "report_sections.environmental_wellbeing",
    wellbeing_highlights: "report_sections.wellbeing_highlights",
    recommendations: "report_sections.recommendations",
    next_steps: "report_sections.next_steps",
    smart_goals: "suggested_smart_goals",
  },
  instructions: [
    "You generate a Wellbeing Assurance Visit Summary from pseudonymised notes.",
    "Return structured JSON only. Do not include markdown, prose outside JSON, or a final Word document.",
    "If no notes are provided, return status needs_notes and ask the user to provide notes.",
    "If notes are too limited, contradictory, or unclear to safely generate a meaningful report, return status needs_clarification.",
    "Otherwise return status ready_for_render and generate the report in one pass.",
    "Infer whether the notes appear to describe an initial assessment where possible.",
    "If the notes clearly or strongly indicate an initial assessment, omit the Wellbeing Highlights section and list it in omitted_sections.",
    "If assessment type is unclear, include Wellbeing Highlights only when there is enough evidence in the notes.",
    "Use direct quotes where possible, but only when the words appear in the notes.",
    "Never invent facts, experiences, preferences, risks, outcomes, relationships, routines, or professional views.",
    "Do not extrapolate beyond the information provided.",
    "Include clear warnings or clarification notes for missing or unclear non-essential information.",
    "Write in a relaxed, professional, concise, empathetic style.",
    "The executive summary must be brief and suitable for a non-specialist reader.",
    "Include detail from every meaningful part of the notes.",
    "Suggest SMART goals based only on the notes. Make clear they are suggested and can be reviewed or amended.",
    "Preserve every bracketed placeholder exactly as written.",
  ],
};

const SIMPLE_SUMMARY = {
  report_type: "simple_summary",
  display_name: "Simple Summary Report",
  example_document: "",
  template_document: "",
  output_mode: "plain_text_summary",
  json_schema: {
    ...WELLBEING_ASSURANCE_VISIT_SCHEMA,
    properties: {
      ...WELLBEING_ASSURANCE_VISIT_SCHEMA.properties,
      report_type: { const: "simple_summary" },
    },
  },
  template_mapping: {},
  instructions: [
    "You generate a simple summary report from pseudonymised care/admin notes.",
    "Use clear, professional language suitable for an internal care/admin note.",
    "Summarise only information present in the notes.",
    "Do not add clinical facts, names, dates, locations, actions, risks, or opinions that are not in the notes.",
    "Keep all bracketed placeholders exactly as written.",
    "Return only the report text, with no preamble.",
  ],
};

const REPORT_TYPES = new Map([
  [WELLBEING_ASSURANCE_VISIT.report_type, WELLBEING_ASSURANCE_VISIT],
  [SIMPLE_SUMMARY.report_type, SIMPLE_SUMMARY],
]);

function listReportTypes() {
  return Array.from(REPORT_TYPES.values()).map((reportType) => ({
    report_type: reportType.report_type,
    display_name: reportType.display_name,
  }));
}

function getReportType(reportType) {
  const key = String(reportType || "").trim();
  const config = REPORT_TYPES.get(key);
  if (!config) {
    const error = new Error(`Unsupported report type "${key || "unknown"}".`);
    error.status = 400;
    throw error;
  }
  return config;
}

module.exports = {
  REPORT_STATUS_VALUES,
  getReportType,
  listReportTypes,
};
