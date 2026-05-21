const test = require("node:test");
const assert = require("node:assert/strict");
const JSZip = require("jszip");

const { getReportType } = require("../api/_lib/reports/registry");
const { buildReportMessages, generateStructuredReport } = require("../api/_lib/reports/service");
const { buildReportDocx } = require("../api/_lib/reports/docx");
const reportDocxHandler = require("../api/reports/docx");

async function docxText(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml").async("string");
  return xml.replace(/<[^>]+>/g, "");
}

test("report registry exposes wellbeing assurance visit configuration", () => {
  const reportType = getReportType("wellbeing_assurance_visit");
  assert.equal(reportType.display_name, "Wellbeing Assurance Visit Summary");
  assert.ok(reportType.example_document.endsWith("Wellbeing_Assurance_Visit_Example.docx"));
  assert.ok(reportType.template_document.endsWith("Wellbeing_Assurance_Report_Template.docx"));
  assert.equal(reportType.json_schema.properties.status.enum.includes("ready_for_render"), true);
  assert.equal(reportType.template_mapping.executive_summary, "report_sections.executive_summary");
  assert.match(reportType.instructions.join("\n"), /preferred-name placeholder/);
});

test("report registry exposes simple summary configuration", () => {
  const reportType = getReportType("simple_summary");
  assert.equal(reportType.display_name, "Simple Summary Report");
  assert.equal(reportType.example_document, "");
  assert.equal(reportType.json_schema.properties.report_type.const, "simple_summary");
  assert.match(reportType.instructions.join("\n"), /simple summary/i);
  assert.match(reportType.instructions.join("\n"), /preferred-name placeholder/);
});

test("report registry rejects unknown report types", () => {
  assert.throws(() => getReportType("unknown_report"), /Unsupported report type/);
});

test("structured report generation returns needs_notes without calling AI for empty notes", async () => {
  let called = false;
  const report = await generateStructuredReport({
    reportType: "wellbeing_assurance_visit",
    notes: "",
    createCompletion: async () => {
      called = true;
      throw new Error("Should not call AI");
    },
  });

  assert.equal(called, false);
  assert.equal(report.status, "needs_notes");
  assert.equal(report.report_type, "wellbeing_assurance_visit");
});

test("structured report generation parses ready_for_render JSON from AI", async () => {
  let capturedOptions = null;
  const report = await generateStructuredReport({
    reportType: "wellbeing_assurance_visit",
    notes: "[CLIENT_001] feels safe.",
    createCompletion: async (options) => {
      capturedOptions = options;
      return {
        content: JSON.stringify({
          status: "ready_for_render",
          report_type: "wellbeing_assurance_visit",
          report_title: "Wellbeing Assurance Visit Summary",
          client_details: { client_name: "[CLIENT_001]" },
          inferred_context: { assessment_type: "review", evidence: ["feels safe"] },
          suggested_smart_goals: [{ goal: "Review companionship options", measurable: "One option agreed" }],
          report_sections: {
            executive_summary: "[CLIENT_001] feels safe.",
            current_situation: "Current situation text.",
            physical_wellbeing: "",
            emotional_wellbeing: "",
            environmental_wellbeing: "",
            wellbeing_highlights: "Feels safe.",
            recommendations: ["Continue support"],
            next_steps: ["Review goals"],
          },
          omitted_sections: [],
          assumptions_avoided: [],
          clarification_notes: [],
          source_notes_used: ["[CLIENT_001] feels safe."],
          warnings: [],
          tone_check: { professional: true },
          revision_prompt: "Request changes if needed.",
        }),
      };
    },
  });

  assert.deepEqual(capturedOptions.responseFormat, { type: "json_object" });
  assert.equal(capturedOptions.maxTokens, 6000);
  assert.match(capturedOptions.messages[0].content, /exactly one valid JSON object/);
  assert.match(capturedOptions.messages[0].content, /JSON\.parse/);
  assert.equal(report.status, "ready_for_render");
  assert.equal(report.client_details.client_name, "[CLIENT_001]");
  assert.equal(report.suggested_smart_goals.length, 1);
});

test("structured report generation omits JSON response format for Azure Foundry", async () => {
  let capturedOptions = null;
  const report = await generateStructuredReport({
    reportType: "wellbeing_assurance_visit",
    notes: "[CLIENT_001] feels safe.",
    provider: "azure_openai",
    model: "primary",
    createCompletion: async (options) => {
      capturedOptions = options;
      return {
        content: JSON.stringify({
          status: "ready_for_render",
          report_type: "wellbeing_assurance_visit",
          report_title: "Wellbeing Assurance Visit Summary",
          client_details: { client_name: "[CLIENT_001]" },
          inferred_context: { assessment_type: "review", evidence: [] },
          suggested_smart_goals: [],
          report_sections: {
            executive_summary: "[CLIENT_001] feels safe.",
            current_situation: "",
            physical_wellbeing: "",
            emotional_wellbeing: "",
            environmental_wellbeing: "",
            wellbeing_highlights: "",
            recommendations: [],
            next_steps: [],
          },
          omitted_sections: [],
          assumptions_avoided: [],
          clarification_notes: [],
          source_notes_used: ["[CLIENT_001] feels safe."],
          warnings: [],
          tone_check: {},
          revision_prompt: "",
        }),
      };
    },
  });

  assert.equal(capturedOptions.responseFormat, undefined);
  assert.equal(capturedOptions.maxTokens, 6000);
  assert.match(capturedOptions.messages[0].content, /JSON\.parse/);
  assert.equal(report.status, "ready_for_render");
});

test("simple summary generation uses plain text AI response without JSON response format", async () => {
  let capturedOptions = null;
  const report = await generateStructuredReport({
    reportType: "simple_summary",
    notes: "[CLIENT_001] feels safe.",
    provider: "azure_openai",
    model: "fast",
    createCompletion: async (options) => {
      capturedOptions = options;
      return {
        content: "[CLIENT_001] feels safe and staff support is positive.",
      };
    },
  });

  assert.equal(capturedOptions.responseFormat, undefined);
  assert.equal(capturedOptions.maxTokens, 1200);
  assert.equal(report.status, "ready_for_render");
  assert.equal(report.report_type, "simple_summary");
  assert.equal(report.report_sections.executive_summary, "[CLIENT_001] feels safe and staff support is positive.");
});

test("structured report generation rejects malformed model JSON", async () => {
  await assert.rejects(
    () =>
      generateStructuredReport({
        reportType: "wellbeing_assurance_visit",
        notes: "Some notes",
        createCompletion: async () => ({ content: "not json" }),
      }),
    /Unexpected token|AI report response/
  );
});

test("revision prompt includes original notes, previous JSON, and requested changes", async () => {
  const reportType = getReportType("wellbeing_assurance_visit");
  const messages = await buildReportMessages({
    reportType,
    notes: "Original notes",
    guidance: "Keep it short and suitable for a family update.",
    audience: "Family member",
    writingStyle: "Plain English",
    previousReport: { status: "ready_for_render", report_type: "wellbeing_assurance_visit" },
    revisionRequest: "Make it warmer",
  });
  const userMessage = messages[1].content;
  assert.match(userMessage, /Original notes/);
  assert.match(userMessage, /Make it warmer/);
  assert.match(userMessage, /Previous structured JSON/);
  assert.match(userMessage, /Family member/);
  assert.match(userMessage, /Plain English/);
  assert.match(userMessage, /Keep it short/);
});

test("report docx replaces placeholders, renders goals, and removes omitted sections and notes", async () => {
  const output = await buildReportDocx({
    reportType: "wellbeing_assurance_visit",
    report: {
      report_type: "wellbeing_assurance_visit",
      report_title: "Wellbeing Assurance Visit Summary",
      client_details: {
        client_name: "Paul Jones",
        report_date: "20.05.2026",
        conducted_by: "Thrive",
        assessment_venue: "Mossbank",
        date_of_birth: "",
        prepared_for: "Family",
      },
      inferred_context: { assessment_type: "initial_assessment" },
      suggested_smart_goals: [
        {
          goal: "Support confidence at home",
          measurable: "One weekly check-in completed",
          owner: "Care team",
          time_bound: "Four weeks",
        },
      ],
      report_sections: {
        executive_summary: "Paul feels safe.",
        current_situation: "Current situation.",
        physical_wellbeing: "Physical wellbeing.",
        emotional_wellbeing: "Emotional wellbeing.",
        environmental_wellbeing: "Environmental wellbeing.",
        wellbeing_highlights: "Should be omitted.",
        recommendations: ["Continue support"],
        next_steps: ["Review plan"],
      },
      omitted_sections: ["Wellbeing Highlights"],
      warnings: ["Limited information provided."],
      clarification_notes: ["Physical wellbeing details are missing."],
    },
  });
  const text = await docxText(output);
  assert.match(text, /Paul Jones/);
  assert.match(text, /Paul feels safe/);
  assert.match(text, /Support confidence at home/);
  assert.doesNotMatch(text, /Limited information provided/);
  assert.doesNotMatch(text, /Physical wellbeing details are missing/);
  assert.doesNotMatch(text, /Wellbeing Highlights/);
  assert.doesNotMatch(text, /Implementation Notes/);
  assert.doesNotMatch(text, /\{\{/);
});

test("report docx filename includes client name and text month date", () => {
  const filename = reportDocxHandler.buildReportFilename(
    {
      report_title: "Wellbeing Assurance Visit Summary",
      client_details: {
        client_name: "Paul Jones",
        report_date: "20.05.2026",
      },
    },
    new Date(2026, 4, 21)
  );

  assert.equal(filename, "Paul Jones - Wellbeing Assurance Visit Summary - 20 May 2026.docx");
});

test("report docx filename falls back to today when report date is missing", () => {
  const filename = reportDocxHandler.buildReportFilename(
    {
      report_title: "Wellbeing Assurance Visit Summary",
      client_details: {
        client_name: "Claire Smith",
      },
    },
    new Date(2026, 4, 21)
  );

  assert.equal(filename, "Claire Smith - Wellbeing Assurance Visit Summary - 21 May 2026.docx");
});

test("report docx filename can derive client name from title and remove duplicate suffix", () => {
  const filename = reportDocxHandler.buildReportFilename(
    {
      report_title: "Wellbeing Assurance Visit Summary for Paul Jones",
      client_details: {},
    },
    new Date(2026, 4, 21)
  );

  assert.equal(filename, "Paul Jones - Wellbeing Assurance Visit Summary - 21 May 2026.docx");
});

test("report docx content disposition is header safe with unicode filename", () => {
  const header = reportDocxHandler.contentDispositionForFilename(
    'Paul Jones – "Warm" Wellbeing Summary, v2 - 21 May 2026.docx'
  );

  assert.doesNotThrow(() => {
    const { validateHeaderValue } = require("node:http");
    validateHeaderValue("Content-Disposition", header);
  });
  assert.match(header, /filename\*=UTF-8''/);
  assert.match(header, /Paul%20Jones/);
});
