const {
  createProblemRow,
  ensureUuid,
  getProblemRowByIdAndOwner,
  listProblemRowsByOwnerEmail,
  normalizeEmail,
  updateProblemRow,
} = require("./repository");

const TYPE_VALUES = new Set(["concern", "opportunity", "decision"]);
const PRIORITY_VALUES = new Set(["low", "medium", "high"]);
const STATE_VALUES = new Set(["new", "in_progress", "done", "parked"]);

function cleanText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function sentenceCase(value) {
  const text = cleanText(value);
  if (!text) {
    return "";
  }
  return text.charAt(0).toUpperCase() + text.slice(1);
}

function clampText(value, maxLength) {
  const text = cleanText(value);
  if (text.length <= maxLength) {
    return text;
  }
  return `${text.slice(0, Math.max(0, maxLength - 1)).trimEnd()}…`;
}

function normalizeProblemType(value) {
  const text = cleanText(value).toLowerCase();
  return TYPE_VALUES.has(text) ? text : "";
}

function normalizePriority(value) {
  const text = cleanText(value).toLowerCase();
  return PRIORITY_VALUES.has(text) ? text : "";
}

function normalizeState(value) {
  const text = cleanText(value).toLowerCase();
  return STATE_VALUES.has(text) ? text : "";
}

function detectProblemType(...parts) {
  const text = cleanText(parts.filter(Boolean).join(" ")).toLowerCase();
  if (!text) {
    return "concern";
  }

  if (
    /\b(decide|decision|choose|choice|which option|pick between|trade[- ]?off|compare options)\b/.test(text)
  ) {
    return "decision";
  }

  if (
    /\b(improve|improvement|better|could|opportunity|optimi[sz]e|streamline|enhance|grow|upgrade)\b/.test(text)
  ) {
    return "opportunity";
  }

  if (
    /\b(problem|issue|wrong|broken|stuck|blocked|delay|risk|concern|bother|bug|friction|failing)\b/.test(text)
  ) {
    return "concern";
  }

  return "concern";
}

function generateTitle(originalInput) {
  const source = cleanText(originalInput)
    .replace(/^(i need to|we need to|need to|there(?:'s| is)|problem with|issue with|concern about)\s+/i, "")
    .replace(/^[^a-z0-9]+/i, "");

  if (!source) {
    return "Untitled problem";
  }

  const firstChunk = source.split(/[.!?;:]/)[0] || source;
  const words = cleanText(firstChunk).split(" ").filter(Boolean);
  const concise = words.slice(0, 8).join(" ");
  return sentenceCase(clampText(concise || source, 64));
}

function deriveFocusPhrase(originalInput, clarification) {
  const clarificationText = cleanText(clarification);
  if (clarificationText) {
    return clarificationText;
  }

  const source = cleanText(originalInput)
    .replace(/^(i need to|we need to|need to|there(?:'s| is)|problem with|issue with|concern about)\s+/i, "")
    .replace(/[.?!].*$/, "");

  return source || "this";
}

function uniqueStrings(values) {
  const seen = new Set();
  const result = [];

  for (const value of values) {
    const text = sentenceCase(value);
    const key = text.toLowerCase();
    if (!text || seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(text);
  }

  return result;
}

function generateReframes(originalInput, clarification, problemType) {
  const type = normalizeProblemType(problemType) || detectProblemType(originalInput, clarification);
  const focus = deriveFocusPhrase(originalInput, clarification).replace(/[?!.]+$/g, "");

  if (type === "opportunity") {
    return uniqueStrings([
      `How might we define ${focus} more clearly?`,
      `How might we improve ${focus} in a way users notice quickly?`,
      `How might we test the best version of ${focus} with low effort?`,
    ]).slice(0, 3);
  }

  if (type === "decision") {
    return uniqueStrings([
      `How might we compare the options for ${focus} more clearly?`,
      `How might we make the trade-offs around ${focus} visible?`,
      `How might we decide on ${focus} with enough confidence to move?`,
    ]).slice(0, 3);
  }

  return uniqueStrings([
    `How might we understand what is driving ${focus}?`,
    `How might we reduce the impact of ${focus} quickly?`,
    `How might we stop ${focus} from happening again?`,
  ]).slice(0, 3);
}

function suggestNextStep(problemType) {
  const type = normalizeProblemType(problemType) || "concern";
  if (type === "opportunity") {
    return "Define the opportunity clearly";
  }
  if (type === "decision") {
    return "List and compare options";
  }
  return "Gather examples / understand root cause";
}

function parseReframes(value, fallbackOriginalInput, fallbackClarification, fallbackType) {
  const source = Array.isArray(value)
    ? value
    : typeof value === "string"
      ? value.split(/\r?\n+/)
      : [];

  const cleaned = source
    .map((entry) => clampText(entry, 160))
    .filter(Boolean)
    .slice(0, 3);

  if (cleaned.length) {
    return cleaned;
  }

  return generateReframes(fallbackOriginalInput, fallbackClarification, fallbackType);
}

function mapProblemRow(row) {
  return {
    id: row?.id || "",
    ownerEmail: normalizeEmail(row?.owner_email || ""),
    ownerName: cleanText(row?.owner_name || ""),
    title: cleanText(row?.title || ""),
    originalInput: cleanText(row?.original_input || ""),
    problemType: normalizeProblemType(row?.problem_type || "") || "concern",
    clarification: cleanText(row?.clarification || ""),
    reframes: Array.isArray(row?.reframes) ? row.reframes.map((entry) => cleanText(entry)).filter(Boolean).slice(0, 3) : [],
    nextStep: cleanText(row?.next_step || ""),
    priority: normalizePriority(row?.priority || "") || "medium",
    state: normalizeState(row?.state || "") || "new",
    createdAt: row?.created_at || null,
    updatedAt: row?.updated_at || null,
  };
}

function priorityRank(priority) {
  if (priority === "high") {
    return 3;
  }
  if (priority === "medium") {
    return 2;
  }
  return 1;
}

function sortProblems(problems) {
  return [...problems].sort((a, b) => {
    const priorityDelta = priorityRank(b.priority) - priorityRank(a.priority);
    if (priorityDelta !== 0) {
      return priorityDelta;
    }

    const aUpdated = Date.parse(a.updatedAt || a.createdAt || "") || 0;
    const bUpdated = Date.parse(b.updatedAt || b.createdAt || "") || 0;
    if (aUpdated !== bUpdated) {
      return bUpdated - aUpdated;
    }

    return String(a.title || "").localeCompare(String(b.title || ""), undefined, { sensitivity: "base" });
  });
}

function buildProblemPayload(input = {}, ownerEmail, existing = null) {
  const originalInput = cleanText(input.originalInput ?? existing?.original_input ?? "");
  if (!originalInput) {
    const error = new Error("Original input is required.");
    error.status = 400;
    error.code = "ORIGINAL_INPUT_REQUIRED";
    throw error;
  }

  const clarification = cleanText(input.clarification ?? existing?.clarification ?? "");
  const problemType =
    normalizeProblemType(input.problemType) ||
    normalizeProblemType(existing?.problem_type) ||
    detectProblemType(originalInput, clarification);
  const title = clampText(input.title || generateTitle(originalInput), 80) || "Untitled problem";
  const nextStep = clampText(input.nextStep || suggestNextStep(problemType), 160) || suggestNextStep(problemType);
  const priority = normalizePriority(input.priority) || normalizePriority(existing?.priority) || "medium";
  const state = normalizeState(input.state) || normalizeState(existing?.state) || "new";
  const ownerName = clampText(input.ownerName ?? existing?.owner_name ?? "", 120);
  const reframes = parseReframes(input.reframes, originalInput, clarification, problemType);
  const nowIso = new Date().toISOString();

  return {
    owner_email: normalizeEmail(ownerEmail),
    owner_name: ownerName || null,
    title,
    original_input: originalInput,
    problem_type: problemType,
    clarification: clarification || null,
    reframes,
    next_step: nextStep,
    priority,
    state,
    updated_at: nowIso,
  };
}

async function listProblemsForUser(email) {
  const rows = await listProblemRowsByOwnerEmail(email);
  const problems = sortProblems(rows.map(mapProblemRow));

  return {
    problems,
    meta: {
      totalProblems: problems.length,
      currentUserEmail: normalizeEmail(email),
    },
  };
}

async function createProblemForUser(email, body) {
  const payload = buildProblemPayload(body || {}, email);
  payload.created_at = new Date().toISOString();

  const row = await createProblemRow(payload);
  return {
    ok: true,
    problem: mapProblemRow(row),
  };
}

async function updateProblemForUser(email, body) {
  const problemId = ensureUuid(body?.id, "Problem");
  const existing = await getProblemRowByIdAndOwner(problemId, email);

  if (!existing) {
    const error = new Error("Problem not found.");
    error.status = 404;
    error.code = "PROBLEM_NOT_FOUND";
    throw error;
  }

  const patch = buildProblemPayload(body || {}, email, existing);
  const row = await updateProblemRow(problemId, email, patch);
  if (!row) {
    const error = new Error("Problem not found.");
    error.status = 404;
    error.code = "PROBLEM_NOT_FOUND";
    throw error;
  }

  return {
    ok: true,
    problem: mapProblemRow(row),
  };
}

function mapProblemError(error) {
  const status = Number(error?.status) >= 400 ? Number(error.status) : 400;
  const code = String(error?.code || "PROBLEMS_REQUEST_FAILED").trim().toUpperCase();

  return {
    status,
    payload: {
      error: {
        code,
        message: error?.message || "Problems request failed.",
      },
      detail: error?.detail || error?.message || "Unknown error.",
    },
  };
}

module.exports = {
  createProblemForUser,
  detectProblemType,
  generateReframes,
  generateTitle,
  listProblemsForUser,
  mapProblemError,
  suggestNextStep,
  updateProblemForUser,
};
