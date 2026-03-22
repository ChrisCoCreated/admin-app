import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260317";

const problemComposerForm = document.getElementById("problemComposerForm");
const problemInput = document.getElementById("problemInput");
const captureProblemBtn = document.getElementById("captureProblemBtn");
const voiceInputBtn = document.getElementById("voiceInputBtn");
const pageStatus = document.getElementById("pageStatus");
const problemsList = document.getElementById("problemsList");
const problemsEmptyState = document.getElementById("problemsEmptyState");
const signOutBtn = document.getElementById("signOutBtn");

const TYPE_LABELS = {
  concern: "Concern",
  opportunity: "Opportunity",
  decision: "Decision",
};

const PRIORITY_LABELS = {
  low: "Low",
  medium: "Medium",
  high: "High",
};

const STATE_LABELS = {
  new: "New",
  in_progress: "In Progress",
  done: "Done",
  parked: "Parked",
};

const TYPE_OPTIONS = ["concern", "opportunity", "decision"];
const PRIORITY_OPTIONS = ["low", "medium", "high"];
const STATE_OPTIONS = ["new", "in_progress", "done", "parked"];

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let currentUser = null;
let problems = [];
let creating = false;
const expandedIds = new Set();
const saveTimers = new Map();
const cardViews = new Map();
const saveSequenceById = new Map();

let recognition = null;
let voiceListening = false;
let voiceSeedText = "";

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
  return TYPE_OPTIONS.includes(text) ? text : "concern";
}

function normalizePriority(value) {
  const text = cleanText(value).toLowerCase();
  return PRIORITY_OPTIONS.includes(text) ? text : "medium";
}

function normalizeState(value) {
  const text = cleanText(value).toLowerCase();
  return STATE_OPTIONS.includes(text) ? text : "new";
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

function sortProblems(list) {
  return [...list].sort((a, b) => {
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
  const type = normalizeProblemType(problemType || detectProblemType(originalInput, clarification));
  const focus = deriveFocusPhrase(originalInput, clarification).replace(/[?!.]+$/g, "");

  if (type === "opportunity") {
    return uniqueStrings([
      `How might we define ${focus} more clearly?`,
      `How might we improve ${focus} in a way people notice quickly?`,
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
  const type = normalizeProblemType(problemType);
  if (type === "opportunity") {
    return "Define the opportunity clearly";
  }
  if (type === "decision") {
    return "List and compare options";
  }
  return "Gather examples / understand root cause";
}

function getClarificationPrompt(problemType) {
  if (problemType === "opportunity") {
    return "What could be better?";
  }
  if (problemType === "decision") {
    return "What choice needs to be made?";
  }
  return "What’s actually happening?";
}

function setPageStatus(message, isError = false) {
  pageStatus.textContent = message;
  pageStatus.classList.toggle("error", isError);
}

function setCreating(value) {
  creating = value === true;
  captureProblemBtn.disabled = creating;
  if (problemInput) {
    problemInput.disabled = creating;
  }
  if (voiceInputBtn && !voiceInputBtn.hidden) {
    voiceInputBtn.disabled = creating;
  }
}

function summarizeWhenUpdated(problem) {
  const updated = new Date(problem?.updatedAt || problem?.createdAt || "");
  if (Number.isNaN(updated.getTime())) {
    return "Autosaves as you edit";
  }
  return `Updated ${updated.toLocaleString()}`;
}

function buttonLabelGroup(values, activeValue, labelMap) {
  const fragment = document.createDocumentFragment();

  for (const value of values) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "problem-option";
    button.dataset.value = value;
    button.dataset.active = value === activeValue ? "true" : "false";
    button.textContent = labelMap[value] || value;
    fragment.appendChild(button);
  }

  return fragment;
}

function setOptionGroupState(group, activeValue) {
  const buttons = group?.querySelectorAll?.(".problem-option") || [];
  for (const button of buttons) {
    const isActive = button.dataset.value === activeValue;
    button.dataset.active = isActive ? "true" : "false";
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-pressed", isActive ? "true" : "false");
  }
}

function getProblemById(problemId) {
  return problems.find((entry) => entry.id === problemId) || null;
}

function upsertLocalProblem(problem) {
  const index = problems.findIndex((entry) => entry.id === problem.id);
  if (index >= 0) {
    problems[index] = { ...problems[index], ...problem };
    return problems[index];
  }
  problems.push(problem);
  return problem;
}

function updateVoiceButton() {
  if (!voiceInputBtn || voiceInputBtn.hidden) {
    return;
  }
  voiceInputBtn.classList.toggle("active", voiceListening);
  voiceInputBtn.textContent = voiceListening ? "Listening…" : "Voice";
}

function setCardSaveState(problemId, message, isError = false) {
  const view = cardViews.get(problemId);
  if (!view?.saveState) {
    return;
  }
  view.saveState.textContent = message;
  view.saveState.classList.toggle("error", isError);
}

function syncProblemSummary(problem, view) {
  if (!problem || !view) {
    return;
  }

  view.title.textContent = problem.title || "Untitled problem";
  view.typeBadge.textContent = TYPE_LABELS[problem.problemType] || "Concern";
  view.typeBadge.dataset.type = problem.problemType;
  view.priorityBadge.textContent = PRIORITY_LABELS[problem.priority] || "Medium";
  view.priorityBadge.dataset.priority = problem.priority;
  view.stateBadge.textContent = STATE_LABELS[problem.state] || "New";
  view.nextStepSummary.textContent = problem.nextStep || "No next step yet";
  view.suggestionNote.textContent = `Suggested: ${TYPE_LABELS[detectProblemType(problem.originalInput, problem.clarification)] || "Concern"}`;
  view.clarificationLabel.textContent = getClarificationPrompt(problem.problemType);
  view.clarificationInput.placeholder = getClarificationPrompt(problem.problemType);
  view.ownerInput.placeholder = currentUser?.email || "Current user";
  setOptionGroupState(view.typeGroup, problem.problemType);
  setOptionGroupState(view.priorityGroup, problem.priority);
  view.stateSelect.value = problem.state;
}

function applyReframesToInputs(problem, view) {
  const reframes = Array.isArray(problem.reframes) ? problem.reframes.slice(0, 3) : [];
  while (reframes.length < 3) {
    reframes.push("");
  }

  view.reframeInputs.forEach((input, index) => {
    input.value = reframes[index] || "";
  });
}

function queueProblemSave(problemId, options = {}) {
  const delay = Number.isFinite(options.delay) ? options.delay : 450;
  const existing = saveTimers.get(problemId);
  if (existing) {
    window.clearTimeout(existing);
  }

  setCardSaveState(problemId, "Saving soon…");
  const timer = window.setTimeout(() => {
    saveTimers.delete(problemId);
    void persistProblem(problemId, options);
  }, delay);
  saveTimers.set(problemId, timer);
}

async function persistProblem(problemId, options = {}) {
  const problem = getProblemById(problemId);
  if (!problem) {
    return;
  }

  const previousSequence = saveSequenceById.get(problemId) || 0;
  const nextSequence = previousSequence + 1;
  saveSequenceById.set(problemId, nextSequence);
  setCardSaveState(problemId, "Saving…");

  try {
    const payload = await directoryApi.updateProblemToSolve({
      id: problem.id,
      title: problem.title,
      originalInput: problem.originalInput,
      problemType: problem.problemType,
      clarification: problem.clarification,
      reframes: problem.reframes,
      nextStep: problem.nextStep,
      priority: problem.priority,
      state: problem.state,
      ownerName: problem.ownerName,
    });

    if ((saveSequenceById.get(problemId) || 0) !== nextSequence) {
      return;
    }

    const saved = payload?.problem ? upsertLocalProblem(payload.problem) : problem;
    setCardSaveState(problemId, summarizeWhenUpdated(saved));

    if (options.resort === true) {
      renderProblems();
      return;
    }

    const view = cardViews.get(problemId);
    if (view) {
      syncProblemSummary(saved, view);
    }
  } catch (error) {
    console.error("[problems] Save failed", error);
    setCardSaveState(problemId, error?.message || "Could not save", true);
  }
}

function buildProblemCard(problem) {
  const article = document.createElement("article");
  article.className = "problem-card";
  article.dataset.problemId = problem.id;

  const head = document.createElement("div");
  head.className = "problem-card-head";

  const identity = document.createElement("div");
  identity.className = "problem-card-identity";

  const title = document.createElement("h2");
  title.className = "problem-card-title";

  const badges = document.createElement("div");
  badges.className = "problem-card-badges";

  const typeBadge = document.createElement("span");
  typeBadge.className = "problem-badge";

  const priorityBadge = document.createElement("span");
  priorityBadge.className = "problem-badge";

  const stateBadge = document.createElement("span");
  stateBadge.className = "problem-badge problem-badge-muted";

  badges.appendChild(typeBadge);
  badges.appendChild(priorityBadge);
  badges.appendChild(stateBadge);

  const nextStepSummary = document.createElement("p");
  nextStepSummary.className = "problem-card-summary";

  identity.appendChild(title);
  identity.appendChild(badges);
  identity.appendChild(nextStepSummary);

  const headActions = document.createElement("div");
  headActions.className = "problem-card-actions";

  const saveState = document.createElement("p");
  saveState.className = "problem-save-state muted";
  saveState.textContent = summarizeWhenUpdated(problem);

  const toggleBtn = document.createElement("button");
  toggleBtn.type = "button";
  toggleBtn.className = "ghost subtle";
  toggleBtn.textContent = expandedIds.has(problem.id) ? "Collapse" : "Expand";

  headActions.appendChild(saveState);
  headActions.appendChild(toggleBtn);

  head.appendChild(identity);
  head.appendChild(headActions);

  const body = document.createElement("div");
  body.className = "problem-card-body";
  body.hidden = !expandedIds.has(problem.id);

  const originalSection = document.createElement("section");
  originalSection.className = "problem-section";
  const originalLabel = document.createElement("label");
  originalLabel.className = "field";
  originalLabel.textContent = "Original input";
  const originalInput = document.createElement("textarea");
  originalInput.value = problem.originalInput || "";
  originalInput.placeholder = "Capture the messy version here";
  originalLabel.appendChild(originalInput);
  originalSection.appendChild(originalLabel);

  const typeSection = document.createElement("section");
  typeSection.className = "problem-section";
  const typeLabel = document.createElement("p");
  typeLabel.className = "problem-section-title";
  typeLabel.textContent = "Type";
  const suggestionNote = document.createElement("p");
  suggestionNote.className = "problem-inline-note muted";
  const typeGroup = document.createElement("div");
  typeGroup.className = "problem-option-group";
  typeGroup.appendChild(buttonLabelGroup(TYPE_OPTIONS, problem.problemType, TYPE_LABELS));
  typeSection.appendChild(typeLabel);
  typeSection.appendChild(suggestionNote);
  typeSection.appendChild(typeGroup);

  const clarificationSection = document.createElement("section");
  clarificationSection.className = "problem-section";
  const clarificationLabelWrap = document.createElement("label");
  clarificationLabelWrap.className = "field";
  const clarificationLabel = document.createElement("span");
  clarificationLabel.className = "problem-section-title";
  const clarificationInput = document.createElement("input");
  clarificationInput.type = "text";
  clarificationInput.value = problem.clarification || "";
  clarificationLabelWrap.appendChild(clarificationLabel);
  clarificationLabelWrap.appendChild(clarificationInput);
  clarificationSection.appendChild(clarificationLabelWrap);

  const reframeSection = document.createElement("section");
  reframeSection.className = "problem-section";
  const reframeHead = document.createElement("div");
  reframeHead.className = "problem-section-head";
  const reframeTitle = document.createElement("p");
  reframeTitle.className = "problem-section-title";
  reframeTitle.textContent = "Reframe";
  const regenerateReframesBtn = document.createElement("button");
  regenerateReframesBtn.type = "button";
  regenerateReframesBtn.className = "ghost subtle";
  regenerateReframesBtn.textContent = "Regenerate";
  reframeHead.appendChild(reframeTitle);
  reframeHead.appendChild(regenerateReframesBtn);
  const reframeHint = document.createElement("p");
  reframeHint.className = "problem-inline-note muted";
  reframeHint.textContent = "Edit these or leave them blank if they are not helpful.";
  const reframeList = document.createElement("div");
  reframeList.className = "problem-reframe-list";
  const reframeInputs = Array.from({ length: 3 }, () => {
    const input = document.createElement("input");
    input.type = "text";
    input.placeholder = "How might we…";
    reframeList.appendChild(input);
    return input;
  });
  reframeSection.appendChild(reframeHead);
  reframeSection.appendChild(reframeHint);
  reframeSection.appendChild(reframeList);

  const nextStepSection = document.createElement("section");
  nextStepSection.className = "problem-section";
  const nextStepLabel = document.createElement("label");
  nextStepLabel.className = "field";
  nextStepLabel.textContent = "Next step";
  const nextStepInput = document.createElement("input");
  nextStepInput.type = "text";
  nextStepInput.value = problem.nextStep || "";
  nextStepInput.placeholder = "One concrete next step";
  nextStepLabel.appendChild(nextStepInput);
  nextStepSection.appendChild(nextStepLabel);

  const metaGrid = document.createElement("div");
  metaGrid.className = "problem-meta-grid";

  const prioritySection = document.createElement("section");
  prioritySection.className = "problem-section";
  const priorityTitle = document.createElement("p");
  priorityTitle.className = "problem-section-title";
  priorityTitle.textContent = "Priority";
  const priorityGroup = document.createElement("div");
  priorityGroup.className = "problem-option-group";
  priorityGroup.appendChild(buttonLabelGroup(PRIORITY_OPTIONS, problem.priority, PRIORITY_LABELS));
  prioritySection.appendChild(priorityTitle);
  prioritySection.appendChild(priorityGroup);

  const stateSection = document.createElement("section");
  stateSection.className = "problem-section";
  const stateLabel = document.createElement("label");
  stateLabel.className = "field";
  stateLabel.textContent = "State";
  const stateSelect = document.createElement("select");
  stateSelect.className = "problem-state-select";
  for (const state of STATE_OPTIONS) {
    const option = document.createElement("option");
    option.value = state;
    option.textContent = STATE_LABELS[state];
    stateSelect.appendChild(option);
  }
  stateLabel.appendChild(stateSelect);
  stateSection.appendChild(stateLabel);

  const ownerSection = document.createElement("section");
  ownerSection.className = "problem-section";
  const ownerLabel = document.createElement("label");
  ownerLabel.className = "field";
  ownerLabel.textContent = "Owner (optional)";
  const ownerInput = document.createElement("input");
  ownerInput.type = "text";
  ownerInput.value = problem.ownerName || "";
  ownerInput.placeholder = currentUser?.email || "Current user";
  ownerLabel.appendChild(ownerInput);
  ownerSection.appendChild(ownerLabel);

  metaGrid.appendChild(prioritySection);
  metaGrid.appendChild(stateSection);
  metaGrid.appendChild(ownerSection);

  body.appendChild(originalSection);
  body.appendChild(typeSection);
  body.appendChild(clarificationSection);
  body.appendChild(reframeSection);
  body.appendChild(nextStepSection);
  body.appendChild(metaGrid);

  article.appendChild(head);
  article.appendChild(body);

  const view = {
    article,
    body,
    title,
    typeBadge,
    priorityBadge,
    stateBadge,
    nextStepSummary,
    saveState,
    suggestionNote,
    clarificationLabel,
    clarificationInput,
    typeGroup,
    priorityGroup,
    stateSelect,
    ownerInput,
    reframeInputs,
  };

  syncProblemSummary(problem, view);
  applyReframesToInputs(problem, view);

  toggleBtn.addEventListener("click", () => {
    const expanded = expandedIds.has(problem.id);
    if (expanded) {
      expandedIds.delete(problem.id);
    } else {
      expandedIds.add(problem.id);
    }
    body.hidden = expanded;
    toggleBtn.textContent = expanded ? "Expand" : "Collapse";
  });

  originalInput.addEventListener("input", () => {
    const current = getProblemById(problem.id);
    if (!current) {
      return;
    }

    const previousSuggestedTitle = generateTitle(current.originalInput);
    current.originalInput = originalInput.value;
    const nextSuggestedTitle = generateTitle(current.originalInput);
    if (!cleanText(current.title) || cleanText(current.title) === cleanText(previousSuggestedTitle)) {
      current.title = nextSuggestedTitle;
    }
    syncProblemSummary(current, view);
    queueProblemSave(current.id);
  });

  typeGroup.addEventListener("click", (event) => {
    const button = event.target.closest(".problem-option");
    if (!button) {
      return;
    }

    const current = getProblemById(problem.id);
    if (!current) {
      return;
    }

    const previousType = current.problemType;
    const previousReframes = generateReframes(current.originalInput, current.clarification, previousType);
    const currentReframes = current.reframes || [];
    current.problemType = button.dataset.value || "concern";

    const existingNextStep = cleanText(current.nextStep);
    if (!existingNextStep || existingNextStep === suggestNextStep(previousType)) {
      current.nextStep = suggestNextStep(current.problemType);
      nextStepInput.value = current.nextStep;
    }

    const currentReframesKey = currentReframes.map((entry) => cleanText(entry)).filter(Boolean).join("|").toLowerCase();
    const previousReframesKey = previousReframes.map((entry) => cleanText(entry)).filter(Boolean).join("|").toLowerCase();
    if (!currentReframesKey || currentReframesKey === previousReframesKey) {
      current.reframes = generateReframes(current.originalInput, current.clarification, current.problemType);
      applyReframesToInputs(current, view);
    }

    syncProblemSummary(current, view);
    queueProblemSave(current.id, { delay: 120 });
  });

  clarificationInput.addEventListener("input", () => {
    const current = getProblemById(problem.id);
    if (!current) {
      return;
    }

    current.clarification = clarificationInput.value;
    syncProblemSummary(current, view);
    queueProblemSave(current.id);
  });

  regenerateReframesBtn.addEventListener("click", () => {
    const current = getProblemById(problem.id);
    if (!current) {
      return;
    }

    current.reframes = generateReframes(current.originalInput, current.clarification, current.problemType);
    applyReframesToInputs(current, view);
    queueProblemSave(current.id, { delay: 120 });
  });

  reframeInputs.forEach((input, index) => {
    input.addEventListener("input", () => {
      const current = getProblemById(problem.id);
      if (!current) {
        return;
      }

      const nextReframes = [...(current.reframes || [])];
      nextReframes[index] = input.value;
      current.reframes = nextReframes.map((entry) => clampText(entry, 160)).slice(0, 3);
      queueProblemSave(current.id);
    });
  });

  nextStepInput.addEventListener("input", () => {
    const current = getProblemById(problem.id);
    if (!current) {
      return;
    }

    current.nextStep = nextStepInput.value;
    syncProblemSummary(current, view);
    queueProblemSave(current.id);
  });

  priorityGroup.addEventListener("click", (event) => {
    const button = event.target.closest(".problem-option");
    if (!button) {
      return;
    }

    const current = getProblemById(problem.id);
    if (!current) {
      return;
    }

    current.priority = button.dataset.value || "medium";
    syncProblemSummary(current, view);
    queueProblemSave(current.id, { delay: 120, resort: true });
  });

  stateSelect.addEventListener("change", () => {
    const current = getProblemById(problem.id);
    if (!current) {
      return;
    }

    current.state = stateSelect.value;
    syncProblemSummary(current, view);
    queueProblemSave(current.id, { delay: 120, resort: true });
  });

  ownerInput.addEventListener("input", () => {
    const current = getProblemById(problem.id);
    if (!current) {
      return;
    }

    current.ownerName = ownerInput.value;
    queueProblemSave(current.id);
  });

  cardViews.set(problem.id, view);
  return article;
}

function renderProblems() {
  problemsList.innerHTML = "";
  cardViews.clear();

  const sorted = sortProblems(problems);
  problems = sorted;

  if (!sorted.length) {
    problemsEmptyState.hidden = false;
    return;
  }

  problemsEmptyState.hidden = true;
  for (const problem of sorted) {
    problemsList.appendChild(buildProblemCard(problem));
  }
}

function normalizeProblem(problem) {
  return {
    id: problem?.id || "",
    ownerEmail: cleanText(problem?.ownerEmail || currentUser?.email || "").toLowerCase(),
    ownerName: cleanText(problem?.ownerName || ""),
    title: cleanText(problem?.title || generateTitle(problem?.originalInput || "")),
    originalInput: cleanText(problem?.originalInput || ""),
    problemType: normalizeProblemType(problem?.problemType || detectProblemType(problem?.originalInput || "", problem?.clarification || "")),
    clarification: cleanText(problem?.clarification || ""),
    reframes: Array.isArray(problem?.reframes)
      ? problem.reframes.map((entry) => clampText(entry, 160)).filter(Boolean).slice(0, 3)
      : [],
    nextStep: cleanText(problem?.nextStep || suggestNextStep(problem?.problemType || "concern")),
    priority: normalizePriority(problem?.priority || "medium"),
    state: normalizeState(problem?.state || "new"),
    createdAt: problem?.createdAt || null,
    updatedAt: problem?.updatedAt || null,
  };
}

async function refreshProblems() {
  setPageStatus("Loading problems...");
  try {
    const payload = await directoryApi.listProblemsToSolve();
    const nextProblems = Array.isArray(payload?.problems) ? payload.problems.map(normalizeProblem) : [];
    problems = sortProblems(nextProblems);
    renderProblems();
    setPageStatus(`${problems.length} problem${problems.length === 1 ? "" : "s"} ready.`);
  } catch (error) {
    console.error("[problems] Refresh failed", error);
    setPageStatus(error?.message || "Could not load problems.", true);
    problems = [];
    renderProblems();
  }
}

async function createProblemFromInput() {
  const originalInput = cleanText(problemInput?.value || "");
  if (!originalInput || creating) {
    return;
  }

  setCreating(true);
  setPageStatus("Capturing problem...");

  try {
    const payload = await directoryApi.createProblemToSolve({
      originalInput,
      ownerName: currentUser?.email || "",
    });

    const problem = normalizeProblem(payload?.problem || {});
    upsertLocalProblem(problem);
    expandedIds.add(problem.id);
    renderProblems();
    problemInput.value = "";
    problemInput.focus();
    setPageStatus("Problem captured. Fill in as much or as little as you need.");
  } catch (error) {
    console.error("[problems] Create failed", error);
    setPageStatus(error?.message || "Could not capture problem.", true);
  } finally {
    setCreating(false);
  }
}

function setupVoiceInput() {
  const Recognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!Recognition || !voiceInputBtn) {
    return;
  }

  recognition = new Recognition();
  recognition.lang = "en-GB";
  recognition.continuous = false;
  recognition.interimResults = true;
  voiceInputBtn.hidden = false;
  updateVoiceButton();

  recognition.addEventListener("start", () => {
    voiceListening = true;
    updateVoiceButton();
  });

  recognition.addEventListener("end", () => {
    voiceListening = false;
    updateVoiceButton();
  });

  recognition.addEventListener("result", (event) => {
    const transcript = Array.from(event.results || [])
      .map((result) => result?.[0]?.transcript || "")
      .join(" ")
      .trim();

    const merged = [voiceSeedText, transcript].filter(Boolean).join(" ").trim();
    if (problemInput) {
      problemInput.value = merged;
    }
  });

  recognition.addEventListener("error", () => {
    voiceListening = false;
    updateVoiceButton();
  });

  voiceInputBtn.addEventListener("click", () => {
    if (!recognition || creating) {
      return;
    }

    if (voiceListening) {
      recognition.stop();
      return;
    }

    voiceSeedText = cleanText(problemInput?.value || "");
    recognition.start();
  });
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();
    if (!canAccessPage(role, "problems")) {
      window.location.href = "./unauthorized.html?page=problems";
      return;
    }

    currentUser = profile;
    renderTopNavigation({ role, currentPathname: window.location.pathname });
    setupVoiceInput();
    await refreshProblems();
  } catch (error) {
    console.error("[problems] Init failed", error);
    setPageStatus(error?.message || "Could not initialize page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

problemComposerForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  await createProblemFromInput();
});

problemInput?.addEventListener("keydown", async (event) => {
  if (event.key !== "Enter" || event.shiftKey) {
    return;
  }
  event.preventDefault();
  await createProblemFromInput();
});

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

void init();
