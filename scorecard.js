import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260317";

const ENTITY_META = {
  values: {
    label: "value",
    nameField: "name",
    reviewLabelField: "nameSnapshot",
    listKey: "values",
  },
  principles: {
    label: "principle",
    nameField: "name",
    reviewLabelField: "nameSnapshot",
    listKey: "principles",
  },
  objectives: {
    label: "objective",
    nameField: "title",
    reviewLabelField: "titleSnapshot",
    listKey: "objectives",
  },
  goals: {
    label: "goal",
    nameField: "title",
    reviewLabelField: "titleSnapshot",
    listKey: "goals",
  },
};

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const saveMessage = document.getElementById("saveMessage");
const saveNowBtn = document.getElementById("saveNowBtn");
const reviewPeriodInput = document.getElementById("reviewPeriodInput");
const loadPeriodBtn = document.getElementById("loadPeriodBtn");
const valuesList = document.getElementById("valuesList");
const principlesList = document.getElementById("principlesList");
const objectivesList = document.getElementById("objectivesList");
const goalsList = document.getElementById("goalsList");
const snapshotGrid = document.getElementById("snapshotGrid");
const resetBtn = document.getElementById("resetBtn");
const reflectionGoingWell = document.getElementById("reflectionGoingWell");
const reflectionNeedsAttention = document.getElementById("reflectionNeedsAttention");
const reflectionNextAction = document.getElementById("reflectionNextAction");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let state = createEmptyState();
let activeReviewPeriod = currentReviewPeriod();
let saveTimer = null;
let loadingPeriod = false;
let savingReview = false;
const pendingDefinitionTimers = new Map();
const pendingDefinitionPatches = new Map();

function currentReviewPeriod() {
  return new Intl.DateTimeFormat(undefined, {
    month: "long",
    year: "numeric",
  }).format(new Date());
}

function normalizeReviewPeriod(value) {
  const reviewPeriod = String(value || "").trim();
  return reviewPeriod || currentReviewPeriod();
}

function createEmptyState(reviewPeriod = currentReviewPeriod()) {
  return {
    review: {
      id: null,
      reviewPeriod,
      reviewType: null,
      createdAt: null,
      updatedAt: null,
      createdBy: null,
      updatedBy: null,
    },
    reflection: {
      goingWell: "",
      needsAttention: "",
      nextAction: "",
    },
    values: [],
    principles: [],
    objectives: [],
    goals: [],
  };
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setSaveMessage(message, isError = false, tone = isError ? "error" : "idle") {
  saveMessage.textContent = message;
  saveMessage.classList.toggle("error", isError);
  saveMessage.dataset.tone = tone;
}

function setPendingSaveMessage(message) {
  setSaveMessage(message, false, "pending");
}

function setSavingMessage(message) {
  setSaveMessage(message, false, "saving");
}

function hasPendingDefinitionSaves() {
  return pendingDefinitionTimers.size > 0;
}

function hasPendingReviewSave() {
  return saveTimer != null;
}

function setBusyState(isBusy) {
  loadingPeriod = isBusy;
  loadPeriodBtn.disabled = isBusy;
  resetBtn.disabled = isBusy || savingReview;
  if (saveNowBtn) {
    saveNowBtn.disabled = isBusy || savingReview;
  }
}

function setSavingReviewState(isSaving) {
  savingReview = isSaving;
  resetBtn.disabled = isSaving || loadingPeriod;
  if (saveNowBtn) {
    saveNowBtn.disabled = isSaving || loadingPeriod;
  }
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function cloneReview(review) {
  return review && typeof review === "object" ? { ...review } : null;
}

function hydrateState(payload, reviewPeriod) {
  const next = payload && typeof payload === "object" ? payload : createEmptyState(reviewPeriod);
  state = {
    review: {
      id: next.review?.id || null,
      reviewPeriod: next.review?.reviewPeriod || reviewPeriod,
      reviewType: next.review?.reviewType || null,
      createdAt: next.review?.createdAt || null,
      updatedAt: next.review?.updatedAt || null,
      createdBy: next.review?.createdBy || null,
      updatedBy: next.review?.updatedBy || null,
    },
    reflection: {
      goingWell: String(next.reflection?.goingWell || ""),
      needsAttention: String(next.reflection?.needsAttention || ""),
      nextAction: String(next.reflection?.nextAction || ""),
    },
    values: Array.isArray(next.values)
      ? next.values.map((item) => ({
          ...item,
          review: cloneReview(item.review),
        }))
      : [],
    principles: Array.isArray(next.principles)
      ? next.principles.map((item) => ({
          ...item,
          review: cloneReview(item.review),
        }))
      : [],
    objectives: Array.isArray(next.objectives)
      ? next.objectives.map((item) => ({
          ...item,
          review: cloneReview(item.review),
        }))
      : [],
    goals: Array.isArray(next.goals)
      ? next.goals.map((item) => ({
          ...item,
          review: cloneReview(item.review),
        }))
      : [],
  };
}

function formatSaveMeta(review) {
  const updatedAt = review?.updatedAt ? new Date(review.updatedAt) : null;
  if (!updatedAt || Number.isNaN(updatedAt.getTime())) {
    return "Review changes save automatically to the shared scorecard.";
  }
  const userLabel = review?.updatedBy ? ` by ${review.updatedBy}` : "";
  return `Saved to the shared scorecard${userLabel} at ${updatedAt.toLocaleString()}.`;
}

function buildScoreOptions(selected) {
  const options = ['<option value="">Select score</option>'];
  for (let score = 1; score <= 5; score += 1) {
    options.push(`<option value="${score}"${String(selected) === String(score) ? " selected" : ""}>${score}</option>`);
  }
  return options.join("");
}

function ensureReviewObject(item) {
  if (!item.review) {
    item.review = {};
  }
  return item.review;
}

function itemDisplayName(item, entityType) {
  const labelField = ENTITY_META[entityType].nameField;
  return String(item?.[labelField] || "").trim();
}

function scoreTone(value) {
  if (value == null) {
    return "is-empty";
  }
  if (value >= 4) {
    return "is-strong";
  }
  if (value >= 3) {
    return "is-watch";
  }
  return "is-risk";
}

function formatAverage(value) {
  return value == null ? "Not scored" : `${value.toFixed(1)} / 5`;
}

function averageScore(items, reviewField) {
  const scores = items
    .map((item) => Number(item?.review?.[reviewField]))
    .filter((value) => Number.isFinite(value) && value >= 1 && value <= 5);

  if (!scores.length) {
    return null;
  }

  return scores.reduce((sum, value) => sum + value, 0) / scores.length;
}

function countScored(items, reviewField) {
  return items.filter((item) => {
    const score = Number(item?.review?.[reviewField]);
    return Number.isFinite(score) && score >= 1 && score <= 5;
  }).length;
}

function getObjectiveOptions() {
  return state.objectives.map((item) => ({
    id: item.id,
    title: item.title || "",
    isActive: item.isActive !== false,
  }));
}

function buildObjectiveSelect(selectedId) {
  const options = ['<option value="">Choose linked objective</option>'];
  for (const objective of getObjectiveOptions()) {
    options.push(
      `<option value="${escapeHtml(objective.id)}"${objective.id === selectedId ? " selected" : ""}>${escapeHtml(
        objective.title || "(Untitled objective)"
      )}</option>`
    );
  }
  return options.join("");
}

function sortEntityItems(entityType, items) {
  const labelField = ENTITY_META[entityType].nameField;
  return [...items].sort((left, right) => {
    const orderDelta = Number(left.sortOrder || 0) - Number(right.sortOrder || 0);
    if (orderDelta !== 0) {
      return orderDelta;
    }
    return String(left?.[labelField] || "").localeCompare(String(right?.[labelField] || ""), undefined, {
      sensitivity: "base",
    });
  });
}

function appendDefinitionToState(entityType, definition) {
  if (!definition?.id) {
    return;
  }

  const listKey = ENTITY_META[entityType].listKey;
  const existing = state[listKey].filter((item) => item.id !== definition.id);
  const hydrated =
    entityType === "goals"
      ? {
          ...definition,
          objectiveTitle: state.objectives.find((item) => item.id === definition.objectiveId)?.title || "",
          review: null,
        }
      : {
          ...definition,
          review: null,
        };

  state[listKey] = sortEntityItems(entityType, [...existing, hydrated]);
}

function definitionSnapshotNote(item, entityType) {
  const snapshotField = ENTITY_META[entityType].reviewLabelField;
  const currentLabel = itemDisplayName(item, entityType);
  const snapshot = String(item?.review?.[snapshotField] || "").trim();
  if (!snapshot || snapshot === currentLabel) {
    return "";
  }
  return `<p class="scorecard-definition-note muted">Reviewed as: ${escapeHtml(snapshot)}</p>`;
}

function goalObjectiveSnapshotNote(item) {
  const currentTitle = String(item?.objectiveTitle || "").trim();
  const snapshotTitle = String(item?.review?.objectiveTitleSnapshot || "").trim();
  if (!snapshotTitle || snapshotTitle === currentTitle) {
    return "";
  }
  return `<p class="scorecard-definition-note muted">Reviewed against objective: ${escapeHtml(snapshotTitle)}</p>`;
}

function definitionDetailNote(item) {
  const detail = String(item?.description || "").trim();
  if (!detail) {
    return "";
  }
  return `<p class="scorecard-definition-note muted">${escapeHtml(detail)}</p>`;
}

function buildEntryShell(item, title, description, content, entityType) {
  const archiveLabel = item.isActive === false ? "Archived" : "Archive";
  return `
    <article class="scorecard-entry ${item.isActive === false ? "is-archived" : ""}" data-row-id="${escapeHtml(item.id)}" data-entity="${entityType}">
      <div class="scorecard-entry-head">
        <div>
          <div class="scorecard-entry-title-row">
            <h3>${title}</h3>
            ${item.isActive === false ? '<span class="scorecard-archived-chip">Archived</span>' : ""}
          </div>
          <p class="muted">${description}</p>
          ${definitionDetailNote(item)}
          ${definitionSnapshotNote(item, entityType)}
          ${entityType === "goals" ? goalObjectiveSnapshotNote(item) : ""}
        </div>
        <button class="secondary scorecard-remove-btn" type="button" data-action="archive"${item.isActive === false ? " disabled" : ""}>
          ${archiveLabel}
        </button>
      </div>
      ${content}
    </article>
  `;
}

function renderValues() {
  if (!state.values.length) {
    valuesList.innerHTML = '<p class="muted scorecard-empty-state">No values defined yet. Add your organisational values to start scoring them each period.</p>';
    return;
  }

  valuesList.innerHTML = state.values
    .map((item, index) => {
      const review = ensureReviewObject(item);
      const showAction = Number(review.score) > 0 && Number(review.score) < 4;
      return buildEntryShell(
        item,
        escapeHtml(item.name || `Value ${index + 1}`),
        "Scores and commentary on this page are specific to the selected review period. Rename shared values from Scorecard Setup.",
        `
          <div class="scorecard-form-grid">
            <label class="field">
              Score (1-5)
              <select data-scope="review" data-field="score">${buildScoreOptions(review.score)}</select>
            </label>
            <label class="field scorecard-span-2">
              Evidence / examples
              <textarea data-scope="review" data-field="evidence" placeholder="What supports this score?">${escapeHtml(review.evidence || "")}</textarea>
            </label>
            <label class="field scorecard-span-2 ${showAction ? "" : "is-muted-block"}">
              Improvement action ${showAction ? "" : "(only needed below 4)"}
              <textarea data-scope="review" data-field="improvementAction" placeholder="What needs to change next?">${escapeHtml(
                review.improvementAction || ""
              )}</textarea>
            </label>
          </div>
        `,
        "values"
      );
    })
    .join("");
}

function renderPrinciples() {
  if (!state.principles.length) {
    principlesList.innerHTML =
      '<p class="muted scorecard-empty-state">No decision principles defined yet. Add the principles leadership wants to assess consistently.</p>';
    return;
  }

  principlesList.innerHTML = state.principles
    .map((item, index) => {
      const review = ensureReviewObject(item);
      return buildEntryShell(
        item,
        escapeHtml(item.name || `Principle ${index + 1}`),
        "Scores and examples below belong only to the selected review period. Edit shared principles from Scorecard Setup.",
        `
          <div class="scorecard-form-grid">
            <label class="field">
              Score (1-5)
              <select data-scope="review" data-field="score">${buildScoreOptions(review.score)}</select>
            </label>
            <label class="field scorecard-span-2">
              Example decision demonstrating the principle
              <textarea data-scope="review" data-field="exampleDecision" placeholder="Which recent decision shows this in practice?">${escapeHtml(
                review.exampleDecision || ""
              )}</textarea>
            </label>
            <label class="field scorecard-span-2">
              Adjustment needed
              <textarea data-scope="review" data-field="adjustmentNeeded" placeholder="What should we refine next period?">${escapeHtml(
                review.adjustmentNeeded || ""
              )}</textarea>
            </label>
          </div>
        `,
        "principles"
      );
    })
    .join("");
}

function renderObjectives() {
  if (!state.objectives.length) {
    objectivesList.innerHTML =
      '<p class="muted scorecard-empty-state">No strategic objectives defined yet. Add the core objectives for the year first.</p>';
    return;
  }

  objectivesList.innerHTML = state.objectives
    .map((item, index) => {
      const review = ensureReviewObject(item);
      return buildEntryShell(
        item,
        escapeHtml(item.title || `Objective ${index + 1}`),
        "Owner, score, and blockers below are review-specific. Edit shared objectives from Scorecard Setup.",
        `
          <div class="scorecard-form-grid">
            <label class="field">
              Owner
              <input data-scope="review" data-field="owner" type="text" value="${escapeHtml(review.owner || "")}" placeholder="Owner name" />
            </label>
            <label class="field">
              Progress score (1-5)
              <select data-scope="review" data-field="progressScore">${buildScoreOptions(review.progressScore)}</select>
            </label>
            <label class="field scorecard-span-2">
              Key risks or blockers
              <textarea data-scope="review" data-field="risks" placeholder="What is slowing progress or increasing risk?">${escapeHtml(
                review.risks || ""
              )}</textarea>
            </label>
          </div>
        `,
        "objectives"
      );
    })
    .join("");
}

function renderGoals() {
  if (!state.goals.length) {
    goalsList.innerHTML =
      '<p class="muted scorecard-empty-state">No active goals defined yet. Add the initiatives that deliver your objectives.</p>';
    return;
  }

  goalsList.innerHTML = state.goals
    .map((item, index) => {
      const review = ensureReviewObject(item);
      return buildEntryShell(
        item,
        escapeHtml(item.title || `Goal ${index + 1}`),
        "Status, score, owner, and next action are specific to the selected review period. Edit shared goals from Goal Setup.",
        `
          <div class="scorecard-form-grid">
            <label class="field scorecard-span-2 is-muted-block">
              Linked objective
              <input type="text" value="${escapeHtml(item.objectiveTitle || "")}" readonly />
            </label>
            <label class="field">
              Owner
              <input data-scope="review" data-field="owner" type="text" value="${escapeHtml(review.owner || "")}" placeholder="Owner name" />
            </label>
            <label class="field">
              Status
              <select data-scope="review" data-field="status">
                <option value=""${!review.status ? " selected" : ""}>Select status</option>
                <option value="On track"${review.status === "On track" ? " selected" : ""}>On track</option>
                <option value="At risk"${review.status === "At risk" ? " selected" : ""}>At risk</option>
                <option value="Off track"${review.status === "Off track" ? " selected" : ""}>Off track</option>
              </select>
            </label>
            <label class="field">
              Progress score (1-5)
              <select data-scope="review" data-field="progressScore">${buildScoreOptions(review.progressScore)}</select>
            </label>
            <label class="field scorecard-span-2">
              Next key action
              <textarea data-scope="review" data-field="nextAction" placeholder="What happens next to move this forward?">${escapeHtml(
                review.nextAction || ""
              )}</textarea>
            </label>
          </div>
        `,
        "goals"
      );
    })
    .join("");
}

function renderSnapshot() {
  const cards = [
    {
      label: "Values Score",
      value: averageScore(state.values, "score"),
      note: `${countScored(state.values, "score")} values scored`,
    },
    {
      label: "Decision Principles Score",
      value: averageScore(state.principles, "score"),
      note: `${countScored(state.principles, "score")} principles scored`,
    },
    {
      label: "Strategic Objectives Score",
      value: averageScore(state.objectives, "progressScore"),
      note: `${countScored(state.objectives, "progressScore")} objectives scored`,
    },
    {
      label: "Goals Delivery Score",
      value: averageScore(state.goals, "progressScore"),
      note: `${countScored(state.goals, "progressScore")} goals scored`,
    },
  ];

  snapshotGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="scorecard-snapshot-card ${scoreTone(card.value)}">
          <p class="scorecard-snapshot-label">${card.label}</p>
          <p class="scorecard-snapshot-value">${formatAverage(card.value)}</p>
          <p class="muted">${card.note}</p>
        </article>
      `
    )
    .join("");
}

function renderReflection() {
  reviewPeriodInput.value = state.review.reviewPeriod || activeReviewPeriod;
  reflectionGoingWell.value = state.reflection.goingWell;
  reflectionNeedsAttention.value = state.reflection.needsAttention;
  reflectionNextAction.value = state.reflection.nextAction;
}

function render() {
  renderSnapshot();
  renderValues();
  renderPrinciples();
  renderObjectives();
  renderGoals();
  renderReflection();
}

function serializeReviewPayload() {
  return {
    reflection: {
      goingWell: state.reflection.goingWell,
      needsAttention: state.reflection.needsAttention,
      nextAction: state.reflection.nextAction,
      reviewType: state.review.reviewType,
    },
    values: state.values.map((item) => ({
      id: item.id,
      score: item.review?.score ?? "",
      evidence: item.review?.evidence || "",
      improvementAction: item.review?.improvementAction || "",
    })),
    principles: state.principles.map((item) => ({
      id: item.id,
      score: item.review?.score ?? "",
      exampleDecision: item.review?.exampleDecision || "",
      adjustmentNeeded: item.review?.adjustmentNeeded || "",
    })),
    objectives: state.objectives.map((item) => ({
      id: item.id,
      owner: item.review?.owner || "",
      progressScore: item.review?.progressScore ?? "",
      risks: item.review?.risks || "",
    })),
    goals: state.goals.map((item) => ({
      id: item.id,
      owner: item.review?.owner || "",
      status: item.review?.status || "",
      progressScore: item.review?.progressScore ?? "",
      nextAction: item.review?.nextAction || "",
      objectiveTitleSnapshot: item.review?.objectiveTitleSnapshot || item.objectiveTitle || "",
    })),
  };
}

async function persistReview() {
  try {
    setSavingReviewState(true);
    setSavingMessage("Saving review...");
    const payload = await directoryApi.upsertPerformanceScorecard({
      reviewPeriod: activeReviewPeriod,
      scorecard: serializeReviewPayload(),
    });
    hydrateState(payload, activeReviewPeriod);
    activeReviewPeriod = state.review.reviewPeriod || activeReviewPeriod;
    render();
    setSaveMessage(formatSaveMeta(state.review));
  } catch (error) {
    console.error("[scorecard] Failed to save review", error);
    setSaveMessage(error?.message || "Could not save review changes.", true);
  } finally {
    setSavingReviewState(false);
  }
}

function scheduleReviewSave() {
  if (loadingPeriod) {
    return;
  }
  setPendingSaveMessage("Unsaved review changes. Saving shortly...");
  window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    saveTimer = null;
    void persistReview();
  }, 260);
}

async function flushPendingReviewSave() {
  if (!saveTimer) {
    return;
  }
  window.clearTimeout(saveTimer);
  saveTimer = null;
  await persistReview();
}

function queueDefinitionSave(entityType, rowId, patch) {
  const key = `${entityType}:${rowId}`;
  pendingDefinitionPatches.set(key, {
    ...(pendingDefinitionPatches.get(key) || {}),
    ...patch,
  });

  const existingTimer = pendingDefinitionTimers.get(key);
  if (existingTimer) {
    window.clearTimeout(existingTimer);
  }

  setPendingSaveMessage("Unsaved definition changes. Saving shortly...");
  const timer = window.setTimeout(async () => {
    pendingDefinitionTimers.delete(key);
    const nextPatch = pendingDefinitionPatches.get(key);
    pendingDefinitionPatches.delete(key);

    try {
      setSavingMessage("Saving shared definitions...");
      await directoryApi.updatePerformanceScorecardDefinition({
        entityType,
        id: rowId,
        patch: nextPatch,
      });
      setSaveMessage("Shared definitions saved.", false, "saved");
    } catch (error) {
      console.error("[scorecard] Failed to save definition", error);
      setSaveMessage(error?.message || "Could not save shared definition changes.", true);
      await loadScorecardForPeriod(activeReviewPeriod);
    }
  }, 260);

  pendingDefinitionTimers.set(key, timer);
}

async function flushDefinitionSaves() {
  const tasks = [];
  for (const [key, timer] of pendingDefinitionTimers.entries()) {
    window.clearTimeout(timer);
    pendingDefinitionTimers.delete(key);
    const [entityType, rowId] = key.split(":");
    const patch = pendingDefinitionPatches.get(key);
    pendingDefinitionPatches.delete(key);
    tasks.push(
      directoryApi.updatePerformanceScorecardDefinition({
        entityType,
        id: rowId,
        patch,
      })
    );
  }

  if (tasks.length) {
    setSavingMessage("Saving shared definitions...");
    await Promise.all(tasks);
    setSaveMessage("Shared definitions saved.", false, "saved");
  }
}

function findEntityRow(entityType, rowId) {
  const list = state[ENTITY_META[entityType].listKey];
  return list.find((item) => item.id === rowId) || null;
}

function updateDefinitionField(entityType, rowId, field, value) {
  const item = findEntityRow(entityType, rowId);
  if (!item) {
    return;
  }

  if (field === "objectiveId") {
    item.objectiveId = value || null;
    renderGoals();
    queueDefinitionSave(entityType, rowId, {
      objective_id: value || null,
    });
    return;
  }

  item[field] = value;

  queueDefinitionSave(entityType, rowId, {
    [field === "title" ? "title" : "name"]: value,
  });
}

function updateReviewField(entityType, rowId, field, value) {
  const item = findEntityRow(entityType, rowId);
  if (!item) {
    return;
  }

  ensureReviewObject(item)[field] = value;

  if (entityType === "values" && field === "score") {
    renderValues();
    renderSnapshot();
  } else if (entityType === "principles" && field === "score") {
    renderPrinciples();
    renderSnapshot();
  } else if (entityType === "objectives" && field === "progressScore") {
    renderObjectives();
    renderSnapshot();
  } else if (entityType === "goals" && (field === "progressScore" || field === "status")) {
    renderGoals();
    renderSnapshot();
  }

  scheduleReviewSave();
}

async function archiveDefinition(entityType, rowId) {
  try {
    await flushPendingReviewSave();
    await flushDefinitionSaves();
    setSaveMessage("Archiving shared definition...");
    await directoryApi.updatePerformanceScorecardDefinition({
      entityType,
      id: rowId,
      patch: {
        is_active: false,
      },
    });
    await loadScorecardForPeriod(activeReviewPeriod);
  } catch (error) {
    console.error("[scorecard] Failed to archive definition", error);
    setSaveMessage(error?.message || "Could not archive the definition.", true);
  }
}

async function addDefinition(entityType) {
  try {
    await flushDefinitionSaves();
    if (entityType === "goals" && !state.objectives.length) {
      setSaveMessage("Add an objective before creating a goal.", true);
      return;
    }

    setSaveMessage("Creating shared definition...");
    const payload =
      entityType === "values"
        ? { name: "" }
        : entityType === "principles"
        ? { name: "" }
        : entityType === "objectives"
        ? { title: "" }
        : {
            title: "",
            objective_id: state.objectives[0]?.id || null,
          };

    const response = await directoryApi.createPerformanceScorecardDefinition({
      entityType,
      definition: payload,
    });
    appendDefinitionToState(entityType, response?.definition || null);
    render();
    setSaveMessage(`New ${ENTITY_META[entityType].label} saved.`, false, "saved");
  } catch (error) {
    console.error("[scorecard] Failed to add definition", error);
    setSaveMessage(error?.message || "Could not create the definition.", true);
  }
}

async function loadScorecardForPeriod(reviewPeriod) {
  const nextPeriod = normalizeReviewPeriod(reviewPeriod);
  try {
    setBusyState(true);
    setSavingMessage(`Loading ${nextPeriod}...`);
    const payload = await directoryApi.getPerformanceScorecard({ reviewPeriod: nextPeriod });
    hydrateState(payload, nextPeriod);
    activeReviewPeriod = state.review.reviewPeriod || nextPeriod;
    render();
    if (state.review.updatedAt) {
      setSaveMessage(formatSaveMeta(state.review));
    } else {
      setSaveMessage(
        `Loaded ${activeReviewPeriod}. Scores and notes are empty until you assess this period.`,
        false,
        "idle"
      );
    }
  } catch (error) {
    console.error("[scorecard] Failed to load period", error);
    if (!snapshotGrid?.children?.length) {
      hydrateState(createEmptyState(nextPeriod), nextPeriod);
      activeReviewPeriod = nextPeriod;
      render();
    }
    setSaveMessage(error?.message || `Could not load ${nextPeriod}.`, true);
  } finally {
    setBusyState(false);
  }
}

function attachListListeners(container, entityType) {
  container?.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target instanceof HTMLSelectElement)) {
      return;
    }
    const rowElement = target.closest("[data-row-id]");
    const scope = target.getAttribute("data-scope");
    const field = target.getAttribute("data-field");
    if (!rowElement || !scope || !field) {
      return;
    }

    const rowId = rowElement.getAttribute("data-row-id");
    if (scope === "definition") {
      updateDefinitionField(entityType, rowId, field, target.value);
      return;
    }

    updateReviewField(entityType, rowId, field, target.value);
  });

  container?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLButtonElement) || target.getAttribute("data-action") !== "archive") {
      return;
    }
    const rowElement = target.closest("[data-row-id]");
    if (!rowElement) {
      return;
    }
    void archiveDefinition(entityType, rowElement.getAttribute("data-row-id"));
  });
}

function attachListeners() {
  attachListListeners(valuesList, "values");
  attachListListeners(principlesList, "principles");
  attachListListeners(objectivesList, "objectives");
  attachListListeners(goalsList, "goals");

  async function handleLoadPeriod() {
    const requestedPeriod = normalizeReviewPeriod(reviewPeriodInput?.value);
    if (requestedPeriod === activeReviewPeriod) {
      reviewPeriodInput.value = activeReviewPeriod;
      return;
    }
    await flushPendingReviewSave();
    await flushDefinitionSaves();
    await loadScorecardForPeriod(requestedPeriod);
  }

  loadPeriodBtn?.addEventListener("click", () => {
    void handleLoadPeriod();
  });

  reviewPeriodInput?.addEventListener("change", () => {
    void handleLoadPeriod();
  });

  reflectionGoingWell?.addEventListener("input", () => {
    state.reflection.goingWell = reflectionGoingWell.value;
    scheduleReviewSave();
  });
  reflectionNeedsAttention?.addEventListener("input", () => {
    state.reflection.needsAttention = reflectionNeedsAttention.value;
    scheduleReviewSave();
  });
  reflectionNextAction?.addEventListener("input", () => {
    state.reflection.nextAction = reflectionNextAction.value;
    scheduleReviewSave();
  });

  resetBtn?.addEventListener("click", () => {
    const confirmed = window.confirm("Reset this review period's scores and reflection to empty?");
    if (!confirmed) {
      return;
    }
    for (const entityType of Object.keys(ENTITY_META)) {
      for (const item of state[ENTITY_META[entityType].listKey]) {
        item.review = null;
      }
    }
    state.reflection = {
      goingWell: "",
      needsAttention: "",
      nextAction: "",
    };
    render();
    void persistReview();
  });

  saveNowBtn?.addEventListener("click", () => {
    void (async () => {
      try {
        const hadPendingDefinitions = hasPendingDefinitionSaves();
        const hadPendingReview = hasPendingReviewSave();

        if (!hadPendingDefinitions && !hadPendingReview) {
          setSaveMessage(formatSaveMeta(state.review), false, state.review.updatedAt ? "saved" : "idle");
          return;
        }
        setSavingMessage("Saving changes now...");
        await flushDefinitionSaves();
        await flushPendingReviewSave();
        if (hadPendingReview && !hasPendingReviewSave() && !savingReview) {
          setSaveMessage(formatSaveMeta(state.review), false, state.review.updatedAt ? "saved" : "idle");
          return;
        }
        if (hadPendingDefinitions && !hasPendingDefinitionSaves()) {
          setSaveMessage("Shared definitions saved.", false, "saved");
        }
      } catch (error) {
        console.error("[scorecard] Manual save failed", error);
        setSaveMessage(error?.message || "Could not save changes.", true);
      }
    })();
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
    if (!canAccessPage(role, "scorecard")) {
      window.location.href = "./unauthorized.html?page=scorecard";
      return;
    }

    renderTopNavigation({ role });
    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
    await loadScorecardForPeriod(currentReviewPeriod());
  } catch (error) {
    console.error("[scorecard] Init failed", error);
    setStatus(error?.message || "Could not initialize scorecard.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

attachListeners();
void init();
