import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260313";

const STORAGE_KEY = "thrive.performance-alignment-scorecard.v1";
const MIN_ROWS = {
  values: 4,
  principles: 3,
  objectives: 3,
  goals: 4,
};

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const saveMessage = document.getElementById("saveMessage");
const reviewPeriodInput = document.getElementById("reviewPeriodInput");
const valuesList = document.getElementById("valuesList");
const principlesList = document.getElementById("principlesList");
const objectivesList = document.getElementById("objectivesList");
const goalsList = document.getElementById("goalsList");
const snapshotGrid = document.getElementById("snapshotGrid");
const addValueBtn = document.getElementById("addValueBtn");
const addPrincipleBtn = document.getElementById("addPrincipleBtn");
const addObjectiveBtn = document.getElementById("addObjectiveBtn");
const addGoalBtn = document.getElementById("addGoalBtn");
const resetBtn = document.getElementById("resetBtn");
const reflectionGoingWell = document.getElementById("reflectionGoingWell");
const reflectionNeedsAttention = document.getElementById("reflectionNeedsAttention");
const reflectionNextAction = document.getElementById("reflectionNextAction");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let state = createDefaultState();
let saveTimer = null;

function uid(prefix) {
  return `${prefix}-${Math.random().toString(36).slice(2, 10)}`;
}

function createValueRow() {
  return {
    id: uid("value"),
    value: "",
    score: "",
    evidence: "",
    action: "",
  };
}

function createPrincipleRow() {
  return {
    id: uid("principle"),
    principle: "",
    score: "",
    example: "",
    adjustment: "",
  };
}

function createObjectiveRow() {
  return {
    id: uid("objective"),
    objective: "",
    owner: "",
    score: "",
    risks: "",
  };
}

function createGoalRow() {
  return {
    id: uid("goal"),
    goal: "",
    linkedObjective: "",
    owner: "",
    status: "On track",
    score: "",
    nextAction: "",
  };
}

function currentReviewPeriod() {
  return new Intl.DateTimeFormat(undefined, {
    month: "long",
    year: "numeric",
  }).format(new Date());
}

function createDefaultState() {
  return {
    reviewPeriod: currentReviewPeriod(),
    values: Array.from({ length: MIN_ROWS.values }, () => createValueRow()),
    principles: Array.from({ length: MIN_ROWS.principles }, () => createPrincipleRow()),
    objectives: Array.from({ length: MIN_ROWS.objectives }, () => createObjectiveRow()),
    goals: Array.from({ length: MIN_ROWS.goals }, () => createGoalRow()),
    reflection: {
      goingWell: "",
      needsAttention: "",
      nextAction: "",
    },
  };
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setSaveMessage(message, isError = false) {
  saveMessage.textContent = message;
  saveMessage.classList.toggle("error", isError);
}

function cloneRow(factory, raw) {
  return {
    ...factory(),
    ...(raw || {}),
  };
}

function ensureMinimumRows(rows, minimum, factory) {
  const nextRows = Array.isArray(rows) ? [...rows] : [];
  while (nextRows.length < minimum) {
    nextRows.push(factory());
  }
  return nextRows;
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      state = createDefaultState();
      return;
    }

    const parsed = JSON.parse(raw);
    state = {
      reviewPeriod: String(parsed?.reviewPeriod || currentReviewPeriod()),
      values: ensureMinimumRows(
        Array.isArray(parsed?.values) ? parsed.values.map((row) => cloneRow(createValueRow, row)) : [],
        MIN_ROWS.values,
        createValueRow
      ),
      principles: ensureMinimumRows(
        Array.isArray(parsed?.principles) ? parsed.principles.map((row) => cloneRow(createPrincipleRow, row)) : [],
        MIN_ROWS.principles,
        createPrincipleRow
      ),
      objectives: ensureMinimumRows(
        Array.isArray(parsed?.objectives) ? parsed.objectives.map((row) => cloneRow(createObjectiveRow, row)) : [],
        MIN_ROWS.objectives,
        createObjectiveRow
      ),
      goals: ensureMinimumRows(
        Array.isArray(parsed?.goals) ? parsed.goals.map((row) => cloneRow(createGoalRow, row)) : [],
        MIN_ROWS.goals,
        createGoalRow
      ),
      reflection: {
        goingWell: String(parsed?.reflection?.goingWell || ""),
        needsAttention: String(parsed?.reflection?.needsAttention || ""),
        nextAction: String(parsed?.reflection?.nextAction || ""),
      },
    };
  } catch (error) {
    console.error("[scorecard] Failed to load saved state", error);
    state = createDefaultState();
    setSaveMessage("Could not load saved scorecard. Showing a fresh page instead.", true);
  }
}

function persistState() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    const time = new Date().toLocaleTimeString([], {
      hour: "2-digit",
      minute: "2-digit",
    });
    setSaveMessage(`Saved locally at ${time}.`);
  } catch (error) {
    console.error("[scorecard] Failed to save state", error);
    setSaveMessage("Could not save changes on this device.", true);
  }
}

function scheduleSave() {
  setSaveMessage("Saving...");
  window.clearTimeout(saveTimer);
  saveTimer = window.setTimeout(() => {
    persistState();
  }, 220);
}

function getObjectiveOptions() {
  return state.objectives
    .map((row) => String(row?.objective || "").trim())
    .filter(Boolean);
}

function averageScore(rows, labelKey) {
  const scores = rows
    .filter((row) => String(row?.[labelKey] || "").trim())
    .map((row) => Number(row?.score))
    .filter((value) => Number.isFinite(value) && value >= 1 && value <= 5);

  if (!scores.length) {
    return null;
  }

  return scores.reduce((sum, value) => sum + value, 0) / scores.length;
}

function formatAverage(value) {
  return value == null ? "Not scored" : `${value.toFixed(1)} / 5`;
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

function buildScoreOptions(selected) {
  const options = ['<option value="">Select score</option>'];
  for (let score = 1; score <= 5; score += 1) {
    const isSelected = String(selected) === String(score) ? " selected" : "";
    options.push(`<option value="${score}"${isSelected}>${score}</option>`);
  }
  return options.join("");
}

function buildRowShell(title, description, content, removeLabel, removeDisabled, rowId) {
  return `
    <article class="scorecard-entry" data-row-id="${rowId}">
      <div class="scorecard-entry-head">
        <div>
          <h3>${title}</h3>
          <p class="muted">${description}</p>
        </div>
        <button class="secondary scorecard-remove-btn" type="button" data-action="remove"${removeDisabled ? " disabled" : ""}>
          ${removeLabel}
        </button>
      </div>
      ${content}
    </article>
  `;
}

function renderValues() {
  valuesList.innerHTML = state.values
    .map((row, index) => {
      const showAction = Number(row?.score) > 0 && Number(row?.score) < 4;
      return buildRowShell(
        `Value ${index + 1}`,
        "Score how clearly this value is showing up in behaviour and decisions.",
        `
          <div class="scorecard-form-grid">
            <label class="field">
              Value
              <input data-field="value" type="text" value="${escapeHtml(row.value)}" placeholder="e.g. Accountability" />
            </label>
            <label class="field">
              Score (1-5)
              <select data-field="score">${buildScoreOptions(row.score)}</select>
            </label>
            <label class="field scorecard-span-2">
              Evidence / examples
              <textarea data-field="evidence" placeholder="What supports this score?">${escapeHtml(row.evidence)}</textarea>
            </label>
            <label class="field scorecard-span-2 ${showAction ? "" : "is-muted-block"}">
              Improvement action ${showAction ? "" : "(only needed below 4)"}
              <textarea data-field="action" placeholder="What needs to change next?">${escapeHtml(row.action)}</textarea>
            </label>
          </div>
        `,
        "Remove",
        state.values.length <= MIN_ROWS.values,
        row.id
      );
    })
    .join("");
}

function renderPrinciples() {
  principlesList.innerHTML = state.principles
    .map((row, index) =>
      buildRowShell(
        `Principle ${index + 1}`,
        "Review whether recent decisions were guided by this principle.",
        `
          <div class="scorecard-form-grid">
            <label class="field">
              Principle
              <input data-field="principle" type="text" value="${escapeHtml(row.principle)}" placeholder="e.g. Make decisions close to the client" />
            </label>
            <label class="field">
              Score (1-5)
              <select data-field="score">${buildScoreOptions(row.score)}</select>
            </label>
            <label class="field scorecard-span-2">
              Example decision demonstrating the principle
              <textarea data-field="example" placeholder="Which recent decision shows this in practice?">${escapeHtml(row.example)}</textarea>
            </label>
            <label class="field scorecard-span-2">
              Adjustment needed
              <textarea data-field="adjustment" placeholder="What should we refine next period?">${escapeHtml(row.adjustment)}</textarea>
            </label>
          </div>
        `,
        "Remove",
        state.principles.length <= MIN_ROWS.principles,
        row.id
      )
    )
    .join("");
}

function renderObjectives() {
  objectivesList.innerHTML = state.objectives
    .map((row, index) =>
      buildRowShell(
        `Objective ${index + 1}`,
        "Track the strategic objective owner, score, and current blockers.",
        `
          <div class="scorecard-form-grid">
            <label class="field scorecard-span-2">
              Objective
              <input data-field="objective" type="text" value="${escapeHtml(row.objective)}" placeholder="Describe the objective" />
            </label>
            <label class="field">
              Owner
              <input data-field="owner" type="text" value="${escapeHtml(row.owner)}" placeholder="Owner name" />
            </label>
            <label class="field">
              Progress score (1-5)
              <select data-field="score">${buildScoreOptions(row.score)}</select>
            </label>
            <label class="field scorecard-span-2">
              Key risks or blockers
              <textarea data-field="risks" placeholder="What is slowing progress or increasing risk?">${escapeHtml(row.risks)}</textarea>
            </label>
          </div>
        `,
        "Remove",
        state.objectives.length <= MIN_ROWS.objectives,
        row.id
      )
    )
    .join("");
}

function buildObjectiveSelect(selected) {
  const options = ['<option value="">Choose linked objective</option>'];
  for (const objective of getObjectiveOptions()) {
    const isSelected = objective === selected ? " selected" : "";
    options.push(`<option value="${escapeHtml(objective)}"${isSelected}>${escapeHtml(objective)}</option>`);
  }
  return options.join("");
}

function renderGoals() {
  goalsList.innerHTML = state.goals
    .map((row, index) =>
      buildRowShell(
        `Goal ${index + 1}`,
        "Link each delivery goal to a strategic objective and capture the next action.",
        `
          <div class="scorecard-form-grid">
            <label class="field scorecard-span-2">
              Goal
              <input data-field="goal" type="text" value="${escapeHtml(row.goal)}" placeholder="Describe the initiative or goal" />
            </label>
            <label class="field">
              Linked objective
              <select data-field="linkedObjective">${buildObjectiveSelect(row.linkedObjective)}</select>
            </label>
            <label class="field">
              Owner
              <input data-field="owner" type="text" value="${escapeHtml(row.owner)}" placeholder="Owner name" />
            </label>
            <label class="field">
              Status
              <select data-field="status">
                <option value="On track"${row.status === "On track" ? " selected" : ""}>On track</option>
                <option value="At risk"${row.status === "At risk" ? " selected" : ""}>At risk</option>
                <option value="Off track"${row.status === "Off track" ? " selected" : ""}>Off track</option>
              </select>
            </label>
            <label class="field">
              Progress score (1-5)
              <select data-field="score">${buildScoreOptions(row.score)}</select>
            </label>
            <label class="field scorecard-span-2">
              Next key action
              <textarea data-field="nextAction" placeholder="What happens next to move this forward?">${escapeHtml(row.nextAction)}</textarea>
            </label>
          </div>
        `,
        "Remove",
        state.goals.length <= MIN_ROWS.goals,
        row.id
      )
    )
    .join("");
}

function renderSnapshot() {
  const cards = [
    {
      label: "Values Score",
      value: averageScore(state.values, "value"),
      note: `${countScoredRows(state.values, "value")} values scored`,
    },
    {
      label: "Decision Principles Score",
      value: averageScore(state.principles, "principle"),
      note: `${countScoredRows(state.principles, "principle")} principles scored`,
    },
    {
      label: "Strategic Objectives Score",
      value: averageScore(state.objectives, "objective"),
      note: `${countScoredRows(state.objectives, "objective")} objectives scored`,
    },
    {
      label: "Goals Delivery Score",
      value: averageScore(state.goals, "goal"),
      note: `${countScoredRows(state.goals, "goal")} goals scored`,
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

function countScoredRows(rows, labelKey) {
  return rows.filter((row) => {
    const label = String(row?.[labelKey] || "").trim();
    const score = Number(row?.score);
    return label && Number.isFinite(score) && score >= 1 && score <= 5;
  }).length;
}

function renderReflection() {
  reviewPeriodInput.value = state.reviewPeriod;
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

function updateCollectionRow(collectionKey, rowId, field, value) {
  const collection = state[collectionKey];
  const row = collection.find((entry) => entry.id === rowId);
  if (!row) {
    return;
  }
  row[field] = value;

  if (collectionKey === "objectives" && field === "objective") {
    for (const goal of state.goals) {
      if (goal.linkedObjective && !getObjectiveOptions().includes(goal.linkedObjective)) {
        goal.linkedObjective = "";
      }
    }
    renderGoals();
    renderSnapshot();
  } else if (collectionKey === "values" && field === "score") {
    renderValues();
    renderSnapshot();
  } else {
    renderSnapshot();
    if (collectionKey === "goals" && field === "status") {
      renderGoals();
    }
  }

  scheduleSave();
}

function addRow(collectionKey) {
  if (collectionKey === "values") {
    state.values.push(createValueRow());
  } else if (collectionKey === "principles") {
    state.principles.push(createPrincipleRow());
  } else if (collectionKey === "objectives") {
    state.objectives.push(createObjectiveRow());
  } else if (collectionKey === "goals") {
    state.goals.push(createGoalRow());
  }
  render();
  scheduleSave();
}

function removeRow(collectionKey, rowId) {
  if (state[collectionKey].length <= MIN_ROWS[collectionKey]) {
    return;
  }

  state[collectionKey] = state[collectionKey].filter((row) => row.id !== rowId);
  if (collectionKey === "objectives") {
    const validObjectives = getObjectiveOptions();
    for (const goal of state.goals) {
      if (goal.linkedObjective && !validObjectives.includes(goal.linkedObjective)) {
        goal.linkedObjective = "";
      }
    }
  }
  render();
  scheduleSave();
}

function attachCollectionListeners(container, collectionKey) {
  container?.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement || target instanceof HTMLSelectElement)) {
      return;
    }
    const rowElement = target.closest("[data-row-id]");
    const field = target.getAttribute("data-field");
    if (!rowElement || !field) {
      return;
    }
    updateCollectionRow(collectionKey, rowElement.getAttribute("data-row-id"), field, target.value);
  });

  container?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLButtonElement) || target.getAttribute("data-action") !== "remove") {
      return;
    }
    const rowElement = target.closest("[data-row-id]");
    if (!rowElement) {
      return;
    }
    removeRow(collectionKey, rowElement.getAttribute("data-row-id"));
  });
}

function attachListeners() {
  attachCollectionListeners(valuesList, "values");
  attachCollectionListeners(principlesList, "principles");
  attachCollectionListeners(objectivesList, "objectives");
  attachCollectionListeners(goalsList, "goals");

  addValueBtn?.addEventListener("click", () => addRow("values"));
  addPrincipleBtn?.addEventListener("click", () => addRow("principles"));
  addObjectiveBtn?.addEventListener("click", () => addRow("objectives"));
  addGoalBtn?.addEventListener("click", () => addRow("goals"));

  reviewPeriodInput?.addEventListener("input", () => {
    state.reviewPeriod = reviewPeriodInput.value;
    scheduleSave();
  });

  reflectionGoingWell?.addEventListener("input", () => {
    state.reflection.goingWell = reflectionGoingWell.value;
    scheduleSave();
  });
  reflectionNeedsAttention?.addEventListener("input", () => {
    state.reflection.needsAttention = reflectionNeedsAttention.value;
    scheduleSave();
  });
  reflectionNextAction?.addEventListener("input", () => {
    state.reflection.nextAction = reflectionNextAction.value;
    scheduleSave();
  });

  resetBtn?.addEventListener("click", () => {
    const confirmed = window.confirm("Reset this scorecard to a fresh template? This only affects the saved copy on this device.");
    if (!confirmed) {
      return;
    }
    state = createDefaultState();
    render();
    persistState();
  });
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
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
    loadState();
    render();

    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
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
