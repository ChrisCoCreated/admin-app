import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260317";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const saveMessage = document.getElementById("saveMessage");
const saveNowBtn = document.getElementById("saveNowBtn");
const addGoalBtn = document.getElementById("addGoalBtn");
const goalsList = document.getElementById("goalsList");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

const state = {
  objectives: [],
  goals: [],
};

const pendingTimers = new Map();
const pendingPatches = new Map();
let loading = false;
let saving = false;

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setSaveMessage(message, tone = "idle", isError = false) {
  saveMessage.textContent = message;
  saveMessage.dataset.tone = tone;
  saveMessage.classList.toggle("error", isError);
}

function setBusy(isBusy) {
  loading = isBusy;
  addGoalBtn.disabled = isBusy || saving;
  saveNowBtn.disabled = isBusy || saving;
  updateMoveButtons();
}

function setSaving(isSaving) {
  saving = isSaving;
  addGoalBtn.disabled = loading || isSaving;
  saveNowBtn.disabled = loading || isSaving;
  updateMoveButtons();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function sortGoals() {
  state.goals = [...state.goals].sort((left, right) => {
    const orderDelta = Number(left.sortOrder || 0) - Number(right.sortOrder || 0);
    if (orderDelta !== 0) {
      return orderDelta;
    }
    return String(left?.title || "").localeCompare(String(right?.title || ""), undefined, {
      sensitivity: "base",
    });
  });
}

function resequenceGoals() {
  state.goals = state.goals.map((goal, index) => ({
    ...goal,
    sortOrder: index + 1,
  }));
}

function syncUpdatedGoal(updated) {
  if (!updated?.id) {
    return;
  }
  const index = state.goals.findIndex((item) => item.id === updated.id);
  if (index >= 0) {
    state.goals[index] = {
      ...state.goals[index],
      ...updated,
    };
  }
}

function updateMoveButtons() {
  const rows = [...goalsList.querySelectorAll("[data-row-id]")];
  rows.forEach((row, index) => {
    const upButton = row.querySelector('[data-action="move-up"]');
    const downButton = row.querySelector('[data-action="move-down"]');
    if (upButton instanceof HTMLButtonElement) {
      upButton.disabled = index === 0 || loading || saving;
    }
    if (downButton instanceof HTMLButtonElement) {
      downButton.disabled = index === rows.length - 1 || loading || saving;
    }
  });
}

function updateGoalObjectiveDisplay(rowId) {
  const row = goalsList.querySelector(`[data-row-id="${CSS.escape(rowId)}"]`);
  if (!row) {
    return;
  }
  const goal = findGoal(rowId);
  if (!goal) {
    return;
  }
  const display = row.querySelector('[data-role="objective-title"]');
  if (display instanceof HTMLInputElement) {
    display.value = objectiveTitle(goal.objectiveId || "");
  }
}

function objectiveOptions(selectedId) {
  const options = ['<option value="">Choose linked objective</option>'];
  for (const objective of state.objectives) {
    options.push(
      `<option value="${escapeHtml(objective.id)}"${objective.id === selectedId ? " selected" : ""}>${escapeHtml(
        objective.title || "(Untitled objective)"
      )}</option>`
    );
  }
  return options.join("");
}

function objectiveTitle(objectiveId) {
  return state.objectives.find((item) => item.id === objectiveId)?.title || "";
}

async function loadData() {
  try {
    setBusy(true);
    setSaveMessage("Loading goals and objectives...", "saving");
    const objectivesPayload = await directoryApi.listPerformanceScorecardDefinitions({
      entityType: "objectives",
      activeOnly: "false",
    });
    state.objectives = Array.isArray(objectivesPayload?.definitions) ? objectivesPayload.definitions : [];

    const goalsPayload = await directoryApi.listPerformanceScorecardDefinitions({
      entityType: "goals",
      activeOnly: "false",
    });
    state.goals = Array.isArray(goalsPayload?.definitions) ? goalsPayload.definitions : [];
    sortGoals();
    render();
    setSaveMessage("Shared goals loaded. Changes save automatically.", "idle");
  } catch (error) {
    console.error("[scorecard-goals] Failed to load goals", error);
    setSaveMessage(error?.message || "Could not load shared goals.", "error", true);
  } finally {
    setBusy(false);
  }
}

function render() {
  if (!state.goals.length) {
    goalsList.innerHTML =
      '<p class="muted scorecard-empty-state">No goals defined yet. Add a goal to build the shared delivery library.</p>';
    return;
  }

  goalsList.innerHTML = state.goals
    .map(
      (goal, index) => `
        <article class="scorecard-entry ${goal.isActive === false ? "is-archived" : ""}" data-row-id="${escapeHtml(goal.id)}">
          <div class="scorecard-entry-head">
            <div>
              <div class="scorecard-entry-title-row">
                <h3>Goal ${index + 1}</h3>
                ${goal.isActive === false ? '<span class="scorecard-archived-chip">Archived</span>' : ""}
              </div>
              <p class="muted">This shared goal appears on each review period and stays linked to one objective.</p>
            </div>
            <div class="scorecard-entry-actions">
              <button class="secondary scorecard-remove-btn" type="button" data-action="move-up">Move up</button>
              <button class="secondary scorecard-remove-btn" type="button" data-action="move-down">Move down</button>
              <button class="secondary scorecard-remove-btn" type="button" data-action="archive"${
                goal.isActive === false ? " disabled" : ""
              }>Archive</button>
            </div>
          </div>
          <div class="scorecard-form-grid">
            <label class="field scorecard-span-2">
              Goal
              <input data-field="title" type="text" value="${escapeHtml(goal.title || "")}" placeholder="Describe the goal or initiative" />
            </label>
            <label class="field scorecard-span-2">
              Details
              <textarea data-field="description" placeholder="Add context, scope, or notes for this shared goal.">${escapeHtml(
                goal.description || ""
              )}</textarea>
            </label>
            <label class="field">
              Linked objective
              <select data-field="objectiveId">${objectiveOptions(goal.objectiveId || "")}</select>
            </label>
            <label class="field is-muted-block">
              Current linked objective
              <input data-role="objective-title" type="text" value="${escapeHtml(objectiveTitle(goal.objectiveId || ""))}" readonly />
            </label>
          </div>
        </article>
      `
    )
    .join("");

  updateMoveButtons();
}

function findGoal(id) {
  return state.goals.find((item) => item.id === id) || null;
}

function queueSave(rowId, patch) {
  pendingPatches.set(rowId, {
    ...(pendingPatches.get(rowId) || {}),
    ...patch,
  });

  const existing = pendingTimers.get(rowId);
  if (existing) {
    window.clearTimeout(existing);
  }

  setSaveMessage("Unsaved goal changes. Saving shortly...", "pending");
  const timer = window.setTimeout(async () => {
    pendingTimers.delete(rowId);
    const nextPatch = pendingPatches.get(rowId);
    pendingPatches.delete(rowId);
    try {
      setSaving(true);
      setSaveMessage("Saving shared goals...", "saving");
      const response = await directoryApi.updatePerformanceScorecardDefinition({
        entityType: "goals",
        id: rowId,
        patch: nextPatch,
      });
      const updated = response?.definition || null;
      if (updated?.id) {
        const index = state.goals.findIndex((item) => item.id === updated.id);
        if (index >= 0) {
          syncUpdatedGoal(updated);
        }
      }
      setSaveMessage("Shared goals saved.", "saved");
    } catch (error) {
      console.error("[scorecard-goals] Failed to save goal", error);
      setSaveMessage(error?.message || "Could not save the goal.", "error", true);
    } finally {
      setSaving(false);
    }
  }, 260);

  pendingTimers.set(rowId, timer);
}

async function flushSaves() {
  const tasks = [];
  for (const [rowId, timer] of pendingTimers.entries()) {
    window.clearTimeout(timer);
    pendingTimers.delete(rowId);
    const patch = pendingPatches.get(rowId);
    pendingPatches.delete(rowId);
    tasks.push(
      directoryApi
        .updatePerformanceScorecardDefinition({
          entityType: "goals",
          id: rowId,
          patch,
        })
        .then((response) => {
          const updated = response?.definition || null;
          if (updated?.id) {
            const index = state.goals.findIndex((item) => item.id === updated.id);
            if (index >= 0) {
              syncUpdatedGoal(updated);
            }
          }
        })
    );
  }

  if (!tasks.length) {
    setSaveMessage("Everything is already saved.", "saved");
    return;
  }

  setSaving(true);
  setSaveMessage("Saving shared goals...", "saving");
  try {
    await Promise.all(tasks);
    sortGoals();
    render();
    setSaveMessage("Shared goals saved.", "saved");
  } catch (error) {
    console.error("[scorecard-goals] Failed to flush goal saves", error);
    setSaveMessage(error?.message || "Could not save all goal changes.", "error", true);
  } finally {
    setSaving(false);
  }
}

async function createGoal() {
  if (!state.objectives.length) {
    setSaveMessage("Add at least one objective in Scorecard Setup before creating goals.", "error", true);
    return;
  }

  try {
    setSaving(true);
    setSaveMessage("Creating goal...", "saving");
    const response = await directoryApi.createPerformanceScorecardDefinition({
      entityType: "goals",
      definition: {
        title: "",
        objective_id: state.objectives[0]?.id || null,
      },
    });
    const goal = response?.definition || null;
    if (goal?.id) {
      state.goals.push(goal);
      sortGoals();
      render();
      setSaveMessage("New goal saved.", "saved");
    }
  } catch (error) {
    console.error("[scorecard-goals] Failed to create goal", error);
    setSaveMessage(error?.message || "Could not create the goal.", "error", true);
  } finally {
    setSaving(false);
  }
}

async function archiveGoal(rowId) {
  try {
    await flushSaves();
    setSaving(true);
    setSaveMessage("Archiving goal...", "saving");
    const response = await directoryApi.updatePerformanceScorecardDefinition({
      entityType: "goals",
      id: rowId,
      patch: { is_active: false },
    });
    const updated = response?.definition || null;
    if (updated?.id) {
      const index = state.goals.findIndex((item) => item.id === updated.id);
      if (index >= 0) {
        state.goals[index] = updated;
        sortGoals();
        render();
      }
    }
    setSaveMessage("Goal archived.", "saved");
  } catch (error) {
    console.error("[scorecard-goals] Failed to archive goal", error);
    setSaveMessage(error?.message || "Could not archive the goal.", "error", true);
  } finally {
    setSaving(false);
  }
}

function attachListeners() {
  addGoalBtn?.addEventListener("click", () => {
    void createGoal();
  });

  saveNowBtn?.addEventListener("click", () => {
    void flushSaves();
  });

  goalsList?.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLSelectElement || target instanceof HTMLTextAreaElement)) {
      return;
    }
    const row = target.closest("[data-row-id]");
    const field = target.getAttribute("data-field");
    if (!row || !field) {
      return;
    }
    const rowId = row.getAttribute("data-row-id");
    const goal = findGoal(rowId);
    if (!goal) {
      return;
    }
    if (field === "objectiveId") {
      goal.objectiveId = target.value || null;
      updateGoalObjectiveDisplay(rowId);
      queueSave(rowId, { objective_id: target.value || null });
      return;
    }
    goal[field] = target.value;
    queueSave(rowId, { [field]: target.value });
  });

  goalsList?.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLButtonElement)) {
      return;
    }
    const row = target.closest("[data-row-id]");
    if (!row) {
      return;
    }
    const action = target.getAttribute("data-action");
    const rowId = row.getAttribute("data-row-id");
    if (action === "archive") {
      void archiveGoal(rowId);
      return;
    }
    if (action === "move-up" || action === "move-down") {
      void moveGoal(rowId, action === "move-up" ? -1 : 1);
    }
  });
}

async function moveGoal(rowId, offset) {
  try {
    await flushSaves();
    const currentIndex = state.goals.findIndex((item) => item.id === rowId);
    const targetIndex = currentIndex + offset;
    if (currentIndex < 0 || targetIndex < 0 || targetIndex >= state.goals.length) {
      return;
    }

    const currentGoal = state.goals[currentIndex];
    state.goals.splice(currentIndex, 1);
    state.goals.splice(targetIndex, 0, currentGoal);
    resequenceGoals();
    render();
    setSaveMessage("Saving order...", "saving");
    await Promise.all(
      state.goals.map((goal) =>
        directoryApi.updatePerformanceScorecardDefinition({
          entityType: "goals",
          id: goal.id,
          patch: { sort_order: goal.sortOrder },
        })
      )
    );
    setSaveMessage("Order saved.", "saved");
  } catch (error) {
    console.error("[scorecard-goals] Failed to move goal", error);
    setSaveMessage(error?.message || "Could not update the order.", "error", true);
    await loadData();
  }
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
    if (!canAccessPage(role, "scorecardgoals")) {
      window.location.href = "./unauthorized.html?page=scorecard-goals";
      return;
    }

    renderTopNavigation({ role });
    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
    await loadData();
  } catch (error) {
    console.error("[scorecard-goals] Init failed", error);
    setStatus(error?.message || "Could not initialize goal setup.", true);
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
