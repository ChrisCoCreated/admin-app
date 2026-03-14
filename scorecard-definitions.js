import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260314";

const ENTITY_CONFIG = {
  values: {
    label: "value",
    field: "name",
    listId: "valuesList",
    addId: "addValueBtn",
    title: "Value",
    placeholder: "e.g. Accountability",
    createPayload: () => ({ name: "" }),
  },
  principles: {
    label: "principle",
    field: "name",
    listId: "principlesList",
    addId: "addPrincipleBtn",
    title: "Principle",
    placeholder: "e.g. Make decisions close to the client",
    createPayload: () => ({ name: "" }),
  },
  objectives: {
    label: "objective",
    field: "title",
    listId: "objectivesList",
    addId: "addObjectiveBtn",
    title: "Objective",
    placeholder: "Describe the strategic objective",
    createPayload: () => ({ title: "" }),
  },
};

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const saveMessage = document.getElementById("saveMessage");
const saveNowBtn = document.getElementById("saveNowBtn");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

const state = {
  values: [],
  principles: [],
  objectives: [],
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
  for (const config of Object.values(ENTITY_CONFIG)) {
    const button = document.getElementById(config.addId);
    if (button) {
      button.disabled = isBusy || saving;
    }
  }
  if (saveNowBtn) {
    saveNowBtn.disabled = isBusy || saving;
  }
  renderAllMoveButtons();
}

function setSaving(isSaving) {
  saving = isSaving;
  for (const config of Object.values(ENTITY_CONFIG)) {
    const button = document.getElementById(config.addId);
    if (button) {
      button.disabled = loading || isSaving;
    }
  }
  if (saveNowBtn) {
    saveNowBtn.disabled = loading || isSaving;
  }
  renderAllMoveButtons();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function sortItems(entityType) {
  const field = ENTITY_CONFIG[entityType].field;
  state[entityType] = [...state[entityType]].sort((left, right) => {
    const orderDelta = Number(left.sortOrder || 0) - Number(right.sortOrder || 0);
    if (orderDelta !== 0) {
      return orderDelta;
    }
    return String(left?.[field] || "").localeCompare(String(right?.[field] || ""), undefined, {
      sensitivity: "base",
    });
  });
}

function resequenceItems(entityType) {
  state[entityType] = state[entityType].map((item, index) => ({
    ...item,
    sortOrder: index + 1,
  }));
}

function itemPatchField(field) {
  if (field === "title") {
    return "title";
  }
  if (field === "description") {
    return "description";
  }
  return "name";
}

function syncUpdatedItem(entityType, updated) {
  if (!updated?.id) {
    return;
  }
  const index = state[entityType].findIndex((item) => item.id === updated.id);
  if (index >= 0) {
    state[entityType][index] = {
      ...state[entityType][index],
      ...updated,
    };
  }
}

function updateMoveButtons(entityType) {
  const container = document.getElementById(ENTITY_CONFIG[entityType].listId);
  if (!container) {
    return;
  }
  const rows = [...container.querySelectorAll("[data-row-id]")];
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

function renderAllMoveButtons() {
  for (const entityType of Object.keys(ENTITY_CONFIG)) {
    updateMoveButtons(entityType);
  }
}

async function loadDefinitions() {
  try {
    setBusy(true);
    setSaveMessage("Loading shared definitions...", "saving");
    for (const entityType of Object.keys(ENTITY_CONFIG)) {
      const payload = await directoryApi.listPerformanceScorecardDefinitions({
        entityType,
        activeOnly: "false",
      });
      state[entityType] = Array.isArray(payload?.definitions) ? payload.definitions : [];
      sortItems(entityType);
    }
    render();
    setSaveMessage("Shared definitions loaded. Changes save automatically.", "idle");
  } catch (error) {
    console.error("[scorecard-definitions] Failed to load definitions", error);
    setSaveMessage(error?.message || "Could not load shared definitions.", "error", true);
  } finally {
    setBusy(false);
  }
}

function renderEntity(entityType) {
  const config = ENTITY_CONFIG[entityType];
  const container = document.getElementById(config.listId);
  if (!container) {
    return;
  }

  if (!state[entityType].length) {
    container.innerHTML = `<p class="muted scorecard-empty-state">No ${config.label}s defined yet. Add one to build the shared scorecard library.</p>`;
    return;
  }

  container.innerHTML = state[entityType]
    .map(
      (item, index) => `
        <article class="scorecard-entry ${item.isActive === false ? "is-archived" : ""}" data-entity="${entityType}" data-row-id="${escapeHtml(item.id)}">
          <div class="scorecard-entry-head">
            <div>
              <div class="scorecard-entry-title-row">
                <h3>${config.title} ${index + 1}</h3>
                ${item.isActive === false ? '<span class="scorecard-archived-chip">Archived</span>' : ""}
              </div>
              <p class="muted">This shared ${config.label} appears on every review period until it is archived.</p>
            </div>
            <div class="scorecard-entry-actions">
              <button class="secondary scorecard-remove-btn" type="button" data-action="move-up">Move up</button>
              <button class="secondary scorecard-remove-btn" type="button" data-action="move-down">Move down</button>
              <button class="secondary scorecard-remove-btn" type="button" data-action="archive"${
                item.isActive === false ? " disabled" : ""
              }>Archive</button>
            </div>
          </div>
          <div class="scorecard-form-grid">
            <label class="field ${entityType === "objectives" ? "scorecard-span-2" : ""}">
              ${config.title}
              <input
                data-field="${config.field}"
                type="text"
                value="${escapeHtml(item[config.field] || "")}"
                placeholder="${escapeHtml(config.placeholder)}"
              />
            </label>
            <label class="field scorecard-span-2">
              Details
              <textarea data-field="description" placeholder="Add context, intent, or notes for this shared ${escapeHtml(
                config.label
              )}.">${escapeHtml(item.description || "")}</textarea>
            </label>
          </div>
        </article>
      `
    )
    .join("");

  updateMoveButtons(entityType);
}

function render() {
  renderEntity("values");
  renderEntity("principles");
  renderEntity("objectives");
}

function findItem(entityType, id) {
  return state[entityType].find((item) => item.id === id) || null;
}

function queueSave(entityType, rowId, patch) {
  const key = `${entityType}:${rowId}`;
  pendingPatches.set(key, {
    ...(pendingPatches.get(key) || {}),
    ...patch,
  });

  const existing = pendingTimers.get(key);
  if (existing) {
    window.clearTimeout(existing);
  }

  setSaveMessage("Unsaved definition changes. Saving shortly...", "pending");
  const timer = window.setTimeout(async () => {
    pendingTimers.delete(key);
    const nextPatch = pendingPatches.get(key);
    pendingPatches.delete(key);
    try {
      setSaving(true);
      setSaveMessage("Saving shared definitions...", "saving");
      const response = await directoryApi.updatePerformanceScorecardDefinition({
        entityType,
        id: rowId,
        patch: nextPatch,
      });
      const updated = response?.definition || null;
      if (updated?.id) {
        const index = state[entityType].findIndex((item) => item.id === updated.id);
        if (index >= 0) {
          syncUpdatedItem(entityType, updated);
        }
      }
      setSaveMessage("Shared definitions saved.", "saved");
    } catch (error) {
      console.error("[scorecard-definitions] Failed to save definition", error);
      setSaveMessage(error?.message || "Could not save the definition.", "error", true);
    } finally {
      setSaving(false);
    }
  }, 260);

  pendingTimers.set(key, timer);
}

async function flushSaves() {
  const tasks = [];
  for (const [key, timer] of pendingTimers.entries()) {
    window.clearTimeout(timer);
    pendingTimers.delete(key);
    const [entityType, rowId] = key.split(":");
    const patch = pendingPatches.get(key);
    pendingPatches.delete(key);
    tasks.push(
      directoryApi
        .updatePerformanceScorecardDefinition({
          entityType,
          id: rowId,
          patch,
        })
        .then((response) => {
          const updated = response?.definition || null;
          if (updated?.id) {
            const index = state[entityType].findIndex((item) => item.id === updated.id);
            if (index >= 0) {
              syncUpdatedItem(entityType, updated);
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
  setSaveMessage("Saving shared definitions...", "saving");
  try {
    await Promise.all(tasks);
    render();
    setSaveMessage("Shared definitions saved.", "saved");
  } catch (error) {
    console.error("[scorecard-definitions] Failed to flush definition saves", error);
    setSaveMessage(error?.message || "Could not save all definition changes.", "error", true);
  } finally {
    setSaving(false);
  }
}

async function createDefinition(entityType) {
  try {
    setSaving(true);
    setSaveMessage(`Creating ${ENTITY_CONFIG[entityType].label}...`, "saving");
    const response = await directoryApi.createPerformanceScorecardDefinition({
      entityType,
      definition: ENTITY_CONFIG[entityType].createPayload(),
    });
    const definition = response?.definition || null;
    if (definition?.id) {
      state[entityType].push(definition);
      sortItems(entityType);
      renderEntity(entityType);
      setSaveMessage(`New ${ENTITY_CONFIG[entityType].label} saved.`, "saved");
    }
  } catch (error) {
    console.error("[scorecard-definitions] Failed to create definition", error);
    setSaveMessage(error?.message || "Could not create the definition.", "error", true);
  } finally {
    setSaving(false);
  }
}

async function archiveDefinition(entityType, rowId) {
  try {
    await flushSaves();
    setSaving(true);
    setSaveMessage(`Archiving ${ENTITY_CONFIG[entityType].label}...`, "saving");
    const response = await directoryApi.updatePerformanceScorecardDefinition({
      entityType,
      id: rowId,
      patch: { is_active: false },
    });
    const updated = response?.definition || null;
    if (updated?.id) {
      const index = state[entityType].findIndex((item) => item.id === updated.id);
      if (index >= 0) {
        state[entityType][index] = updated;
        sortItems(entityType);
        renderEntity(entityType);
      }
    }
    setSaveMessage(`${ENTITY_CONFIG[entityType].title} archived.`, "saved");
  } catch (error) {
    console.error("[scorecard-definitions] Failed to archive definition", error);
    setSaveMessage(error?.message || "Could not archive the definition.", "error", true);
  } finally {
    setSaving(false);
  }
}

function attachEntityListeners(entityType) {
  const container = document.getElementById(ENTITY_CONFIG[entityType].listId);
  container?.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement || target instanceof HTMLTextAreaElement)) {
      return;
    }
    const row = target.closest("[data-row-id]");
    const field = target.getAttribute("data-field");
    if (!row || !field) {
      return;
    }
    const rowId = row.getAttribute("data-row-id");
    const item = findItem(entityType, rowId);
    if (!item) {
      return;
    }
    item[field] = target.value;
    queueSave(entityType, rowId, {
      [itemPatchField(field)]: target.value,
    });
  });

  container?.addEventListener("click", (event) => {
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
      void archiveDefinition(entityType, rowId);
      return;
    }
    if (action === "move-up" || action === "move-down") {
      void moveDefinition(entityType, rowId, action === "move-up" ? -1 : 1);
    }
  });

  const addButton = document.getElementById(ENTITY_CONFIG[entityType].addId);
  addButton?.addEventListener("click", () => {
    void createDefinition(entityType);
  });
}

async function moveDefinition(entityType, rowId, offset) {
  try {
    await flushSaves();
    const list = state[entityType];
    const currentIndex = list.findIndex((item) => item.id === rowId);
    const targetIndex = currentIndex + offset;
    if (currentIndex < 0 || targetIndex < 0 || targetIndex >= list.length) {
      return;
    }

    const currentItem = list[currentIndex];
    const targetItem = list[targetIndex];
    list.splice(currentIndex, 1);
    list.splice(targetIndex, 0, currentItem);
    state[entityType] = list;
    resequenceItems(entityType);
    renderEntity(entityType);
    setSaveMessage("Saving order...", "saving");
    await Promise.all(
      state[entityType].map((item) =>
        directoryApi.updatePerformanceScorecardDefinition({
          entityType,
          id: item.id,
          patch: { sort_order: item.sortOrder },
        })
      )
    );
    setSaveMessage("Order saved.", "saved");
  } catch (error) {
    console.error("[scorecard-definitions] Failed to move definition", error);
    setSaveMessage(error?.message || "Could not update the order.", "error", true);
    await loadDefinitions();
  }
}

function attachListeners() {
  attachEntityListeners("values");
  attachEntityListeners("principles");
  attachEntityListeners("objectives");
  saveNowBtn?.addEventListener("click", () => {
    void flushSaves();
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
    if (!canAccessPage(role, "scorecarddefinitions")) {
      window.location.href = "./unauthorized.html?page=scorecard-definitions";
      return;
    }

    renderTopNavigation({ role });
    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
    await loadDefinitions();
  } catch (error) {
    console.error("[scorecard-definitions] Init failed", error);
    setStatus(error?.message || "Could not initialize scorecard setup.", true);
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
