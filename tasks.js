import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const searchInput = document.getElementById("searchInput");
const refreshBtn = document.getElementById("refreshBtn");
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const metaMessage = document.getElementById("metaMessage");
const tasksTableBody = document.getElementById("tasksTableBody");
const emptyState = document.getElementById("emptyState");

const detailRoot = document.getElementById("taskDetail");
const detailFields = {
  provider: detailRoot?.querySelector('[data-field="provider"]'),
  title: detailRoot?.querySelector('[data-field="title"]'),
  taskId: detailRoot?.querySelector('[data-field="taskId"]'),
  containerId: detailRoot?.querySelector('[data-field="containerId"]'),
  due: detailRoot?.querySelector('[data-field="due"]'),
  completed: detailRoot?.querySelector('[data-field="completed"]'),
};

const workingStatusSelect = document.getElementById("workingStatusSelect");
const workTypeInput = document.getElementById("workTypeInput");
const tagsInput = document.getElementById("tagsInput");
const overlayNotesInput = document.getElementById("overlayNotesInput");
const pinnedInput = document.getElementById("pinnedInput");
const startBtn = document.getElementById("startBtn");
const stopBtn = document.getElementById("stopBtn");
const saveOverlayBtn = document.getElementById("saveOverlayBtn");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});

const directoryApi = createDirectoryApi(authController);

let allTasks = [];
let selectedTaskKey = "";
let busy = false;

function taskKey(task) {
  return `${String(task?.provider || "").trim().toLowerCase()}|${String(task?.externalTaskId || "").trim()}`;
}

function setBusy(value) {
  busy = value;
  refreshBtn.disabled = value;
  saveOverlayBtn.disabled = value;
  startBtn.disabled = value;
  stopBtn.disabled = value;
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function logTaskError(context, error) {
  console.error(`[tasks-ui] ${context}`, {
    status: error?.status,
    code: error?.code,
    retryable: error?.retryable,
    correlationId: error?.correlationId,
    detail: error?.detail || error?.message || String(error),
  });
}

function formatDateTime(value) {
  if (!value) {
    return "-";
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return String(value);
  }
  return parsed.toLocaleString();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function compareNullableDateAsc(a, b) {
  if (!a && !b) {
    return 0;
  }
  if (!a) {
    return 1;
  }
  if (!b) {
    return -1;
  }

  const aTime = Date.parse(a);
  const bTime = Date.parse(b);
  if (!Number.isFinite(aTime) && !Number.isFinite(bTime)) {
    return 0;
  }
  if (!Number.isFinite(aTime)) {
    return 1;
  }
  if (!Number.isFinite(bTime)) {
    return -1;
  }

  return aTime - bTime;
}

function sortTasks(tasks) {
  return [...tasks].sort((a, b) => {
    const aPinned = a?.overlay?.pinned === true;
    const bPinned = b?.overlay?.pinned === true;
    if (aPinned !== bPinned) {
      return aPinned ? -1 : 1;
    }

    const aActive = String(a?.overlay?.workingStatus || "").toLowerCase() === "active";
    const bActive = String(b?.overlay?.workingStatus || "").toLowerCase() === "active";
    if (aActive !== bActive) {
      return aActive ? -1 : 1;
    }

    const dueCmp = compareNullableDateAsc(a?.dueDateTimeUtc, b?.dueDateTimeUtc);
    if (dueCmp !== 0) {
      return dueCmp;
    }

    const titleCmp = String(a?.title || "").localeCompare(String(b?.title || ""), undefined, {
      sensitivity: "base",
    });
    if (titleCmp !== 0) {
      return titleCmp;
    }

    return String(a?.externalTaskId || "").localeCompare(String(b?.externalTaskId || ""));
  });
}

function getSelectedTask() {
  return allTasks.find((task) => taskKey(task) === selectedTaskKey) || null;
}

function getFilteredTasks() {
  const query = String(searchInput?.value || "").trim().toLowerCase();
  if (!query) {
    return allTasks;
  }

  return allTasks.filter((task) => {
    return (
      String(task.title || "").toLowerCase().includes(query) ||
      String(task.provider || "").toLowerCase().includes(query) ||
      String(task.externalTaskId || "").toLowerCase().includes(query)
    );
  });
}

function setDetail(task) {
  if (!task) {
    detailFields.provider.textContent = "-";
    detailFields.title.textContent = "Select a task";
    detailFields.taskId.textContent = "-";
    detailFields.containerId.textContent = "-";
    detailFields.due.textContent = "-";
    detailFields.completed.textContent = "-";

    workingStatusSelect.value = "";
    workTypeInput.value = "";
    tagsInput.value = "";
    overlayNotesInput.value = "";
    pinnedInput.checked = false;
    return;
  }

  detailFields.provider.textContent = task.provider || "-";
  detailFields.title.textContent = task.title || "-";
  detailFields.taskId.textContent = task.externalTaskId || "-";
  detailFields.containerId.textContent = task.externalContainerId || "-";
  detailFields.due.textContent = formatDateTime(task.dueDateTimeUtc);
  detailFields.completed.textContent = task.isCompleted ? "Yes" : "No";

  const overlay = task.overlay || {};
  workingStatusSelect.value = overlay.workingStatus || "";
  workTypeInput.value = overlay.workType || "";
  tagsInput.value = Array.isArray(overlay.tags) ? overlay.tags.join(", ") : "";
  overlayNotesInput.value = overlay.overlayNotes || "";
  pinnedInput.checked = overlay.pinned === true;
}

function renderTasks() {
  const filtered = getFilteredTasks();
  tasksTableBody.innerHTML = "";

  if (!filtered.length) {
    emptyState.hidden = false;
    setDetail(null);
    return;
  }

  emptyState.hidden = true;

  const selected = filtered.find((task) => taskKey(task) === selectedTaskKey) || filtered[0];
  selectedTaskKey = taskKey(selected);

  for (const task of filtered) {
    const rowKey = taskKey(task);
    const tr = document.createElement("tr");
    tr.classList.toggle("selected", rowKey === selectedTaskKey);
    tr.dataset.taskKey = rowKey;

    const workingStatus = task?.overlay?.workingStatus || "-";
    const pinned = task?.overlay?.pinned === true ? "Yes" : "No";

    tr.innerHTML = `
      <td>${escapeHtml(task.provider || "-")}</td>
      <td>${escapeHtml(task.title || "-")}</td>
      <td>${escapeHtml(formatDateTime(task.dueDateTimeUtc))}</td>
      <td>${escapeHtml(task.isCompleted ? "Yes" : "No")}</td>
      <td>${escapeHtml(workingStatus)}</td>
      <td>${escapeHtml(pinned)}</td>
      <td>
        <div class="task-action-group">
          <button class="secondary task-action-btn" type="button" data-action="start" data-task-key="${escapeHtml(rowKey)}">Start</button>
          <button class="secondary task-action-btn" type="button" data-action="stop" data-task-key="${escapeHtml(rowKey)}">Stop</button>
          <button class="secondary task-action-btn" type="button" data-action="pin" data-task-key="${escapeHtml(rowKey)}">Pin</button>
        </div>
      </td>
    `;

    tr.addEventListener("click", (event) => {
      const actionTarget = event.target?.closest?.("button[data-action]");
      if (actionTarget) {
        return;
      }
      selectedTaskKey = rowKey;
      setDetail(task);
      renderTasks();
    });

    tasksTableBody.appendChild(tr);
  }

  setDetail(selected);
}

function updateTaskOverlayInMemory(provider, externalTaskId, overlay) {
  const key = `${String(provider || "").trim().toLowerCase()}|${String(externalTaskId || "").trim()}`;
  const task = allTasks.find((entry) => taskKey(entry) === key);
  if (!task) {
    return;
  }

  task.overlay = {
    itemId: overlay.itemId,
    workingStatus: overlay.workingStatus || "",
    workType: overlay.workType || "",
    tags: Array.isArray(overlay.tags) ? overlay.tags : [],
    activeStartedAt: overlay.activeStartedAt || null,
    lastWorkedAt: overlay.lastWorkedAt || null,
    energy: overlay.energy || "",
    effortMinutes: overlay.effortMinutes ?? null,
    impact: overlay.impact || "",
    overlayNotes: overlay.overlayNotes || "",
    pinned: overlay.pinned === true,
    lastOverlayUpdatedAt: overlay.lastOverlayUpdatedAt || null,
  };

  allTasks = sortTasks(allTasks);
}

async function submitOverlayPatch(provider, externalTaskId, patch) {
  setBusy(true);
  try {
    const key = `${String(provider || "").trim().toLowerCase()}|${String(externalTaskId || "").trim()}`;
    const task = allTasks.find((entry) => taskKey(entry) === key);
    const patchWithTitle = {
      ...(patch && typeof patch === "object" ? patch : {}),
      title: String(task?.title || "").trim(),
    };
    const result = await directoryApi.upsertTaskOverlay({ provider, externalTaskId, patch: patchWithTitle });
    updateTaskOverlayInMemory(provider, externalTaskId, result?.overlay || {});
    renderTasks();
    setStatus("Overlay updated.");
  } catch (error) {
    logTaskError("Overlay update failed", error);
    setStatus(error?.message || "Could not update overlay.", true);
  } finally {
    setBusy(false);
  }
}

async function refreshTasks() {
  setBusy(true);
  setStatus("Loading unified tasks...");

  try {
    const payload = await directoryApi.getUnifiedTasks();
    allTasks = sortTasks(Array.isArray(payload?.tasks) ? payload.tasks : []);

    const meta = payload?.meta || {};
    const providerErrors = meta.providerErrors && typeof meta.providerErrors === "object"
      ? meta.providerErrors
      : {};
    const providerErrorLabels = Object.keys(providerErrors);
    metaMessage.textContent =
      `Total: ${meta.total || 0} | To Do: ${meta.todoCount || 0} | ` +
      `Planner: ${meta.plannerCount || 0} | Overlay matched: ${meta.overlayMatchedCount || 0} | ` +
      `Overlay orphans: ${meta.overlayOrphanCount || 0}`;
    if (meta.partial && providerErrorLabels.length > 0) {
      metaMessage.textContent += ` | Partial: ${providerErrorLabels.join(", ")}`;
    }

    setStatus(`Loaded ${allTasks.length} task(s).`);
    renderTasks();
  } catch (error) {
    logTaskError("Unified task fetch failed", error);
    setStatus(error?.message || "Could not load unified tasks.", true);
    emptyState.hidden = false;
    tasksTableBody.innerHTML = "";
  } finally {
    setBusy(false);
  }
}

function getPatchFromDetailForm() {
  return {
    workingStatus: String(workingStatusSelect.value || "").trim().toLowerCase(),
    workType: String(workTypeInput.value || "").trim(),
    tags: String(tagsInput.value || "")
      .split(",")
      .map((entry) => entry.trim())
      .filter(Boolean),
    overlayNotes: String(overlayNotesInput.value || "").trim(),
    pinned: pinnedInput.checked,
  };
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
    if (!canAccessPage(role, "tasks")) {
      window.location.href = "./unauthorized.html?page=tasks";
      return;
    }

    renderTopNavigation({ role });
    await refreshTasks();
  } catch (error) {
    logTaskError("Tasks page init failed", error);
    setStatus(error?.message || "Could not initialize page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

tasksTableBody?.addEventListener("click", async (event) => {
  const button = event.target?.closest?.("button[data-action]");
  if (!button || busy) {
    return;
  }

  const rowKey = String(button.dataset.taskKey || "").trim();
  const action = String(button.dataset.action || "").trim();
  const task = allTasks.find((entry) => taskKey(entry) === rowKey);
  if (!task) {
    return;
  }

  selectedTaskKey = rowKey;
  renderTasks();

  if (action === "start") {
    await submitOverlayPatch(task.provider, task.externalTaskId, { workingStatus: "active" });
    return;
  }

  if (action === "stop") {
    await submitOverlayPatch(task.provider, task.externalTaskId, { workingStatus: "parked" });
    return;
  }

  if (action === "pin") {
    const currentlyPinned = task?.overlay?.pinned === true;
    await submitOverlayPatch(task.provider, task.externalTaskId, { pinned: !currentlyPinned });
  }
});

refreshBtn?.addEventListener("click", async () => {
  if (busy) {
    return;
  }
  await refreshTasks();
});

searchInput?.addEventListener("input", () => {
  renderTasks();
});

startBtn?.addEventListener("click", async () => {
  const task = getSelectedTask();
  if (!task || busy) {
    return;
  }
  await submitOverlayPatch(task.provider, task.externalTaskId, { workingStatus: "active" });
});

stopBtn?.addEventListener("click", async () => {
  const task = getSelectedTask();
  if (!task || busy) {
    return;
  }
  await submitOverlayPatch(task.provider, task.externalTaskId, { workingStatus: "parked" });
});

saveOverlayBtn?.addEventListener("click", async () => {
  const task = getSelectedTask();
  if (!task || busy) {
    return;
  }
  await submitOverlayPatch(task.provider, task.externalTaskId, getPatchFromDetailForm());
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
