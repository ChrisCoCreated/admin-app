import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const refreshBtn = document.getElementById("refreshBtn");
const drawBoxBtn = document.getElementById("drawBoxBtn");
const clearBoxesBtn = document.getElementById("clearBoxesBtn");
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const boardCanvas = document.getElementById("boardCanvas");
const stagingLane = document.getElementById("stagingLane");

const DEFAULT_CARD_W = 220;
const DEFAULT_CARD_H = 72;
const SAVE_DEBOUNCE_MS = 320;

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let busy = false;
let drawMode = false;
let tasksByKey = new Map();
let categoryBoxes = [];
let drawDraft = null;
let movingBox = null;

const persistQueue = new Map();
const persistSequenceByKey = new Map();
const lastPersistedByKey = new Map();

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setBusy(value) {
  busy = value;
  refreshBtn.disabled = value;
}

function taskKey(task) {
  return `${String(task?.provider || "").trim().toLowerCase()}|${String(task?.externalTaskId || "").trim()}`;
}

function formatDate(value) {
  if (!value) {
    return "No due date";
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return String(value);
  }
  return parsed.toLocaleDateString();
}

function isPinned(value) {
  if (value === true) {
    return true;
  }
  if (value === 1) {
    return true;
  }
  const text = String(value || "")
    .trim()
    .toLowerCase();
  return text === "true" || text === "1" || text === "yes";
}

function parseLayout(raw) {
  if (!raw || typeof raw !== "string") {
    return null;
  }

  try {
    const parsed = JSON.parse(raw);
    const x = Number(parsed?.x);
    const y = Number(parsed?.y);
    if (!Number.isFinite(x) || !Number.isFinite(y)) {
      return null;
    }

    const w = Number(parsed?.w);
    const h = Number(parsed?.h);
    const z = Number(parsed?.z);

    return {
      x,
      y,
      w: Number.isFinite(w) && w > 0 ? w : DEFAULT_CARD_W,
      h: Number.isFinite(h) && h > 0 ? h : DEFAULT_CARD_H,
      z: Number.isFinite(z) ? z : 1,
    };
  } catch {
    return null;
  }
}

function clamp(value, min, max) {
  return Math.max(min, Math.min(max, value));
}

function boardBounds() {
  const rect = boardCanvas.getBoundingClientRect();
  return {
    width: rect.width,
    height: rect.height,
  };
}

function updateDrawModeUi() {
  drawBoxBtn.classList.toggle("active", drawMode);
  drawBoxBtn.textContent = drawMode ? "Drawing... (Click board)" : "Draw Category Box";
}

function topZForCards() {
  let top = 1;
  for (const task of tasksByKey.values()) {
    const layout = parseLayout(task?.overlay?.layout || "");
    if (layout && layout.z > top) {
      top = layout.z;
    }
  }
  return top;
}

function findCategoryLabelForPoint(x, y) {
  for (const box of categoryBoxes) {
    if (x >= box.x && x <= box.x + box.w && y >= box.y && y <= box.y + box.h) {
      return box.label;
    }
  }
  return "";
}

function updateTaskOverlayLocally(key, patch) {
  const task = tasksByKey.get(key);
  if (!task) {
    return;
  }
  task.overlay = {
    ...(task.overlay || {}),
    ...patch,
  };
}

function schedulePersist(key, patch, rollbackSnapshot) {
  const previousSeq = persistSequenceByKey.get(key) || 0;
  const seq = previousSeq + 1;
  persistSequenceByKey.set(key, seq);

  const existing = persistQueue.get(key);
  if (existing?.timer) {
    clearTimeout(existing.timer);
  }

  const timer = setTimeout(async () => {
    const queued = persistQueue.get(key);
    if (!queued || queued.seq !== seq) {
      return;
    }

    const task = tasksByKey.get(key);
    if (!task) {
      persistQueue.delete(key);
      return;
    }

    try {
      const result = await directoryApi.upsertTaskOverlay({
        provider: task.provider,
        externalTaskId: task.externalTaskId,
        patch: queued.patch,
      });

      const latestSeq = persistSequenceByKey.get(key) || 0;
      if (seq !== latestSeq) {
        return;
      }

      const overlay = result?.overlay || {};
      updateTaskOverlayLocally(key, {
        layout: String(overlay.layout || ""),
        category: String(overlay.category || ""),
      });
      lastPersistedByKey.set(key, {
        layout: String(overlay.layout || ""),
        category: String(overlay.category || ""),
      });
      renderBoard();
    } catch (error) {
      const latestSeq = persistSequenceByKey.get(key) || 0;
      if (seq !== latestSeq) {
        return;
      }

      console.error("[whiteboard] Persist failed", {
        key,
        status: error?.status,
        code: error?.code,
        detail: error?.detail || error?.message || String(error),
      });

      updateTaskOverlayLocally(key, rollbackSnapshot);
      setStatus(error?.message || "Could not save whiteboard layout.", true);
      renderBoard();
    } finally {
      const active = persistQueue.get(key);
      if (active?.seq === seq) {
        persistQueue.delete(key);
      }
    }
  }, SAVE_DEBOUNCE_MS);

  persistQueue.set(key, {
    seq,
    patch,
    rollbackSnapshot,
    timer,
  });
}

function renderBoxes() {
  for (const box of categoryBoxes) {
    const boxEl = document.createElement("div");
    boxEl.className = "whiteboard-category-box";
    boxEl.style.left = `${box.x}px`;
    boxEl.style.top = `${box.y}px`;
    boxEl.style.width = `${box.w}px`;
    boxEl.style.height = `${box.h}px`;
    boxEl.style.borderColor = box.color;
    boxEl.dataset.boxId = box.id;

    const label = document.createElement("div");
    label.className = "whiteboard-category-label";
    label.textContent = box.label;
    label.style.backgroundColor = box.color;
    label.dataset.boxMoveHandle = "1";

    const deleteBtn = document.createElement("button");
    deleteBtn.type = "button";
    deleteBtn.className = "whiteboard-category-delete";
    deleteBtn.textContent = "×";
    deleteBtn.title = "Delete category box";
    deleteBtn.addEventListener("click", (event) => {
      event.stopPropagation();
      categoryBoxes = categoryBoxes.filter((entry) => entry.id !== box.id);
      renderBoard();
    });

    label.appendChild(deleteBtn);
    boxEl.appendChild(label);
    boardCanvas.appendChild(boxEl);
  }
}

function renderTaskCards() {
  const pinnedTasks = Array.from(tasksByKey.values()).filter((task) => isPinned(task?.overlay?.pinned));
  const staged = [];
  const placed = [];

  for (const task of pinnedTasks) {
    const key = taskKey(task);
    const layout = parseLayout(task?.overlay?.layout || "");
    if (layout) {
      placed.push({ task, key, layout });
    } else {
      staged.push({ task, key });
    }
  }

  staged.sort((a, b) => String(a.task?.title || "").localeCompare(String(b.task?.title || ""), undefined, {
    sensitivity: "base",
  }));

  placed.sort((a, b) => {
    const zCmp = (a.layout.z || 0) - (b.layout.z || 0);
    if (zCmp !== 0) {
      return zCmp;
    }
    return String(a.task?.title || "").localeCompare(String(b.task?.title || ""), undefined, {
      sensitivity: "base",
    });
  });

  stagingLane.innerHTML = "";
  for (const entry of staged) {
    const card = document.createElement("div");
    card.className = "whiteboard-task-card staged";
    card.draggable = true;
    card.dataset.taskKey = entry.key;
    card.innerHTML = `
      <div class="whiteboard-task-title">${entry.task?.title || "Untitled task"}</div>
      <div class="whiteboard-task-meta">${formatDate(entry.task?.dueDateTimeUtc)}${entry.task?.overlay?.category ? ` • ${entry.task.overlay.category}` : ""}</div>
    `;
    card.addEventListener("dragstart", onTaskDragStart);
    stagingLane.appendChild(card);
  }

  for (const entry of placed) {
    const card = document.createElement("div");
    card.className = "whiteboard-task-card";
    card.draggable = true;
    card.dataset.taskKey = entry.key;
    card.style.left = `${entry.layout.x}px`;
    card.style.top = `${entry.layout.y}px`;
    card.style.width = `${entry.layout.w}px`;
    card.style.height = `${entry.layout.h}px`;
    card.style.zIndex = String(entry.layout.z || 1);
    card.innerHTML = `
      <div class="whiteboard-task-title">${entry.task?.title || "Untitled task"}</div>
      <div class="whiteboard-task-meta">${formatDate(entry.task?.dueDateTimeUtc)}${entry.task?.overlay?.category ? ` • ${entry.task.overlay.category}` : ""}</div>
    `;
    card.addEventListener("dragstart", onTaskDragStart);
    boardCanvas.appendChild(card);
  }
}

function renderBoard() {
  boardCanvas.innerHTML = "";
  renderBoxes();
  renderTaskCards();
}

function onTaskDragStart(event) {
  const key = String(event.currentTarget?.dataset?.taskKey || "").trim();
  if (!key) {
    return;
  }
  event.dataTransfer.effectAllowed = "move";
  event.dataTransfer.setData("text/task-key", key);
}

function applyBoardDrop(task, key, boardX, boardY) {
  const bounds = boardBounds();
  const current = parseLayout(task?.overlay?.layout || "");
  const cardW = current?.w || DEFAULT_CARD_W;
  const cardH = current?.h || DEFAULT_CARD_H;
  const x = clamp(Math.round(boardX - cardW / 2), 0, Math.max(0, bounds.width - cardW));
  const y = clamp(Math.round(boardY - cardH / 2), 0, Math.max(0, bounds.height - cardH));
  const z = topZForCards() + 1;
  const nextLayout = { x, y, w: cardW, h: cardH, z };
  const category = findCategoryLabelForPoint(x + cardW / 2, y + cardH / 2);
  const layoutRaw = JSON.stringify(nextLayout);

  const rollbackSnapshot = {
    layout: String(task?.overlay?.layout || ""),
    category: String(task?.overlay?.category || ""),
  };

  updateTaskOverlayLocally(key, {
    layout: layoutRaw,
    category,
  });
  renderBoard();

  schedulePersist(
    key,
    {
      layout: layoutRaw,
      category,
    },
    rollbackSnapshot
  );
}

function applyStagingDrop(task, key) {
  const rollbackSnapshot = {
    layout: String(task?.overlay?.layout || ""),
    category: String(task?.overlay?.category || ""),
  };

  updateTaskOverlayLocally(key, {
    layout: "",
    category: "",
  });
  renderBoard();

  schedulePersist(
    key,
    {
      layout: "",
      category: "",
    },
    rollbackSnapshot
  );
}

function createBoxFromDraft() {
  if (!drawDraft) {
    return;
  }

  const minW = 80;
  const minH = 60;
  const x = Math.min(drawDraft.startX, drawDraft.endX);
  const y = Math.min(drawDraft.startY, drawDraft.endY);
  const w = Math.abs(drawDraft.endX - drawDraft.startX);
  const h = Math.abs(drawDraft.endY - drawDraft.startY);

  drawDraft = null;
  renderBoard();

  if (w < minW || h < minH) {
    return;
  }

  const label = window.prompt("Category label:", "");
  const normalized = String(label || "").trim();
  if (!normalized) {
    return;
  }

  const colorPalette = ["#4f74d9", "#24a38b", "#cb7a1a", "#9c4fd9", "#cc456a"];
  const color = colorPalette[categoryBoxes.length % colorPalette.length];

  categoryBoxes.push({
    id: `box_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
    label: normalized,
    x,
    y,
    w,
    h,
    color,
  });

  renderBoard();
}

function updateDrawDraftElement() {
  const existing = boardCanvas.querySelector(".whiteboard-draw-draft");
  if (!drawDraft) {
    if (existing) {
      existing.remove();
    }
    return;
  }

  const x = Math.min(drawDraft.startX, drawDraft.endX);
  const y = Math.min(drawDraft.startY, drawDraft.endY);
  const w = Math.abs(drawDraft.endX - drawDraft.startX);
  const h = Math.abs(drawDraft.endY - drawDraft.startY);

  const el = existing || document.createElement("div");
  el.className = "whiteboard-draw-draft";
  el.style.left = `${x}px`;
  el.style.top = `${y}px`;
  el.style.width = `${w}px`;
  el.style.height = `${h}px`;

  if (!existing) {
    boardCanvas.appendChild(el);
  }
}

function onBoardMouseDown(event) {
  const boardRect = boardCanvas.getBoundingClientRect();
  const x = event.clientX - boardRect.left;
  const y = event.clientY - boardRect.top;

  const boxHandle = event.target.closest("[data-box-move-handle]");
  if (!drawMode && boxHandle) {
    const boxElement = event.target.closest(".whiteboard-category-box");
    const boxId = boxElement?.dataset?.boxId;
    const box = categoryBoxes.find((entry) => entry.id === boxId);
    if (!box) {
      return;
    }

    movingBox = {
      id: box.id,
      startX: x,
      startY: y,
      originX: box.x,
      originY: box.y,
    };
    return;
  }

  if (!drawMode) {
    return;
  }

  if (event.target !== boardCanvas) {
    return;
  }

  drawDraft = {
    startX: x,
    startY: y,
    endX: x,
    endY: y,
  };
  updateDrawDraftElement();
}

function onBoardMouseMove(event) {
  const boardRect = boardCanvas.getBoundingClientRect();
  const x = event.clientX - boardRect.left;
  const y = event.clientY - boardRect.top;

  if (movingBox) {
    const box = categoryBoxes.find((entry) => entry.id === movingBox.id);
    if (!box) {
      movingBox = null;
      return;
    }

    const nextX = movingBox.originX + (x - movingBox.startX);
    const nextY = movingBox.originY + (y - movingBox.startY);
    const bounds = boardBounds();
    box.x = clamp(Math.round(nextX), 0, Math.max(0, bounds.width - box.w));
    box.y = clamp(Math.round(nextY), 0, Math.max(0, bounds.height - box.h));
    renderBoard();
    return;
  }

  if (!drawDraft) {
    return;
  }

  const bounds = boardBounds();
  drawDraft.endX = clamp(x, 0, bounds.width);
  drawDraft.endY = clamp(y, 0, bounds.height);
  updateDrawDraftElement();
}

function onBoardMouseUp() {
  if (movingBox) {
    movingBox = null;
    return;
  }
  if (!drawDraft) {
    return;
  }
  createBoxFromDraft();
}

async function refreshBoard() {
  setBusy(true);
  setStatus("Loading whiteboard tasks...");
  try {
    const payload = await directoryApi.getWhiteboardTasks();
    const tasks = Array.isArray(payload?.tasks) ? payload.tasks : [];

    tasksByKey = new Map();
    for (const task of tasks) {
      const key = taskKey(task);
      const overlay = task.overlay || {};
      tasksByKey.set(key, {
        ...task,
        overlay: {
          ...overlay,
          layout: String(overlay.layout || ""),
          category: String(overlay.category || ""),
        },
      });
      lastPersistedByKey.set(key, {
        layout: String(overlay.layout || ""),
        category: String(overlay.category || ""),
      });
    }

    const meta = payload?.meta || {};
    const totalOverlayRows = Number(meta?.totalOverlayRows || 0);
    const totalByProvider = meta?.totalByProvider && typeof meta.totalByProvider === "object"
      ? meta.totalByProvider
      : {};
    const pinnedByProvider = meta?.pinnedByProvider && typeof meta.pinnedByProvider === "object"
      ? meta.pinnedByProvider
      : {};
    setStatus(
      `TaskOverlay rows: ${totalOverlayRows} ` +
      `(planner ${totalByProvider.planner || 0}, todo ${totalByProvider.todo || 0}) | ` +
      `Pinned shown: ${tasksByKey.size} ` +
      `(planner ${pinnedByProvider.planner || 0}, todo ${pinnedByProvider.todo || 0})`
    );
    renderBoard();
  } catch (error) {
    console.error("[whiteboard] Refresh failed", error);
    setStatus(error?.message || "Could not load whiteboard tasks.", true);
    tasksByKey = new Map();
    renderBoard();
  } finally {
    setBusy(false);
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
    if (!canAccessPage(role, "whiteboard")) {
      window.location.href = "./unauthorized.html?page=whiteboard";
      return;
    }

    renderTopNavigation({ role });
    await refreshBoard();
  } catch (error) {
    console.error("[whiteboard] Init failed", error);
    setStatus(error?.message || "Could not initialize whiteboard.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

boardCanvas?.addEventListener("dragover", (event) => {
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
});

boardCanvas?.addEventListener("drop", (event) => {
  event.preventDefault();
  const key = String(event.dataTransfer.getData("text/task-key") || "").trim();
  if (!key) {
    return;
  }

  const task = tasksByKey.get(key);
  if (!task) {
    return;
  }

  const boardRect = boardCanvas.getBoundingClientRect();
  const x = event.clientX - boardRect.left;
  const y = event.clientY - boardRect.top;
  applyBoardDrop(task, key, x, y);
});

stagingLane?.addEventListener("dragover", (event) => {
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
});

stagingLane?.addEventListener("drop", (event) => {
  event.preventDefault();
  const key = String(event.dataTransfer.getData("text/task-key") || "").trim();
  if (!key) {
    return;
  }

  const task = tasksByKey.get(key);
  if (!task) {
    return;
  }

  applyStagingDrop(task, key);
});

boardCanvas?.addEventListener("mousedown", onBoardMouseDown);
window.addEventListener("mousemove", onBoardMouseMove);
window.addEventListener("mouseup", onBoardMouseUp);

refreshBtn?.addEventListener("click", async () => {
  if (busy) {
    return;
  }
  await refreshBoard();
});

drawBoxBtn?.addEventListener("click", () => {
  drawMode = !drawMode;
  drawDraft = null;
  updateDrawModeUi();
  renderBoard();
});

clearBoxesBtn?.addEventListener("click", () => {
  categoryBoxes = [];
  renderBoard();
});

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

updateDrawModeUi();
void init();
