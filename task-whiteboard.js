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
const stagingDropZone = document.getElementById("stagingDropZone");
const togglePinnedBtn = document.getElementById("togglePinnedBtn");
const stagingWrap = document.querySelector(".whiteboard-staging-wrap");

const DEFAULT_CARD_W = 220;
const DEFAULT_CARD_H = 72;
const SAVE_DEBOUNCE_MS = 320;
const GRID_SIZE = 24;

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
let stagedTaskCount = 0;
let pinnedMinimized = false;

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

function looksOpaqueTaskId(value) {
  const text = String(value || "").trim();
  if (!text) {
    return false;
  }
  if (/\s/.test(text)) {
    return false;
  }
  return text.length >= 28;
}

function displayTaskTitle(task) {
  const title = String(task?.title || "").trim();
  const externalTaskId = String(task?.externalTaskId || "").trim();
  if (!title) {
    return "Untitled task";
  }
  if (externalTaskId && title === externalTaskId) {
    return "Untitled task";
  }
  if (looksOpaqueTaskId(title)) {
    return "Untitled task";
  }
  return title;
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

function snapToGrid(value) {
  return Math.round(value / GRID_SIZE) * GRID_SIZE;
}

function normalizeCategoryLabel(value) {
  return String(value || "").trim();
}

function parseCategoryValue(raw) {
  const text = String(raw || "").trim();
  if (!text) {
    return { raw: "", label: "", box: null };
  }

  try {
    const parsed = JSON.parse(text);
    const label = normalizeCategoryLabel(parsed?.label || parsed?.name || parsed?.category);
    if (!label) {
      return { raw: text, label: "", box: null };
    }
    const box = parsed?.box && typeof parsed.box === "object" ? parsed.box : null;
    if (!box) {
      return { raw: text, label, box: null };
    }
    const x = Number(box.x);
    const y = Number(box.y);
    const w = Number(box.w);
    const h = Number(box.h);
    const color = String(box.color || "").trim();
    if (!Number.isFinite(x) || !Number.isFinite(y) || !Number.isFinite(w) || !Number.isFinite(h)) {
      return { raw: text, label, box: null };
    }
    return {
      raw: text,
      label,
      box: {
        x: snapToGrid(x),
        y: snapToGrid(y),
        w: Math.max(GRID_SIZE * 2, snapToGrid(w)),
        h: Math.max(GRID_SIZE * 2, snapToGrid(h)),
        color: color || "#4f74d9",
      },
    };
  } catch {
    return { raw: text, label: text, box: null };
  }
}

function buildCategoryValue(label, box) {
  const normalizedLabel = normalizeCategoryLabel(label);
  if (!normalizedLabel) {
    return "";
  }
  if (!box) {
    return normalizedLabel;
  }
  return JSON.stringify({
    label: normalizedLabel,
    box: {
      x: Number(box.x) || 0,
      y: Number(box.y) || 0,
      w: Number(box.w) || GRID_SIZE * 8,
      h: Number(box.h) || GRID_SIZE * 6,
      color: String(box.color || "#4f74d9"),
    },
  });
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

function updatePinnedPanelUi() {
  const collapsed = pinnedMinimized || stagedTaskCount === 0;
  stagingWrap?.classList.toggle("is-collapsed", collapsed);
  stagingWrap?.classList.toggle("is-empty", stagedTaskCount === 0);
  if (togglePinnedBtn) {
    togglePinnedBtn.textContent = pinnedMinimized ? "Expand" : "Minimize";
    togglePinnedBtn.setAttribute("aria-expanded", pinnedMinimized ? "false" : "true");
  }
  if (stagingDropZone) {
    stagingDropZone.textContent = stagedTaskCount === 0
      ? "Pinned staging is empty. Drag tasks here to unplace them."
      : "Drag here to return a task to pinned staging";
  }
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

function findCategoryBoxForPoint(x, y) {
  for (const box of categoryBoxes) {
    if (x >= box.x && x <= box.x + box.w && y >= box.y && y <= box.y + box.h) {
      return box;
    }
  }
  return null;
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

function persistCategoryForLabel(label, boxOrNull) {
  const normalizedLabel = normalizeCategoryLabel(label);
  if (!normalizedLabel) {
    return;
  }

  for (const [key, task] of tasksByKey.entries()) {
    const currentCategory = parseCategoryValue(task?.overlay?.category || "");
    if (normalizeCategoryLabel(currentCategory.label) !== normalizedLabel) {
      continue;
    }

    const nextCategory = boxOrNull ? buildCategoryValue(normalizedLabel, boxOrNull) : "";
    const rollbackSnapshot = {
      layout: String(task?.overlay?.layout || ""),
      category: String(task?.overlay?.category || ""),
    };

    updateTaskOverlayLocally(key, { category: nextCategory });
    schedulePersist(
      key,
      {
        title: String(task?.title || "").trim(),
        layout: String(task?.overlay?.layout || ""),
        category: nextCategory,
      },
      rollbackSnapshot
    );
  }
}

function hydrateCategoryBoxesFromTasks() {
  const byLabel = new Map();

  for (const task of tasksByKey.values()) {
    const parsedCategory = parseCategoryValue(task?.overlay?.category || "");
    if (!parsedCategory.label || !parsedCategory.box) {
      continue;
    }
    const key = normalizeCategoryLabel(parsedCategory.label).toLowerCase();
    if (byLabel.has(key)) {
      continue;
    }
    byLabel.set(key, {
      id: `box_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
      label: parsedCategory.label,
      x: parsedCategory.box.x,
      y: parsedCategory.box.y,
      w: parsedCategory.box.w,
      h: parsedCategory.box.h,
      color: parsedCategory.box.color || "#4f74d9",
    });
  }

  categoryBoxes = Array.from(byLabel.values());
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
      persistCategoryForLabel(box.label, null);
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

  stagedTaskCount = staged.length;
  updatePinnedPanelUi();

  staged.sort((a, b) => displayTaskTitle(a.task).localeCompare(displayTaskTitle(b.task), undefined, {
    sensitivity: "base",
  }));

  placed.sort((a, b) => {
    const zCmp = (a.layout.z || 0) - (b.layout.z || 0);
    if (zCmp !== 0) {
      return zCmp;
    }
    return displayTaskTitle(a.task).localeCompare(displayTaskTitle(b.task), undefined, {
      sensitivity: "base",
    });
  });

  stagingLane.innerHTML = "";
  for (const entry of staged) {
    const card = document.createElement("div");
    card.className = "whiteboard-task-card staged";
    card.draggable = true;
    card.dataset.taskKey = entry.key;
    const title = document.createElement("div");
    title.className = "whiteboard-task-title";
    title.textContent = displayTaskTitle(entry.task);
    title.title = title.textContent;

    const meta = document.createElement("div");
    meta.className = "whiteboard-task-meta";
    const parsedCategory = parseCategoryValue(entry.task?.overlay?.category || "");
    meta.textContent = `${formatDate(entry.task?.dueDateTimeUtc)}${
      parsedCategory.label ? ` • ${parsedCategory.label}` : ""
    }`;

    card.appendChild(title);
    card.appendChild(meta);
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
    const title = document.createElement("div");
    title.className = "whiteboard-task-title";
    title.textContent = displayTaskTitle(entry.task);
    title.title = title.textContent;

    const meta = document.createElement("div");
    meta.className = "whiteboard-task-meta";
    const parsedCategory = parseCategoryValue(entry.task?.overlay?.category || "");
    meta.textContent = `${formatDate(entry.task?.dueDateTimeUtc)}${
      parsedCategory.label ? ` • ${parsedCategory.label}` : ""
    }`;

    card.appendChild(title);
    card.appendChild(meta);
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
  const x = clamp(snapToGrid(boardX - cardW / 2), 0, Math.max(0, bounds.width - cardW));
  const y = clamp(snapToGrid(boardY - cardH / 2), 0, Math.max(0, bounds.height - cardH));
  const z = topZForCards() + 1;
  const nextLayout = { x, y, w: cardW, h: cardH, z };
  const categoryBox = findCategoryBoxForPoint(x + cardW / 2, y + cardH / 2);
  const category = categoryBox ? buildCategoryValue(categoryBox.label, categoryBox) : "";
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
      title: String(task?.title || "").trim(),
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
      title: String(task?.title || "").trim(),
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
  const x = snapToGrid(Math.min(drawDraft.startX, drawDraft.endX));
  const y = snapToGrid(Math.min(drawDraft.startY, drawDraft.endY));
  const w = Math.max(GRID_SIZE * 2, snapToGrid(Math.abs(drawDraft.endX - drawDraft.startX)));
  const h = Math.max(GRID_SIZE * 2, snapToGrid(Math.abs(drawDraft.endY - drawDraft.startY)));

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

  const created = categoryBoxes[categoryBoxes.length - 1];
  persistCategoryForLabel(created.label, created);
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
    box.x = clamp(snapToGrid(nextX), 0, Math.max(0, bounds.width - box.w));
    box.y = clamp(snapToGrid(nextY), 0, Math.max(0, bounds.height - box.h));
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
    const box = categoryBoxes.find((entry) => entry.id === movingBox.id) || null;
    if (box) {
      persistCategoryForLabel(box.label, box);
    }
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
    hydrateCategoryBoxesFromTasks();

    const meta = payload?.meta || {};
    const totalOverlayRows = Number(meta?.totalOverlayRows || 0);
    const requestedUserUpn = String(meta?.requestedUserUpn || "").trim().toLowerCase();
    const totalByProvider = meta?.totalByProvider && typeof meta.totalByProvider === "object"
      ? meta.totalByProvider
      : {};
    const pinnedByProvider = meta?.pinnedByProvider && typeof meta.pinnedByProvider === "object"
      ? meta.pinnedByProvider
      : {};
    const graphMatchedCount = Number(meta?.graphMatchedCount || 0);
    const graphMissCount = Number(meta?.graphMissCount || 0);
    const titleBackfilledCount = Number(meta?.titleBackfilledCount || 0);
    const partial = meta?.partial === true;
    setStatus(
      `TaskOverlay rows: ${totalOverlayRows} ` +
      `(planner ${totalByProvider.planner || 0}, todo ${totalByProvider.todo || 0}) | ` +
      `Pinned shown: ${tasksByKey.size} ` +
      `(planner ${pinnedByProvider.planner || 0}, todo ${pinnedByProvider.todo || 0}) | ` +
      `Graph matched: ${graphMatchedCount}, unmatched: ${graphMissCount}` +
      (titleBackfilledCount > 0 ? ` | title synced: ${titleBackfilledCount}` : "") +
      (partial ? " | partial provider data" : "") +
      `${requestedUserUpn ? ` | user ${requestedUserUpn}` : ""}`
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

function handleStagingDrop(event) {
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
}

stagingLane?.addEventListener("dragover", (event) => {
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
});
stagingLane?.addEventListener("drop", handleStagingDrop);
stagingDropZone?.addEventListener("dragover", (event) => {
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
});
stagingDropZone?.addEventListener("drop", handleStagingDrop);

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
  const labels = categoryBoxes.map((entry) => entry.label);
  categoryBoxes = [];
  for (const label of labels) {
    persistCategoryForLabel(label, null);
  }
  renderBoard();
});

togglePinnedBtn?.addEventListener("click", () => {
  pinnedMinimized = !pinnedMinimized;
  updatePinnedPanelUi();
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
updatePinnedPanelUi();
void init();
