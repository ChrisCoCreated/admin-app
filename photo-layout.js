import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const MAX_SELECTION = 6;
const PAN_LIMIT = 100;
const EXPORT_WIDTH = 2000;
const DEFAULT_GAP_PX = 24;
const DEFAULT_RADIUS_PX = 36;
const CLIENT_LIST_CACHE_KEY = "photoLayoutClientListV1";
const CLIENT_LIST_CACHE_TTL_MS = 6 * 60 * 60 * 1000;

const LAYOUTS = [
  {
    id: "single",
    name: "Single",
    aspect: 1.25,
    slots: [{ x: 0, y: 0, w: 1, h: 1 }],
  },
  {
    id: "split_two",
    name: "Split Two",
    aspect: 1.5,
    slots: [
      { x: 0, y: 0, w: 0.5, h: 1 },
      { x: 0.5, y: 0, w: 0.5, h: 1 },
    ],
  },
  {
    id: "feature_three",
    name: "Feature Three",
    aspect: 1.3333,
    slots: [
      { x: 0, y: 0, w: 1, h: 0.55 },
      { x: 0, y: 0.55, w: 0.5, h: 0.45 },
      { x: 0.5, y: 0.55, w: 0.5, h: 0.45 },
    ],
  },
  {
    id: "grid_four",
    name: "Grid Four",
    aspect: 1,
    slots: [
      { x: 0, y: 0, w: 0.5, h: 0.5 },
      { x: 0.5, y: 0, w: 0.5, h: 0.5 },
      { x: 0, y: 0.5, w: 0.5, h: 0.5 },
      { x: 0.5, y: 0.5, w: 0.5, h: 0.5 },
    ],
  },
  {
    id: "mosaic_five",
    name: "Mosaic Five",
    aspect: 1.3333,
    slots: [
      { x: 0, y: 0, w: 0.6, h: 0.62 },
      { x: 0.6, y: 0, w: 0.4, h: 0.31 },
      { x: 0.6, y: 0.31, w: 0.4, h: 0.31 },
      { x: 0, y: 0.62, w: 0.5, h: 0.38 },
      { x: 0.5, y: 0.62, w: 0.5, h: 0.38 },
    ],
  },
  {
    id: "grid_six",
    name: "Grid Six",
    aspect: 1.5,
    slots: [
      { x: 0, y: 0, w: 1 / 3, h: 0.5 },
      { x: 1 / 3, y: 0, w: 1 / 3, h: 0.5 },
      { x: 2 / 3, y: 0, w: 1 / 3, h: 0.5 },
      { x: 0, y: 0.5, w: 1 / 3, h: 0.5 },
      { x: 1 / 3, y: 0.5, w: 1 / 3, h: 0.5 },
      { x: 2 / 3, y: 0.5, w: 1 / 3, h: 0.5 },
    ],
  },
];

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const clientSelect = document.getElementById("clientSelect");
const imagesStatus = document.getElementById("imagesStatus");
const imagesGrid = document.getElementById("imagesGrid");
const layoutPicker = document.getElementById("layoutPicker");
const composeStatus = document.getElementById("composeStatus");
const layoutStage = document.getElementById("layoutStage");
const selectedImagesList = document.getElementById("selectedImagesList");
const adjustStatus = document.getElementById("adjustStatus");
const adjustControls = document.getElementById("adjustControls");
const zoomRange = document.getElementById("zoomRange");
const panXRange = document.getElementById("panXRange");
const panYRange = document.getElementById("panYRange");
const resetAdjustBtn = document.getElementById("resetAdjustBtn");
const gapEnabledInput = document.getElementById("gapEnabled");
const gapRange = document.getElementById("gapRange");
const gapValue = document.getElementById("gapValue");
const roundedEnabledInput = document.getElementById("roundedEnabled");
const cornerRadiusRange = document.getElementById("cornerRadiusRange");
const cornerRadiusValue = document.getElementById("cornerRadiusValue");
const generateOutputBtn = document.getElementById("generateOutputBtn");
const copyOutputBtn = document.getElementById("copyOutputBtn");
const saveOutputBtn = document.getElementById("saveOutputBtn");
const exportStatus = document.getElementById("exportStatus");
const outputPreviewImage = document.getElementById("outputPreviewImage");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let allPhotos = [];
let clientPhotoPool = [];
let cachedClients = [];
let selectedClient = "";
let selectedImages = [];
let selectedSlotIndex = -1;
let activeLayoutId = LAYOUTS[0].id;
let dragState = null;
let latestOutputBlob = null;
let latestOutputUrl = "";
const layoutStyle = {
  gapEnabled: true,
  gapPx: DEFAULT_GAP_PX,
  roundedEnabled: true,
  cornerRadiusPx: DEFAULT_RADIUS_PX,
};

const exportImageCache = new Map();
const clientPhotoCache = new Map();

function loadCachedClientList() {
  try {
    const raw =
      localStorage.getItem(CLIENT_LIST_CACHE_KEY) || sessionStorage.getItem(CLIENT_LIST_CACHE_KEY);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    const clients = Array.isArray(parsed) ? parsed : parsed?.clients;
    const cachedAt = Number(parsed?.cachedAt || 0);
    if (!Array.isArray(clients)) {
      return [];
    }
    if (cachedAt && Date.now() - cachedAt > CLIENT_LIST_CACHE_TTL_MS) {
      return [];
    }
    return clients
      .map((value) => normalizeClientName(value))
      .filter((value, index, array) => Boolean(value) && array.indexOf(value) === index);
  } catch {
    return [];
  }
}

function saveCachedClientList(clients) {
  const payload = {
    clients: Array.isArray(clients) ? clients : [],
    cachedAt: Date.now(),
  };
  try {
    localStorage.setItem(CLIENT_LIST_CACHE_KEY, JSON.stringify(payload));
  } catch {
    // ignore local cache errors
  }
  try {
    sessionStorage.setItem(CLIENT_LIST_CACHE_KEY, JSON.stringify(payload));
  } catch {
    // ignore session cache errors
  }
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setImagesStatus(message, isError = false) {
  imagesStatus.textContent = message;
  imagesStatus.classList.toggle("error", isError);
}

function setComposeStatus(message, isError = false) {
  composeStatus.textContent = message;
  composeStatus.classList.toggle("error", isError);
}

function setExportStatus(message, isError = false) {
  exportStatus.textContent = message;
  exportStatus.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "photolayout").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function normalizeClientName(value) {
  return String(value || "").trim() || "Unassigned";
}

function isImagePhoto(photo) {
  const type = String(photo?.mediaType || "").toLowerCase();
  if (type === "video") {
    return false;
  }
  const url = String(photo?.imageUrl || photo?.mediaUrl || "").toLowerCase();
  if (!url) {
    return false;
  }
  if (/\.(mp4|mov|webm)(?:$|[?#])/.test(url)) {
    return false;
  }
  return true;
}

function getActiveLayout() {
  return LAYOUTS.find((layout) => layout.id === activeLayoutId) || LAYOUTS[0];
}

function getClientPhotos() {
  return clientPhotoPool;
}

function updateClientOptions() {
  const current = selectedClient;
  const clients = cachedClients.slice().sort((a, b) => a.localeCompare(b));
  clientSelect.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Select client";
  clientSelect.append(placeholder);
  if (!clients.length) {
    placeholder.textContent = "No clients with images";
    clientSelect.value = "";
    selectedClient = "";
    return;
  }
  for (const client of clients) {
    const option = document.createElement("option");
    option.value = client;
    option.textContent = client;
    clientSelect.append(option);
  }
  if (current && clients.includes(current)) {
    clientSelect.value = current;
    selectedClient = current;
  } else {
    selectedClient = "";
    clientSelect.value = "";
  }
}

function removeImageFromSelection(photoId) {
  const index = selectedImages.findIndex((item) => item.id === photoId);
  if (index < 0) {
    return;
  }
  selectedImages.splice(index, 1);
  if (selectedSlotIndex >= selectedImages.length) {
    selectedSlotIndex = selectedImages.length - 1;
  }
  invalidateOutput();
}

function toggleImageSelection(photo) {
  const existing = selectedImages.find((item) => item.id === photo.id);
  if (existing) {
    removeImageFromSelection(photo.id);
    renderAll();
    return;
  }
  if (selectedImages.length >= MAX_SELECTION) {
    setImagesStatus(`You can select up to ${MAX_SELECTION} images.`, true);
    return;
  }
  selectedImages.push({
    id: photo.id,
    title: photo.title || photo.client || "Untitled",
    client: normalizeClientName(photo.client),
    sourceCandidates: [photo.attachmentUrl, photo.mediaUrl, photo.imageUrl]
      .map((value) => String(value || "").trim())
      .filter((value, index, array) => Boolean(value) && array.indexOf(value) === index),
    previewUrl: photo.imageUrl || photo.mediaUrl || "",
    zoom: 1,
    panX: 0,
    panY: 0,
  });
  if (selectedSlotIndex < 0) {
    selectedSlotIndex = 0;
  }
  invalidateOutput();
  renderAll();
}

function renderImagesGrid() {
  imagesGrid.innerHTML = "";
  if (!selectedClient) {
    setImagesStatus("Select a client to load images.");
    return;
  }
  const photos = getClientPhotos();
  if (!photos.length) {
    setImagesStatus("No images found for this client.");
    return;
  }
  const selectedSet = new Set(selectedImages.map((item) => item.id));
  for (const photo of photos) {
    const selected = selectedSet.has(photo.id);
    const card = document.createElement("button");
    card.type = "button";
    card.className = `layout-image-card${selected ? " selected" : ""}`;
    card.disabled = !selected && selectedImages.length >= MAX_SELECTION;
    card.setAttribute("aria-pressed", selected ? "true" : "false");

    const img = document.createElement("img");
    img.className = "layout-image-thumb";
    img.src = photo.imageUrl || photo.mediaUrl;
    img.alt = photo.title || photo.client || "Client image";
    img.loading = "lazy";

    const caption = document.createElement("span");
    caption.className = "layout-image-caption";
    caption.textContent = photo.title || photo.fileName || "Untitled image";

    card.append(img, caption);
    if (selected) {
      const badge = document.createElement("span");
      badge.className = "layout-selected-badge";
      badge.textContent = "Selected";
      card.append(badge);
    }
    card.addEventListener("click", () => toggleImageSelection(photo));
    imagesGrid.append(card);
  }
  setImagesStatus(
    `${photos.length} image${photos.length === 1 ? "" : "s"} for ${selectedClient}. Selected ${
      selectedImages.length
    }/${MAX_SELECTION}.`
  );
}

function renderLayoutPicker() {
  layoutPicker.innerHTML = "";
  for (const layout of LAYOUTS) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `layout-thumb${layout.id === activeLayoutId ? " active" : ""}`;
    btn.setAttribute("aria-pressed", layout.id === activeLayoutId ? "true" : "false");
    btn.addEventListener("click", () => {
      activeLayoutId = layout.id;
      selectedSlotIndex = Math.min(selectedSlotIndex, getActiveLayout().slots.length - 1);
      invalidateOutput();
      renderAll();
    });

    const mini = document.createElement("div");
    mini.className = "layout-thumb-canvas";
    mini.style.setProperty("--layout-thumb-aspect", String(layout.aspect));
    for (const slot of layout.slots) {
      const cell = document.createElement("span");
      cell.className = "layout-thumb-slot";
      cell.style.left = `${slot.x * 100}%`;
      cell.style.top = `${slot.y * 100}%`;
      cell.style.width = `${slot.w * 100}%`;
      cell.style.height = `${slot.h * 100}%`;
      mini.append(cell);
    }

    const label = document.createElement("span");
    label.className = "layout-thumb-label";
    label.textContent = layout.name;
    btn.append(mini, label);
    layoutPicker.append(btn);
  }
}

function clamp(value, min, max) {
  return Math.min(Math.max(value, min), max);
}

function rangesOverlap(startA, endA, startB, endB) {
  return Math.min(endA, endB) - Math.max(startA, startB) > 1e-6;
}

function nearlyEqual(valueA, valueB) {
  return Math.abs(valueA - valueB) <= 1e-6;
}

function getSlotNeighborFlags(slots, index) {
  const slot = slots[index];
  const slotLeft = slot.x;
  const slotRight = slot.x + slot.w;
  const slotTop = slot.y;
  const slotBottom = slot.y + slot.h;
  const flags = {
    hasLeftNeighbor: false,
    hasRightNeighbor: false,
    hasTopNeighbor: false,
    hasBottomNeighbor: false,
  };

  for (let otherIndex = 0; otherIndex < slots.length; otherIndex += 1) {
    if (otherIndex === index) {
      continue;
    }
    const other = slots[otherIndex];
    const otherLeft = other.x;
    const otherRight = other.x + other.w;
    const otherTop = other.y;
    const otherBottom = other.y + other.h;

    if (nearlyEqual(slotLeft, otherRight) && rangesOverlap(slotTop, slotBottom, otherTop, otherBottom)) {
      flags.hasLeftNeighbor = true;
    }
    if (nearlyEqual(slotRight, otherLeft) && rangesOverlap(slotTop, slotBottom, otherTop, otherBottom)) {
      flags.hasRightNeighbor = true;
    }
    if (nearlyEqual(slotTop, otherBottom) && rangesOverlap(slotLeft, slotRight, otherLeft, otherRight)) {
      flags.hasTopNeighbor = true;
    }
    if (nearlyEqual(slotBottom, otherTop) && rangesOverlap(slotLeft, slotRight, otherLeft, otherRight)) {
      flags.hasBottomNeighbor = true;
    }
  }

  return flags;
}

function computeAdjustedSlotRect(slot, neighborFlags, width, height, gapPx) {
  const halfGap = Math.max(0, Number(gapPx) || 0) * 0.5;
  let x = slot.x * width;
  let y = slot.y * height;
  let w = slot.w * width;
  let h = slot.h * height;

  if (neighborFlags.hasLeftNeighbor) {
    x += halfGap;
    w -= halfGap;
  }
  if (neighborFlags.hasRightNeighbor) {
    w -= halfGap;
  }
  if (neighborFlags.hasTopNeighbor) {
    y += halfGap;
    h -= halfGap;
  }
  if (neighborFlags.hasBottomNeighbor) {
    h -= halfGap;
  }

  return {
    x,
    y,
    w: Math.max(0, w),
    h: Math.max(0, h),
  };
}

function getExportGapPx() {
  return layoutStyle.gapEnabled ? clamp(layoutStyle.gapPx, 0, 120) : 0;
}

function getExportCornerRadiusPx() {
  return layoutStyle.roundedEnabled ? clamp(layoutStyle.cornerRadiusPx, 0, 240) : 0;
}

function updateStyleControls() {
  if (gapEnabledInput) {
    gapEnabledInput.checked = layoutStyle.gapEnabled;
  }
  if (roundedEnabledInput) {
    roundedEnabledInput.checked = layoutStyle.roundedEnabled;
  }
  if (gapRange) {
    gapRange.value = String(layoutStyle.gapPx);
    gapRange.disabled = !layoutStyle.gapEnabled;
  }
  if (cornerRadiusRange) {
    cornerRadiusRange.value = String(layoutStyle.cornerRadiusPx);
    cornerRadiusRange.disabled = !layoutStyle.roundedEnabled;
  }
  if (gapValue) {
    gapValue.textContent = `${Math.round(layoutStyle.gapPx)}px`;
  }
  if (cornerRadiusValue) {
    cornerRadiusValue.textContent = `${Math.round(layoutStyle.cornerRadiusPx)}px`;
  }
}

function invalidateOutput() {
  latestOutputBlob = null;
  if (latestOutputUrl) {
    URL.revokeObjectURL(latestOutputUrl);
    latestOutputUrl = "";
  }
  outputPreviewImage.removeAttribute("src");
  outputPreviewImage.hidden = true;
  setExportStatus("Generate an output image to copy or save.");
}

function swapSelectedImages(fromIndex, toIndex) {
  if (fromIndex < 0 || toIndex < 0 || fromIndex >= selectedImages.length || toIndex >= selectedImages.length) {
    return;
  }
  const temp = selectedImages[fromIndex];
  selectedImages[fromIndex] = selectedImages[toIndex];
  selectedImages[toIndex] = temp;
  selectedSlotIndex = toIndex;
  invalidateOutput();
  renderAll();
}

function renderSelectedImagesList() {
  selectedImagesList.innerHTML = "";
  if (!selectedImages.length) {
    selectedImagesList.innerHTML = '<p class="muted">No images selected yet.</p>';
    return;
  }
  selectedImages.forEach((item, index) => {
    const row = document.createElement("div");
    row.className = `selected-image-row${index === selectedSlotIndex ? " active" : ""}`;

    const info = document.createElement("button");
    info.type = "button";
    info.className = "selected-image-info";
    info.textContent = `${index + 1}. ${item.title}`;
    info.addEventListener("click", () => {
      selectedSlotIndex = index;
      updateAdjustControls();
      renderStage();
      renderSelectedImagesList();
    });

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "secondary selected-image-remove";
    removeBtn.textContent = "Remove";
    removeBtn.addEventListener("click", () => {
      removeImageFromSelection(item.id);
      renderAll();
    });

    const moveUpBtn = document.createElement("button");
    moveUpBtn.type = "button";
    moveUpBtn.className = "secondary selected-image-move";
    moveUpBtn.textContent = "Up";
    moveUpBtn.disabled = index === 0;
    moveUpBtn.addEventListener("click", () => {
      swapSelectedImages(index, index - 1);
    });

    const moveDownBtn = document.createElement("button");
    moveDownBtn.type = "button";
    moveDownBtn.className = "secondary selected-image-move";
    moveDownBtn.textContent = "Down";
    moveDownBtn.disabled = index === selectedImages.length - 1;
    moveDownBtn.addEventListener("click", () => {
      swapSelectedImages(index, index + 1);
    });

    row.append(info, moveUpBtn, moveDownBtn, removeBtn);
    selectedImagesList.append(row);
  });
}

function renderStage() {
  const layout = getActiveLayout();
  layoutStage.innerHTML = "";
  layoutStage.style.setProperty("--layout-stage-aspect", String(layout.aspect));
  const stageWidth = layoutStage.clientWidth || layoutStage.getBoundingClientRect().width || 1;
  const stageHeight = layoutStage.clientHeight || stageWidth / layout.aspect || 1;
  const previewScale = stageWidth / EXPORT_WIDTH;
  const previewGapPx = getExportGapPx() * previewScale;
  const previewCornerRadiusPx = getExportCornerRadiusPx() * previewScale;

  layout.slots.forEach((slot, slotIndex) => {
    const slotEl = document.createElement("div");
    const assignedImage = selectedImages[slotIndex];
    const neighborFlags = getSlotNeighborFlags(layout.slots, slotIndex);
    const slotRect = computeAdjustedSlotRect(slot, neighborFlags, stageWidth, stageHeight, previewGapPx);
    slotEl.className = `layout-slot${slotIndex === selectedSlotIndex ? " active" : ""}${assignedImage ? " has-image" : ""}`;
    slotEl.style.left = `${slotRect.x}px`;
    slotEl.style.top = `${slotRect.y}px`;
    slotEl.style.width = `${slotRect.w}px`;
    slotEl.style.height = `${slotRect.h}px`;
    slotEl.style.borderRadius = `${previewCornerRadiusPx}px`;
    slotEl.dataset.slotIndex = String(slotIndex);
    slotEl.addEventListener("click", () => {
      selectedSlotIndex = slotIndex;
      updateAdjustControls();
      renderStage();
      renderSelectedImagesList();
    });

    if (assignedImage) {
      const media = document.createElement("img");
      media.className = "layout-slot-image";
      media.src = assignedImage.previewUrl;
      media.alt = assignedImage.title;
      media.draggable = false;
      media.style.objectPosition = `${50 - assignedImage.panX * 0.5}% ${50 - assignedImage.panY * 0.5}%`;
      media.style.transform = `scale(${assignedImage.zoom})`;

      media.addEventListener("pointerdown", (event) => {
        selectedSlotIndex = slotIndex;
        const rect = slotEl.getBoundingClientRect();
        dragState = {
          slotIndex,
          pointerId: event.pointerId,
          startX: event.clientX,
          startY: event.clientY,
          startPanX: assignedImage.panX,
          startPanY: assignedImage.panY,
          width: rect.width || 1,
          height: rect.height || 1,
        };
        media.setPointerCapture(event.pointerId);
      });
      media.addEventListener("pointermove", (event) => {
        if (!dragState || dragState.pointerId !== event.pointerId || dragState.slotIndex !== slotIndex) {
          return;
        }
        const image = selectedImages[slotIndex];
        if (!image) {
          return;
        }
        const deltaX = ((event.clientX - dragState.startX) / dragState.width) * 100;
        const deltaY = ((event.clientY - dragState.startY) / dragState.height) * 100;
        image.panX = clamp(dragState.startPanX + deltaX, -PAN_LIMIT, PAN_LIMIT);
        image.panY = clamp(dragState.startPanY + deltaY, -PAN_LIMIT, PAN_LIMIT);
        media.style.objectPosition = `${50 - image.panX * 0.5}% ${50 - image.panY * 0.5}%`;
        media.style.transform = `scale(${image.zoom})`;
        if (slotIndex === selectedSlotIndex) {
          updateAdjustControls();
        }
      });
      const clearDrag = () => {
        dragState = null;
      };
      media.addEventListener("pointerup", clearDrag);
      media.addEventListener("pointercancel", clearDrag);

      slotEl.append(media);
    } else {
      const placeholder = document.createElement("span");
      placeholder.className = "layout-slot-placeholder";
      placeholder.textContent = `Slot ${slotIndex + 1}`;
      slotEl.append(placeholder);
    }

    layoutStage.append(slotEl);
  });

  const usedCount = Math.min(layout.slots.length, selectedImages.length);
  setComposeStatus(
    `${layout.name} layout: using ${usedCount} of ${selectedImages.length} selected image${
      selectedImages.length === 1 ? "" : "s"
    }. Drag inside an image to reposition; use Up/Down to reorder.`
  );
}

function updateAdjustControls() {
  const image = selectedImages[selectedSlotIndex];
  if (!image) {
    adjustControls.hidden = true;
    adjustStatus.textContent = "Select a filled slot to adjust zoom and position.";
    return;
  }
  adjustControls.hidden = false;
  adjustStatus.textContent = `Adjusting: ${image.title}`;
  zoomRange.value = String(image.zoom);
  panXRange.value = String(Math.round(image.panX));
  panYRange.value = String(Math.round(image.panY));
}

function applyAdjustmentsFromControls() {
  const image = selectedImages[selectedSlotIndex];
  if (!image) {
    return;
  }
  image.zoom = clamp(Number(zoomRange.value) || 1, 1, 3);
  image.panX = clamp(Number(panXRange.value) || 0, -PAN_LIMIT, PAN_LIMIT);
  image.panY = clamp(Number(panYRange.value) || 0, -PAN_LIMIT, PAN_LIMIT);
  invalidateOutput();
  renderStage();
}

function drawRoundedRectPath(ctx, x, y, w, h, radius) {
  const maxRadius = Math.min(w, h) * 0.5;
  const r = clamp(Number(radius) || 0, 0, maxRadius);
  if (!r) {
    ctx.rect(x, y, w, h);
    return;
  }
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.arcTo(x + w, y, x + w, y + r, r);
  ctx.lineTo(x + w, y + h - r);
  ctx.arcTo(x + w, y + h, x + w - r, y + h, r);
  ctx.lineTo(x + r, y + h);
  ctx.arcTo(x, y + h, x, y + h - r, r);
  ctx.lineTo(x, y + r);
  ctx.arcTo(x, y, x + r, y, r);
}

function drawImageIntoSlot(ctx, img, slot, transform, cornerRadiusPx = 0) {
  const slotX = slot.x;
  const slotY = slot.y;
  const slotW = slot.w;
  const slotH = slot.h;
  const zoom = clamp(Number(transform.zoom) || 1, 1, 3);
  const panX = clamp(Number(transform.panX) || 0, -PAN_LIMIT, PAN_LIMIT);
  const panY = clamp(Number(transform.panY) || 0, -PAN_LIMIT, PAN_LIMIT);

  const baseScale = Math.max(slotW / img.width, slotH / img.height);
  const drawW = img.width * baseScale * zoom;
  const drawH = img.height * baseScale * zoom;
  const overflowX = Math.max(0, drawW - slotW);
  const overflowY = Math.max(0, drawH - slotH);
  const drawX = slotX + (slotW - drawW) / 2 + (overflowX / 2) * (panX / PAN_LIMIT);
  const drawY = slotY + (slotH - drawH) / 2 + (overflowY / 2) * (panY / PAN_LIMIT);

  ctx.save();
  ctx.beginPath();
  drawRoundedRectPath(ctx, slotX, slotY, slotW, slotH, cornerRadiusPx);
  ctx.clip();
  ctx.drawImage(img, drawX, drawY, drawW, drawH);
  ctx.restore();
}

function decodeBase64ToBlob(base64, mimeType) {
  const binary = atob(String(base64 || ""));
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return new Blob([bytes], { type: mimeType || "image/jpeg" });
}

function loadImageElement(url) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error(`Could not load image from ${url}`));
    img.src = url;
  });
}

async function loadExportImage(sourceCandidates = []) {
  const candidates = Array.isArray(sourceCandidates)
    ? sourceCandidates.map((value) => String(value || "").trim()).filter(Boolean)
    : [];
  if (!candidates.length) {
    throw new Error("Missing source image URL.");
  }

  let lastError = null;
  for (const candidate of candidates) {
    if (exportImageCache.has(candidate)) {
      return exportImageCache.get(candidate);
    }
    try {
      const payload = await directoryApi.getMarketingMedia({ url: candidate });
      const blob = decodeBase64ToBlob(payload?.dataBase64, payload?.mimeType);
      const objectUrl = URL.createObjectURL(blob);
      try {
        const image = await loadImageElement(objectUrl);
        exportImageCache.set(candidate, image);
        return image;
      } finally {
        URL.revokeObjectURL(objectUrl);
      }
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("Could not load export image.");
}

async function buildOutputCanvas() {
  const layout = getActiveLayout();
  const outputHeight = Math.round(EXPORT_WIDTH / layout.aspect);
  const canvas = document.createElement("canvas");
  canvas.width = EXPORT_WIDTH;
  canvas.height = outputHeight;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    throw new Error("Could not initialize output canvas.");
  }

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  const renderableSlots = Math.min(layout.slots.length, selectedImages.length);
  const gapPx = getExportGapPx();
  const cornerRadiusPx = getExportCornerRadiusPx();
  for (let index = 0; index < renderableSlots; index += 1) {
    const imageState = selectedImages[index];
    const sourceCandidates = Array.isArray(imageState?.sourceCandidates) ? imageState.sourceCandidates : [];
    if (!sourceCandidates.length) {
      continue;
    }
    const img = await loadExportImage(sourceCandidates);
    const slot = layout.slots[index];
    const neighborFlags = getSlotNeighborFlags(layout.slots, index);
    const slotRect = computeAdjustedSlotRect(slot, neighborFlags, canvas.width, canvas.height, gapPx);
    drawImageIntoSlot(
      ctx,
      img,
      {
        x: slotRect.x,
        y: slotRect.y,
        w: slotRect.w,
        h: slotRect.h,
      },
      imageState,
      cornerRadiusPx
    );
  }

  return canvas;
}

function canvasToBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error("Could not encode output image."));
        return;
      }
      resolve(blob);
    }, "image/png");
  });
}

async function generateOutput() {
  if (!selectedImages.length) {
    setExportStatus("Select at least one image first.", true);
    return null;
  }
  generateOutputBtn.disabled = true;
  setExportStatus("Generating output image...");
  try {
    const canvas = await buildOutputCanvas();
    const blob = await canvasToBlob(canvas);
    latestOutputBlob = blob;
    if (latestOutputUrl) {
      URL.revokeObjectURL(latestOutputUrl);
    }
    latestOutputUrl = URL.createObjectURL(blob);
    outputPreviewImage.src = latestOutputUrl;
    outputPreviewImage.hidden = false;
    setExportStatus(`Output ready (${EXPORT_WIDTH}px wide PNG).`);
    return blob;
  } catch (error) {
    console.error(error);
    setExportStatus(error?.message || "Could not generate output image.", true);
    return null;
  } finally {
    generateOutputBtn.disabled = false;
  }
}

async function copyOutput() {
  let blob = latestOutputBlob;
  if (!blob) {
    blob = await generateOutput();
  }
  if (!blob) {
    return;
  }
  if (!(navigator.clipboard && window.ClipboardItem)) {
    setExportStatus("Clipboard image copy is not supported in this browser. Use Save PNG.", true);
    return;
  }
  try {
    await navigator.clipboard.write([new ClipboardItem({ [blob.type]: blob })]);
    setExportStatus("Image copied to clipboard.");
  } catch (error) {
    console.error(error);
    setExportStatus("Could not copy to clipboard. Use Save PNG.", true);
  }
}

async function saveOutput() {
  let blob = latestOutputBlob;
  if (!blob) {
    blob = await generateOutput();
  }
  if (!blob) {
    return;
  }
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  const clientToken = normalizeClientName(selectedClient).replace(/[^a-z0-9]+/gi, "-").replace(/^-+|-+$/g, "");
  const layoutToken = getActiveLayout().id.replace(/[^a-z0-9]+/gi, "-");
  link.href = url;
  link.download = `${clientToken || "client"}-${layoutToken || "layout"}-2000w.png`;
  link.click();
  URL.revokeObjectURL(url);
  setExportStatus("PNG download started.");
}

function renderAll() {
  renderImagesGrid();
  renderLayoutPicker();
  renderSelectedImagesList();
  renderStage();
  updateAdjustControls();
  updateStyleControls();
}

async function loadPhotos() {
  const localLoadStartedAt = performance.now();
  cachedClients = loadCachedClientList();
  const localLoadElapsedMs = performance.now() - localLoadStartedAt;
  console.log("[Photo Layout Debug] Local client cache load:", {
    clients: cachedClients.length,
    elapsedMs: Number(localLoadElapsedMs.toFixed(2)),
  });
  updateClientOptions();
  renderImagesGrid();

  if (!cachedClients.length) {
    setImagesStatus("Loading clients...");
    await refreshClientListFromApi();
    return;
  }

  void refreshClientListFromApi().catch((error) => {
    console.error("[Photo Layout Debug] Client list refresh failed.", error);
  });
}

async function refreshClientListFromApi() {
  const clientsPayload = await directoryApi.listMarketingPhotos({ clientsOnly: 1 });
  const clients = Array.isArray(clientsPayload?.clients) ? clientsPayload.clients : [];
  cachedClients = clients
    .map((item) => normalizeClientName(item?.name))
    .filter((value, index, array) => Boolean(value) && array.indexOf(value) === index);
  saveCachedClientList(cachedClients);
  updateClientOptions();
  renderImagesGrid();
}

async function loadPhotosForClient(clientName) {
  const requestedClient = String(clientName || "").trim();
  if (!requestedClient) {
    clientPhotoPool = [];
    renderAll();
    setImagesStatus("Select a client to load images.");
    return;
  }
  const normalizedClient = normalizeClientName(requestedClient);

  setImagesStatus(`Loading images for ${normalizedClient}...`);

  if (clientPhotoCache.has(normalizedClient)) {
    const localLoadStartedAt = performance.now();
    clientPhotoPool = clientPhotoCache.get(normalizedClient) || [];
    const localLoadElapsedMs = performance.now() - localLoadStartedAt;
    console.log("[Photo Layout Debug] Local client image cache load:", {
      client: normalizedClient,
      images: clientPhotoPool.length,
      elapsedMs: Number(localLoadElapsedMs.toFixed(2)),
    });
    renderAll();
    return;
  }

  const payload = await directoryApi.listMarketingPhotos({ client: normalizedClient });
  const photos = Array.isArray(payload?.photos) ? payload.photos : [];
  allPhotos = photos;
  console.log("[Photo Layout Debug] First 5 photo records:", allPhotos.slice(0, 5));
  const imagePhotos = photos.filter((photo) => isImagePhoto(photo));
  clientPhotoCache.set(normalizedClient, imagePhotos);
  clientPhotoPool = imagePhotos;
  renderAll();
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
    if (!canAccessPage(role, "photolayout")) {
      redirectToUnauthorized("photolayout");
      return;
    }

    renderTopNavigation({ role });
    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
    await loadPhotos();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("photolayout");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
    setImagesStatus(error?.message || "Could not load images.", true);
    setComposeStatus("Could not load composer.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

clientSelect?.addEventListener("change", () => {
  selectedClient = String(clientSelect.value || "").trim();
  void loadPhotosForClient(selectedClient);
});

zoomRange?.addEventListener("input", applyAdjustmentsFromControls);
panXRange?.addEventListener("input", applyAdjustmentsFromControls);
panYRange?.addEventListener("input", applyAdjustmentsFromControls);
resetAdjustBtn?.addEventListener("click", () => {
  const image = selectedImages[selectedSlotIndex];
  if (!image) {
    return;
  }
  image.zoom = 1;
  image.panX = 0;
  image.panY = 0;
  invalidateOutput();
  updateAdjustControls();
  renderStage();
});

gapEnabledInput?.addEventListener("change", () => {
  layoutStyle.gapEnabled = Boolean(gapEnabledInput.checked);
  invalidateOutput();
  updateStyleControls();
  renderStage();
});

gapRange?.addEventListener("input", () => {
  layoutStyle.gapPx = clamp(Number(gapRange.value) || 0, 0, 120);
  invalidateOutput();
  updateStyleControls();
  renderStage();
});

roundedEnabledInput?.addEventListener("change", () => {
  layoutStyle.roundedEnabled = Boolean(roundedEnabledInput.checked);
  invalidateOutput();
  updateStyleControls();
  renderStage();
});

cornerRadiusRange?.addEventListener("input", () => {
  layoutStyle.cornerRadiusPx = clamp(Number(cornerRadiusRange.value) || 0, 0, 240);
  invalidateOutput();
  updateStyleControls();
  renderStage();
});

generateOutputBtn?.addEventListener("click", () => {
  void generateOutput();
});
copyOutputBtn?.addEventListener("click", () => {
  void copyOutput();
});
saveOutputBtn?.addEventListener("click", () => {
  void saveOutput();
});

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

window.addEventListener("beforeunload", () => {
  if (latestOutputUrl) {
    URL.revokeObjectURL(latestOutputUrl);
  }
});

window.addEventListener("resize", () => {
  renderStage();
});

void init();
