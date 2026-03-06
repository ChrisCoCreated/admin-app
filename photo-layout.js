import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const MAX_SELECTION = 6;
const PAN_LIMIT = 100;
const EXPORT_WIDTH = 2000;

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
let imagePhotos = [];
let selectedClient = "";
let selectedImages = [];
let selectedSlotIndex = -1;
let activeLayoutId = LAYOUTS[0].id;
let dragState = null;
let latestOutputBlob = null;
let latestOutputUrl = "";

const exportImageCache = new Map();

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
  const target = normalizeClientName(selectedClient);
  return imagePhotos.filter((photo) => normalizeClientName(photo.client) === target);
}

function updateClientOptions() {
  const current = selectedClient;
  const clients = Array.from(new Set(imagePhotos.map((photo) => normalizeClientName(photo.client)))).sort((a, b) =>
    a.localeCompare(b)
  );
  clientSelect.innerHTML = "";
  if (!clients.length) {
    clientSelect.innerHTML = '<option value="">No clients with images</option>';
    selectedClient = "";
    return;
  }
  for (const client of clients) {
    const option = document.createElement("option");
    option.value = client;
    option.textContent = client;
    clientSelect.append(option);
  }
  if (clients.includes(current)) {
    clientSelect.value = current;
    selectedClient = current;
  } else {
    selectedClient = clients[0];
    clientSelect.value = selectedClient;
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
    sourceUrl: photo.imageUrl || photo.mediaUrl || "",
    previewUrl: photo.imageUrl || photo.mediaUrl || "",
    zoom: 1,
    panX: 0,
    panY: 0,
  });
  if (selectedSlotIndex < 0) {
    selectedSlotIndex = 0;
  }
  renderAll();
}

function renderImagesGrid() {
  imagesGrid.innerHTML = "";
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

function swapSelectedImages(fromIndex, toIndex) {
  if (fromIndex < 0 || toIndex < 0 || fromIndex >= selectedImages.length || toIndex >= selectedImages.length) {
    return;
  }
  const temp = selectedImages[fromIndex];
  selectedImages[fromIndex] = selectedImages[toIndex];
  selectedImages[toIndex] = temp;
  selectedSlotIndex = toIndex;
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

    row.append(info, removeBtn);
    selectedImagesList.append(row);
  });
}

function renderStage() {
  const layout = getActiveLayout();
  layoutStage.innerHTML = "";
  layoutStage.style.setProperty("--layout-stage-aspect", String(layout.aspect));

  layout.slots.forEach((slot, slotIndex) => {
    const slotEl = document.createElement("div");
    const assignedImage = selectedImages[slotIndex];
    slotEl.className = `layout-slot${slotIndex === selectedSlotIndex ? " active" : ""}${assignedImage ? " has-image" : ""}`;
    slotEl.style.left = `${slot.x * 100}%`;
    slotEl.style.top = `${slot.y * 100}%`;
    slotEl.style.width = `${slot.w * 100}%`;
    slotEl.style.height = `${slot.h * 100}%`;
    slotEl.dataset.slotIndex = String(slotIndex);
    slotEl.addEventListener("click", () => {
      selectedSlotIndex = slotIndex;
      updateAdjustControls();
      renderStage();
      renderSelectedImagesList();
    });

    slotEl.draggable = Boolean(assignedImage);
    slotEl.addEventListener("dragstart", (event) => {
      if (!assignedImage) {
        event.preventDefault();
        return;
      }
      event.dataTransfer?.setData("text/plain", String(slotIndex));
      event.dataTransfer.effectAllowed = "move";
    });
    slotEl.addEventListener("dragover", (event) => {
      event.preventDefault();
      if (assignedImage) {
        event.dataTransfer.dropEffect = "move";
      }
    });
    slotEl.addEventListener("drop", (event) => {
      event.preventDefault();
      const fromIndex = Number(event.dataTransfer?.getData("text/plain"));
      if (!Number.isFinite(fromIndex) || fromIndex === slotIndex) {
        return;
      }
      swapSelectedImages(fromIndex, slotIndex);
    });

    if (assignedImage) {
      const media = document.createElement("img");
      media.className = "layout-slot-image";
      media.src = assignedImage.previewUrl;
      media.alt = assignedImage.title;
      media.draggable = false;
      media.style.transform = `translate(${assignedImage.panX}%, ${assignedImage.panY}%) scale(${assignedImage.zoom})`;

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
        media.style.transform = `translate(${image.panX}%, ${image.panY}%) scale(${image.zoom})`;
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
    }. Drag one slot onto another to reorder.`
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
  renderStage();
}

function drawImageIntoSlot(ctx, img, slot, transform) {
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
  ctx.rect(slotX, slotY, slotW, slotH);
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

async function loadExportImage(sourceUrl) {
  const cacheKey = String(sourceUrl || "").trim();
  if (!cacheKey) {
    throw new Error("Missing source image URL.");
  }
  if (exportImageCache.has(cacheKey)) {
    return exportImageCache.get(cacheKey);
  }

  const payload = await directoryApi.getMarketingMedia({ url: cacheKey });
  const blob = decodeBase64ToBlob(payload?.dataBase64, payload?.mimeType);
  const objectUrl = URL.createObjectURL(blob);
  try {
    const image = await loadImageElement(objectUrl);
    exportImageCache.set(cacheKey, image);
    return image;
  } finally {
    URL.revokeObjectURL(objectUrl);
  }
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
  for (let index = 0; index < renderableSlots; index += 1) {
    const imageState = selectedImages[index];
    const sourceUrl = imageState?.sourceUrl;
    if (!sourceUrl) {
      continue;
    }
    const img = await loadExportImage(sourceUrl);
    const slot = layout.slots[index];
    drawImageIntoSlot(
      ctx,
      img,
      {
        x: slot.x * canvas.width,
        y: slot.y * canvas.height,
        w: slot.w * canvas.width,
        h: slot.h * canvas.height,
      },
      imageState
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
}

async function loadPhotos() {
  setImagesStatus("Loading photos...");
  const payload = await directoryApi.listMarketingPhotos();
  allPhotos = Array.isArray(payload?.photos) ? payload.photos : [];
  imagePhotos = allPhotos.filter((photo) => isImagePhoto(photo));
  updateClientOptions();
  if (!selectedClient && imagePhotos.length) {
    selectedClient = normalizeClientName(imagePhotos[0].client);
  }
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
  selectedClient = normalizeClientName(clientSelect.value);
  renderImagesGrid();
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
  updateAdjustControls();
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

void init();
