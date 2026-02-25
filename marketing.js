import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { renderTopNavigation } from "./navigation.js";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const photosStatus = document.getElementById("photosStatus");
const photosGrid = document.getElementById("photosGrid");
const lightboxEl = document.getElementById("photoLightbox");
const lightboxImage = document.getElementById("lightboxImage");
const lightboxCaption = document.getElementById("lightboxCaption");
const lightboxCloseBtn = document.getElementById("lightboxCloseBtn");
const lightboxPrevBtn = document.getElementById("lightboxPrevBtn");
const lightboxNextBtn = document.getElementById("lightboxNextBtn");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);
let marketingPhotos = [];
let activePhotoIndex = -1;

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setPhotosStatus(message, isError = false) {
  if (!photosStatus) {
    return;
  }
  photosStatus.textContent = message;
  photosStatus.classList.toggle("error", isError);
}

function renderLightboxPhoto() {
  if (!lightboxImage || !lightboxCaption || !lightboxPrevBtn || !lightboxNextBtn) {
    return;
  }
  const photo = marketingPhotos[activePhotoIndex];
  if (!photo) {
    return;
  }

  lightboxImage.src = photo.imageUrl;
  lightboxImage.alt = photo.title || photo.client || "Client photo";
  lightboxCaption.textContent = `${photo.client || photo.title}${photo.title && photo.client !== photo.title ? ` - ${photo.title}` : ""}`;
  const canNavigate = marketingPhotos.length > 1;
  lightboxPrevBtn.disabled = !canNavigate;
  lightboxNextBtn.disabled = !canNavigate;
}

function openLightbox(index) {
  if (!lightboxEl || index < 0 || index >= marketingPhotos.length) {
    return;
  }
  activePhotoIndex = index;
  renderLightboxPhoto();
  lightboxEl.hidden = false;
  document.body.classList.add("lightbox-open");
}

function closeLightbox() {
  if (!lightboxEl || !lightboxImage) {
    return;
  }
  lightboxEl.hidden = true;
  lightboxImage.src = "";
  lightboxImage.alt = "";
  activePhotoIndex = -1;
  document.body.classList.remove("lightbox-open");
}

function stepLightbox(direction) {
  if (!marketingPhotos.length || activePhotoIndex < 0) {
    return;
  }
  const nextIndex = (activePhotoIndex + direction + marketingPhotos.length) % marketingPhotos.length;
  activePhotoIndex = nextIndex;
  renderLightboxPhoto();
}

function renderPhotoGrid(photos) {
  if (!photosGrid) {
    return;
  }

  photosGrid.innerHTML = "";
  if (!photos.length) {
    setPhotosStatus("No photos with client consent were found.");
    return;
  }

  const fragment = document.createDocumentFragment();
  for (const [index, photo] of photos.entries()) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "marketing-photo-card";
    button.setAttribute("aria-label", `Open ${photo.client || photo.title}`);

    const image = document.createElement("img");
    image.src = photo.imageUrl;
    image.alt = photo.title || photo.client || "Client photo";
    image.loading = "lazy";
    image.className = "marketing-photo-image";

    const caption = document.createElement("span");
    caption.className = "marketing-photo-caption";
    caption.textContent = photo.client || photo.title || "Untitled";

    button.append(image, caption);
    button.addEventListener("click", () => openLightbox(index));
    fragment.append(button);
  }

  photosGrid.append(fragment);
  setPhotosStatus(`${photos.length} consented photo${photos.length === 1 ? "" : "s"} loaded.`);
}

async function fetchCurrentUser() {
  return directoryApi.getCurrentUser();
}

async function loadPhotos() {
  setPhotosStatus("Loading consented photos...");
  const payload = await directoryApi.listMarketingPhotos();
  const photos = Array.isArray(payload?.photos) ? payload.photos : [];
  marketingPhotos = photos;
  renderPhotoGrid(photos);
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await fetchCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();

    if (role !== "marketing" && role !== "admin") {
      window.location.href = "./clients.html";
      return;
    }
    renderTopNavigation({ role });

    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
    await loadPhotos();
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
    setPhotosStatus(error?.message || "Could not load marketing photos.", true);
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

lightboxCloseBtn?.addEventListener("click", closeLightbox);
lightboxPrevBtn?.addEventListener("click", () => stepLightbox(-1));
lightboxNextBtn?.addEventListener("click", () => stepLightbox(1));
lightboxEl?.addEventListener("click", (event) => {
  if (event.target === lightboxEl) {
    closeLightbox();
  }
});

document.addEventListener("keydown", (event) => {
  if (!lightboxEl || lightboxEl.hidden) {
    return;
  }
  if (event.key === "Escape") {
    closeLightbox();
    return;
  }
  if (event.key === "ArrowLeft") {
    stepLightbox(-1);
    return;
  }
  if (event.key === "ArrowRight") {
    stepLightbox(1);
  }
});

void init();
