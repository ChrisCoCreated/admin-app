import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const photosStatus = document.getElementById("photosStatus");
const photosGrid = document.getElementById("photosGrid");
const consentFilterBtn = document.getElementById("consentFilterBtn");
const lightboxEl = document.getElementById("photoLightbox");
const lightboxImage = document.getElementById("lightboxImage");
const lightboxVideo = document.getElementById("lightboxVideo");
const lightboxCaption = document.getElementById("lightboxCaption");
const lightboxCopyLinkBtn = document.getElementById("lightboxCopyLinkBtn");
const lightboxCloseBtn = document.getElementById("lightboxCloseBtn");
const lightboxPrevBtn = document.getElementById("lightboxPrevBtn");
const lightboxNextBtn = document.getElementById("lightboxNextBtn");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);
let allMarketingPhotos = [];
let visibleMarketingPhotos = [];
let consentFilterOn = false;
let activePhotoIndex = -1;
let attemptedVideoSources = new Set();
let copyLabelResetTimer = null;

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "marketing").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function setPhotosStatus(message, isError = false) {
  if (!photosStatus) {
    return;
  }
  photosStatus.textContent = message;
  photosStatus.classList.toggle("error", isError);
}

async function copyToClipboard(text) {
  const value = String(text || "").trim();
  if (!value) {
    throw new Error("No link available.");
  }
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(value);
    return;
  }
  throw new Error("Clipboard API unavailable.");
}

function isVideoMedia(photo) {
  const explicitType = String(photo?.mediaType || "").toLowerCase();
  if (explicitType === "video") {
    return true;
  }
  const source = String(photo?.mediaUrl || photo?.fileName || photo?.title || "").toLowerCase();
  return source.includes(".mp4") || source.includes(".mov") || source.includes(".webm");
}

function isLikelyImageUrl(url) {
  const source = String(url || "").toLowerCase();
  if (!source) {
    return false;
  }
  return (
    source.includes(".png") ||
    source.includes(".jpg") ||
    source.includes(".jpeg") ||
    source.includes(".webp") ||
    source.includes(".gif") ||
    source.includes(".bmp") ||
    source.includes("getpreview")
  );
}

function renderLightboxPhoto() {
  if (!lightboxImage || !lightboxVideo || !lightboxCaption || !lightboxPrevBtn || !lightboxNextBtn) {
    return;
  }
  const photo = visibleMarketingPhotos[activePhotoIndex];
  if (!photo) {
    return;
  }
  const originalUrl = photo.mediaUrl || photo.attachmentUrl || photo.imageUrl || "";
  if (lightboxCopyLinkBtn) {
    lightboxCopyLinkBtn.dataset.url = originalUrl;
    lightboxCopyLinkBtn.hidden = !originalUrl;
    lightboxCopyLinkBtn.textContent = "Copy link";
  }

  if (isVideoMedia(photo)) {
    attemptedVideoSources = new Set();
    lightboxImage.hidden = true;
    lightboxImage.src = "";
    lightboxImage.alt = "";
    lightboxVideo.hidden = false;
    lightboxVideo.src = photo.mediaUrl || photo.imageUrl;
    attemptedVideoSources.add(String(lightboxVideo.src || ""));
    lightboxVideo.load();
  } else {
    lightboxVideo.pause();
    lightboxVideo.src = "";
    lightboxVideo.hidden = true;
    lightboxImage.hidden = false;
    lightboxImage.src = photo.imageUrl;
    lightboxImage.alt = photo.title || photo.client || "Client photo";
  }
  lightboxCaption.textContent = `${photo.client || photo.title}${photo.title && photo.client !== photo.title ? ` - ${photo.title}` : ""}`;
  const canNavigate = visibleMarketingPhotos.length > 1;
  lightboxPrevBtn.disabled = !canNavigate;
  lightboxNextBtn.disabled = !canNavigate;
}

function openLightbox(index) {
  if (!lightboxEl || index < 0 || index >= visibleMarketingPhotos.length) {
    return;
  }
  activePhotoIndex = index;
  renderLightboxPhoto();
  lightboxEl.hidden = false;
  document.body.classList.add("lightbox-open");
}

function closeLightbox() {
  if (!lightboxEl || !lightboxImage || !lightboxVideo) {
    return;
  }
  lightboxEl.hidden = true;
  lightboxImage.src = "";
  lightboxImage.alt = "";
  lightboxImage.hidden = false;
  lightboxVideo.pause();
  lightboxVideo.src = "";
  lightboxVideo.hidden = true;
  if (lightboxCopyLinkBtn) {
    lightboxCopyLinkBtn.dataset.url = "";
    lightboxCopyLinkBtn.hidden = true;
    lightboxCopyLinkBtn.textContent = "Copy link";
  }
  if (copyLabelResetTimer) {
    clearTimeout(copyLabelResetTimer);
    copyLabelResetTimer = null;
  }
  attemptedVideoSources = new Set();
  activePhotoIndex = -1;
  document.body.classList.remove("lightbox-open");
}

function stepLightbox(direction) {
  if (!visibleMarketingPhotos.length || activePhotoIndex < 0) {
    return;
  }
  const nextIndex =
    (activePhotoIndex + direction + visibleMarketingPhotos.length) % visibleMarketingPhotos.length;
  activePhotoIndex = nextIndex;
  renderLightboxPhoto();
}

function renderPhotoGrid(photos) {
  if (!photosGrid) {
    return;
  }

  photosGrid.innerHTML = "";
  if (!photos.length) {
    setPhotosStatus(consentFilterOn ? "No consented photos were found." : "No photos were found.");
    return;
  }

  const fragment = document.createDocumentFragment();
  for (const [index, photo] of photos.entries()) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "marketing-photo-card";
    button.setAttribute("aria-label", `Open ${photo.client || photo.title}`);

    const mediaIsVideo = isVideoMedia(photo);
    const previewUrl = photo.imageUrl || photo.mediaUrl;
    let media;
    if (mediaIsVideo && !isLikelyImageUrl(previewUrl)) {
      media = document.createElement("video");
      media.src = photo.mediaUrl || previewUrl;
      media.className = "marketing-photo-image";
      media.muted = true;
      media.playsInline = true;
      media.preload = "metadata";
      media.setAttribute("aria-label", photo.title || photo.client || "Client video");
    } else {
      media = document.createElement("img");
      media.src = previewUrl;
      media.className = "marketing-photo-image";
      media.alt = photo.title || photo.client || "Client photo";
      media.loading = "lazy";
    }

    const caption = document.createElement("span");
    caption.className = "marketing-photo-caption";
    caption.textContent = photo.client || photo.title || "Untitled";

    const actions = document.createElement("div");
    actions.className = "marketing-photo-actions";
    const copyLinkBtn = document.createElement("button");
    copyLinkBtn.type = "button";
    copyLinkBtn.className = "marketing-photo-copy-link";
    copyLinkBtn.textContent = "Copy link";
    copyLinkBtn.addEventListener("click", async (event) => {
      event.stopPropagation();
      const sourceUrl = photo.mediaUrl || photo.attachmentUrl || photo.imageUrl || "";
      try {
        await copyToClipboard(sourceUrl);
        copyLinkBtn.textContent = "Copied";
        setTimeout(() => {
          copyLinkBtn.textContent = "Copy link";
        }, 1400);
      } catch {
        copyLinkBtn.textContent = "Failed";
        setTimeout(() => {
          copyLinkBtn.textContent = "Copy link";
        }, 1400);
      }
    });
    actions.append(copyLinkBtn);

    button.append(media, caption, actions);
    button.addEventListener("click", () => openLightbox(index));
    fragment.append(button);
  }

  photosGrid.append(fragment);
  const consentedCount = allMarketingPhotos.filter((photo) => photo.consented).length;
  if (consentFilterOn) {
    setPhotosStatus(`${photos.length} consented photo${photos.length === 1 ? "" : "s"} shown.`);
  } else {
    setPhotosStatus(
      `${photos.length} photo${photos.length === 1 ? "" : "s"} shown (${consentedCount} consented).`
    );
  }
}

function applyPhotoFilter() {
  visibleMarketingPhotos = consentFilterOn
    ? allMarketingPhotos.filter((photo) => photo.consented)
    : allMarketingPhotos.slice();
  renderPhotoGrid(visibleMarketingPhotos);
}

function updateConsentFilterButton() {
  if (!consentFilterBtn) {
    return;
  }
  consentFilterBtn.textContent = consentFilterOn ? "Show all photos" : "Show consented only";
  consentFilterBtn.classList.toggle("active", consentFilterOn);
}

async function fetchCurrentUser() {
  return directoryApi.getCurrentUser();
}

async function loadPhotos() {
  setPhotosStatus("Loading photos...");
  const payload = await directoryApi.listMarketingPhotos();
  const photos = Array.isArray(payload?.photos) ? payload.photos : [];
  allMarketingPhotos = photos;
  updateConsentFilterButton();
  applyPhotoFilter();
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

    if (!canAccessPage(role, "marketing")) {
      redirectToUnauthorized("marketing");
      return;
    }
    renderTopNavigation({ role });

    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
    await loadPhotos();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("marketing");
      return;
    }
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
lightboxCopyLinkBtn?.addEventListener("click", async () => {
  const sourceUrl = String(lightboxCopyLinkBtn.dataset.url || "").trim();
  try {
    await copyToClipboard(sourceUrl);
    lightboxCopyLinkBtn.textContent = "Copied";
  } catch {
    lightboxCopyLinkBtn.textContent = "Failed";
  }
  if (copyLabelResetTimer) {
    clearTimeout(copyLabelResetTimer);
  }
  copyLabelResetTimer = setTimeout(() => {
    if (lightboxCopyLinkBtn) {
      lightboxCopyLinkBtn.textContent = "Copy link";
    }
  }, 1400);
});
lightboxVideo?.addEventListener("error", () => {
  const photo = visibleMarketingPhotos[activePhotoIndex];
  if (!photo || !isVideoMedia(photo)) {
    return;
  }
  const candidates = [photo.mediaUrl, photo.attachmentUrl, photo.imageUrl]
    .map((value) => String(value || "").trim())
    .filter((value, index, array) => Boolean(value) && array.indexOf(value) === index);
  const nextSource = candidates.find((candidate) => !attemptedVideoSources.has(candidate));
  if (!nextSource) {
    return;
  }
  attemptedVideoSources.add(nextSource);
  lightboxVideo.src = nextSource;
  lightboxVideo.load();
});
consentFilterBtn?.addEventListener("click", () => {
  consentFilterOn = !consentFilterOn;
  closeLightbox();
  updateConsentFilterButton();
  applyPhotoFilter();
});
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
