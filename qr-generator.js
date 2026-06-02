import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260601";

const DEFAULT_BASE_URL = "https://www.thrivehomecare.co.uk/";
const QR_SERVICE_URL = "https://api.qrserver.com/v1/create-qr-code/";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const baseUrlInput = document.getElementById("baseUrlInput");
const sourceInput = document.getElementById("sourceInput");
const mediumInput = document.getElementById("mediumInput");
const mediumField = document.getElementById("mediumField");
const editMediumBtn = document.getElementById("editMediumBtn");
const campaignInput = document.getElementById("campaignInput");
const generatedUrlOutput = document.getElementById("generatedUrlOutput");
const copyGeneratedUrlBtn = document.getElementById("copyGeneratedUrlBtn");
const openGeneratedUrlLink = document.getElementById("openGeneratedUrlLink");
const downloadQrLink = document.getElementById("downloadQrLink");
const qrPreviewImage = document.getElementById("qrPreviewImage");
const qrFormStatus = document.getElementById("qrFormStatus");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

function setStatus(message, isError = false) {
  if (!statusMessage) {
    return;
  }
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setFormStatus(message, isError = false) {
  if (!qrFormStatus) {
    return;
  }
  qrFormStatus.textContent = message;
  qrFormStatus.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "qrgenerator").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function normalizeUrlInput(value) {
  const trimmed = String(value || "").trim();
  if (!trimmed) {
    return DEFAULT_BASE_URL;
  }
  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed;
  }
  return `https://${trimmed}`;
}

function buildTrackedUrl() {
  const pageUrl = new URL(normalizeUrlInput(baseUrlInput?.value), window.location.origin);
  const source = String(sourceInput?.value || "").trim();
  const medium = String(mediumInput?.value || "").trim() || "referral";
  const campaign = String(campaignInput?.value || "").trim();

  pageUrl.searchParams.set("utm_source", source);
  pageUrl.searchParams.set("utm_medium", medium);
  pageUrl.searchParams.set("utm_campaign", campaign);

  return {
    campaign,
    medium,
    source,
    url: pageUrl.toString(),
  };
}

function getQrImageUrl(value, size = 420) {
  const params = new URLSearchParams({
    data: value,
    size: `${size}x${size}`,
    format: "png",
    margin: "18",
  });
  return `${QR_SERVICE_URL}?${params.toString()}`;
}

function updateQrPreview() {
  try {
    const tracked = buildTrackedUrl();
    const hasRequiredInputs = Boolean(tracked.source && tracked.campaign);

    generatedUrlOutput.value = tracked.url;
    openGeneratedUrlLink.href = tracked.url;

    const previewUrl = getQrImageUrl(tracked.url);
    qrPreviewImage.src = previewUrl;
    downloadQrLink.href = getQrImageUrl(tracked.url, 1000);
    downloadQrLink.classList.toggle("is-disabled", !hasRequiredInputs);
    copyGeneratedUrlBtn.disabled = !hasRequiredInputs;

    if (!hasRequiredInputs) {
      setFormStatus("Add a source and campaign to finish the QR link.");
      return;
    }

    setFormStatus("QR code ready.");
  } catch (error) {
    generatedUrlOutput.value = "";
    qrPreviewImage.removeAttribute("src");
    openGeneratedUrlLink.href = DEFAULT_BASE_URL;
    downloadQrLink.href = "#";
    downloadQrLink.classList.add("is-disabled");
    copyGeneratedUrlBtn.disabled = true;
    setFormStatus(error?.message || "Enter a valid Thrive page link.", true);
  }
}

async function copyGeneratedUrl() {
  const value = String(generatedUrlOutput?.value || "").trim();
  if (!value || !navigator?.clipboard?.writeText) {
    setFormStatus("Copy is unavailable in this browser.", true);
    return;
  }

  await navigator.clipboard.writeText(value);
  setFormStatus("Link copied.");
}

async function fetchCurrentUser() {
  return directoryApi.getCurrentUser();
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

    if (!canAccessPage(role, "qrgenerator")) {
      redirectToUnauthorized("qrgenerator");
      return;
    }

    renderTopNavigation({ role });
    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
    updateQrPreview();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("qrgenerator");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

[baseUrlInput, sourceInput, mediumInput, campaignInput].forEach((input) => {
  input?.addEventListener("input", updateQrPreview);
});

copyGeneratedUrlBtn?.addEventListener("click", () => {
  void copyGeneratedUrl();
});

editMediumBtn?.addEventListener("click", () => {
  const isExpanded = !mediumField?.hidden;
  if (mediumField) {
    mediumField.hidden = isExpanded;
  }
  editMediumBtn.setAttribute("aria-expanded", String(!isExpanded));
  editMediumBtn.textContent = isExpanded ? "Edit medium" : "Hide medium";
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
