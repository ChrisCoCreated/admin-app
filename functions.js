import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260702";
import { getRecruitmentWhatsAppMessage, getRecruitmentWhatsAppUrl } from "./recruitment-whatsapp.js?v=20260702";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const recruitmentWhatsappForm = document.getElementById("recruitmentWhatsappForm");
const candidateNameInput = document.getElementById("candidateNameInput");
const phoneNumberInput = document.getElementById("phoneNumberInput");
const copyRecruitmentMessageBtn = document.getElementById("copyRecruitmentMessageBtn");
const actionStatus = document.getElementById("actionStatus");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setActionStatus(message, isError = false) {
  actionStatus.textContent = message;
  actionStatus.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "functions").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function getFormValues() {
  return {
    candidateName: String(candidateNameInput?.value || "").trim(),
    phoneNumber: String(phoneNumberInput?.value || "").trim(),
  };
}

function openRecruitmentWhatsapp() {
  const { candidateName, phoneNumber } = getFormValues();
  const url = getRecruitmentWhatsAppUrl(phoneNumber, candidateName);
  if (!url) {
    setActionStatus("Enter a phone number first.", true);
    phoneNumberInput?.focus();
    return;
  }
  window.open(url, "_blank", "noopener,noreferrer");
  setActionStatus("Opened WhatsApp.");
}

async function copyRecruitmentMessage() {
  const { candidateName } = getFormValues();
  const message = getRecruitmentWhatsAppMessage(candidateName);
  try {
    await navigator.clipboard.writeText(message);
    setActionStatus("Copied message.");
  } catch (error) {
    console.error(error);
    setActionStatus("Could not copy message.", true);
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
    if (!canAccessPage(role, "functions")) {
      redirectToUnauthorized("functions");
      return;
    }

    renderTopNavigation({ role });

    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("functions");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize functions.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

recruitmentWhatsappForm?.addEventListener("submit", (event) => {
  event.preventDefault();
  openRecruitmentWhatsapp();
});

copyRecruitmentMessageBtn?.addEventListener("click", () => {
  void copyRecruitmentMessage();
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
