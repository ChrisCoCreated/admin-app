import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";

const signOutBtn = document.getElementById("signOutBtn");
const deniedMessage = document.getElementById("deniedMessage");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});

const PAGE_LABELS = {
  clients: "Clients",
  mapping: "Time Mapping",
  marketing: "Marketing",
};

function getPageLabel(page) {
  const key = String(page || "").trim().toLowerCase();
  return PAGE_LABELS[key] || "this page";
}

function setDeniedMessage() {
  const params = new URLSearchParams(window.location.search);
  const page = params.get("page");
  if (!deniedMessage) {
    return;
  }
  deniedMessage.textContent = `You do not have permission to view ${getPageLabel(page)}.`;
}

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

setDeniedMessage();
