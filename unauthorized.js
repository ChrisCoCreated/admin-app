import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { renderTopNavigation } from "./navigation.js?v=20260317";

const signOutBtn = document.getElementById("signOutBtn");
const deniedMessage = document.getElementById("deniedMessage");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

const PAGE_LABELS = {
  clients: "Clients",
  recruitment: "Recruitment",
  mapping: "Time Mapping",
  drivetime: "Our Geography",
  consultant: "Consultant",
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

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    renderTopNavigation({ role: profile?.role, currentPathname: window.location.pathname });
  } catch (error) {
    console.error(error);
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

setDeniedMessage();
void init();
