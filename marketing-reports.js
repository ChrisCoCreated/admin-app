import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260601";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");

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

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "marketingreports").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
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

    if (!canAccessPage(role, "marketingreports")) {
      redirectToUnauthorized("marketingreports");
      return;
    }

    renderTopNavigation({ role });
    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("marketingreports");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
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

void init();
