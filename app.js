import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage } from "./navigation.js";

const signInBtn = document.getElementById("signInBtn");
const authState = document.getElementById("authState");
const authCard = document.querySelector(".auth-card");
const mainContainer = document.querySelector("main.container");
if (authCard) {
  authCard.hidden = true;
}

function setStatus(message, isError = false) {
  authState.textContent = message;
  authState.classList.toggle("error", isError);
}

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
  authCard,
  mainContainer,
  onSignedIn: async () => {
    setStatus("Signed in");
    await routeToRoleHome();
  },
  onSignedOut: () => {
    setStatus("Signed out.");
  },
});
const directoryApi = createDirectoryApi(authController);

async function fetchCurrentUser() {
  return directoryApi.getCurrentUser();
}

async function routeToRoleHome() {
  const profile = await fetchCurrentUser();
  const role = String(profile?.role || "").trim().toLowerCase();
  if (canAccessPage(role, "marketing") && !canAccessPage(role, "clients")) {
    window.location.href = "./marketing.html";
    return;
  }
  if (canAccessPage(role, "clients")) {
    window.location.href = "./clients.html";
    return;
  }
  window.location.href = "./unauthorized.html";
}

async function init() {
  try {
    setStatus("Checking session...");
    const account = await authController.restoreSession();
    if (!account) {
      setStatus("Please sign in");
      authCard.hidden = false;
      return;
    }
    await routeToRoleHome();
  } catch (error) {
    if (error?.status === 403) {
      window.location.href = "./unauthorized.html";
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
    authCard.hidden = false;
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

signInBtn?.addEventListener("click", async () => {
  try {
    signInBtn.disabled = true;
    setStatus("Signing in...");
    await authController.signIn({
      scopes: ["openid", "profile", FRONTEND_CONFIG.apiScope],
    });
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Sign-in failed.", true);
    signInBtn.disabled = false;
  }
});

void init();
