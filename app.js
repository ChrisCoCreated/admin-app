import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";

const signInBtn = document.getElementById("signInBtn");
const authState = document.getElementById("authState");
const authCard = document.querySelector(".auth-card");
const mainContainer = document.querySelector("main.container");
const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
const ME_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/auth/me` : "/api/auth/me";

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

async function fetchCurrentUser() {
  const token = await authController.acquireToken([FRONTEND_CONFIG.apiScope]);
  const response = await fetch(ME_ENDPOINT, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Profile request failed (${response.status}): ${text || "Unknown error"}`);
  }

  return response.json();
}

async function routeToRoleHome() {
  const profile = await fetchCurrentUser();
  const role = String(profile?.role || "").trim().toLowerCase();
  if (role === "marketing") {
    window.location.href = "./marketing.html";
    return;
  }
  window.location.href = "./clients.html";
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
