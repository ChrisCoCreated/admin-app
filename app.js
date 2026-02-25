import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";

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
  onSignedIn: () => {
    setStatus("Signed in");
    window.location.href = "./clients.html";
  },
  onSignedOut: () => {
    setStatus("Signed out.");
  },
});

async function init() {
  try {
    setStatus("Checking session...");
    const account = await authController.restoreSession();
    if (!account) {
      setStatus("Please sign in");
      authCard.hidden = false;
    }
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
