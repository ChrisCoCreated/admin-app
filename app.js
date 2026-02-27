import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { getAccessiblePages, getPageMeta, renderTopNavigation } from "./navigation.js";

const signInBtn = document.getElementById("signInBtn");
const authState = document.getElementById("authState");
const authCard = document.querySelector(".auth-card");
const mainContainer = document.querySelector("main.container");
const signOutBtn = document.getElementById("signOutBtn");
const topbarActions = document.getElementById("topbarActions");
const homeMenuSection = document.getElementById("homeMenuSection");
const homeMenuGrid = document.getElementById("homeMenuGrid");
if (authCard) {
  authCard.hidden = true;
}

function setStatus(message, isError = false) {
  authState.textContent = message;
  authState.classList.toggle("error", isError);
}

function setSignedOutUi() {
  if (authCard) {
    authCard.hidden = false;
  }
  if (homeMenuSection) {
    homeMenuSection.hidden = true;
  }
  if (homeMenuGrid) {
    homeMenuGrid.innerHTML = "";
  }
  if (topbarActions) {
    topbarActions.hidden = true;
  }
  if (signInBtn) {
    signInBtn.disabled = false;
  }
}

function setSignedInUi() {
  if (authCard) {
    authCard.hidden = true;
  }
  if (homeMenuSection) {
    homeMenuSection.hidden = false;
  }
  if (topbarActions) {
    topbarActions.hidden = false;
  }
}

function renderHomeMenu(role) {
  if (!homeMenuGrid) {
    return;
  }

  const pages = getAccessiblePages(role);
  homeMenuGrid.innerHTML = "";
  renderTopNavigation({ role, currentPathname: "./index.html" });

  for (const pageKey of pages) {
    const page = getPageMeta(pageKey);
    if (!page) {
      continue;
    }
    const link = document.createElement("a");
    link.className = "home-menu-link";
    link.href = page.href;
    link.textContent = page.label;
    homeMenuGrid.appendChild(link);
  }
}

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
  authCard,
  mainContainer,
  onSignedIn: async () => {
    setStatus("Signed in");
    await renderRoleMenu();
  },
  onSignedOut: () => {
    setStatus("Signed out.");
    setSignedOutUi();
  },
});
const directoryApi = createDirectoryApi(authController);

async function fetchCurrentUser() {
  return directoryApi.getCurrentUser();
}

async function renderRoleMenu() {
  const profile = await fetchCurrentUser();
  const role = String(profile?.role || "").trim().toLowerCase();
  const pages = getAccessiblePages(role);
  if (!pages.length) {
    window.location.href = "./unauthorized.html";
    return;
  }
  setSignedInUi();
  renderHomeMenu(role);
}

async function init() {
  try {
    setStatus("Checking session...");
    const account = await authController.restoreSession();
    if (!account) {
      setStatus("Please sign in.");
      setSignedOutUi();
      return;
    }
    await renderRoleMenu();
  } catch (error) {
    if (error?.status === 403) {
      window.location.href = "./unauthorized.html";
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize authentication.", true);
    setSignedOutUi();
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

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Sign-out failed.", true);
    signOutBtn.disabled = false;
  }
});

void init();
