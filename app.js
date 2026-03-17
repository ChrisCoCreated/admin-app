import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { getAccessiblePages, getHomePageTiles, getPageMeta, renderTopNavigation } from "./navigation.js?v=20260317";

const signInBtn = document.getElementById("signInBtn");
const authState = document.getElementById("authState");
const authCard = document.querySelector(".auth-card");
const mainContainer = document.querySelector("main.container");
const signOutBtn = document.getElementById("signOutBtn");
const topbarActions = document.getElementById("topbarActions");
const heroSignInMessage = document.getElementById("heroSignInMessage");
const homeUserEmail = document.getElementById("homeUserEmail");
const homeUserPermissions = document.getElementById("homeUserPermissions");
const homeQuickLinks = document.getElementById("homeQuickLinks");
const homeQuickLinksMessage = document.getElementById("homeQuickLinksMessage");
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
  if (homeUserEmail) {
    homeUserEmail.textContent = "";
    homeUserEmail.hidden = true;
  }
  if (homeUserPermissions) {
    homeUserPermissions.textContent = "";
    homeUserPermissions.hidden = true;
  }
  if (heroSignInMessage) {
    heroSignInMessage.hidden = false;
  }
  if (homeQuickLinks) {
    homeQuickLinks.hidden = true;
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
  if (topbarActions) {
    topbarActions.hidden = false;
  }
  if (heroSignInMessage) {
    heroSignInMessage.hidden = true;
  }
}

function renderUserSummary(profile) {
  const role = String(profile?.role || "").trim().toLowerCase();
  const actualRole = String(profile?.actualRole || role).trim().toLowerCase();
  const email = String(profile?.email || "").trim();
  const pageLabels = getAccessiblePages(role)
    .map((pageKey) => getPageMeta(pageKey)?.label)
    .filter(Boolean);

  if (homeUserEmail) {
    homeUserEmail.textContent = email ? `Signed in as: ${email}` : "Signed in.";
    homeUserEmail.hidden = false;
  }
  if (homeUserPermissions) {
    const permissions = pageLabels.length ? pageLabels.join(", ") : "None";
    homeUserPermissions.textContent = profile?.previewingLoggedInUser
      ? `Permissions: ${permissions} (previewing as a signed-in user without admin permissions).`
      : `Permissions: ${permissions}`;
    homeUserPermissions.hidden = false;
  }
  if (heroSignInMessage) {
    heroSignInMessage.hidden = true;
  }
  if (profile?.previewingLoggedInUser && actualRole === "admin" && homeUserPermissions) {
    homeUserPermissions.classList.add("status-preview");
  } else if (homeUserPermissions) {
    homeUserPermissions.classList.remove("status-preview");
  }
}

function renderHomeQuickLinks(role) {
  if (!homeQuickLinks || !homeMenuGrid) {
    return;
  }

  const pageKeys = getHomePageTiles(role);
  homeMenuGrid.innerHTML = "";

  if (!pageKeys.length) {
    homeQuickLinks.hidden = true;
    return;
  }

  const isAdmin = String(role || "").trim().toLowerCase() === "admin";
  if (homeQuickLinksMessage) {
    homeQuickLinksMessage.textContent = isAdmin
      ? "Common admin destinations are pinned here for quicker access."
      : "All of your available pages are shown here because you only have access to a few.";
  }

  for (const pageKey of pageKeys) {
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

  homeQuickLinks.hidden = !homeMenuGrid.children.length;
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
  if (!pages.length && !profile?.previewingLoggedInUser) {
    window.location.href = "./unauthorized.html";
    return;
  }
  setSignedInUi();
  renderUserSummary(profile);
  renderHomeQuickLinks(role);
  renderTopNavigation({ role, currentPathname: "./index.html" });
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
