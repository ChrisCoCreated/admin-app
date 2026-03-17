const PREVIEW_LOGGED_IN_USER_KEY = "thrive.preview.loggedInUser";
const ACTUAL_ROLE_KEY = "thrive.preview.actualRole";

function normalizeRole(role) {
  return String(role || "").trim().toLowerCase();
}

function canUseSessionStorage() {
  return typeof window !== "undefined" && !!window.sessionStorage;
}

export function getStoredActualRole() {
  if (!canUseSessionStorage()) {
    return "";
  }
  return normalizeRole(window.sessionStorage.getItem(ACTUAL_ROLE_KEY));
}

export function setStoredActualRole(role) {
  if (!canUseSessionStorage()) {
    return;
  }
  const normalizedRole = normalizeRole(role);
  if (!normalizedRole) {
    window.sessionStorage.removeItem(ACTUAL_ROLE_KEY);
    return;
  }
  window.sessionStorage.setItem(ACTUAL_ROLE_KEY, normalizedRole);
}

export function isLoggedInUserPreviewEnabled() {
  if (!canUseSessionStorage()) {
    return false;
  }
  return window.sessionStorage.getItem(PREVIEW_LOGGED_IN_USER_KEY) === "true";
}

export function setLoggedInUserPreviewEnabled(enabled) {
  if (!canUseSessionStorage()) {
    return;
  }
  if (enabled) {
    window.sessionStorage.setItem(PREVIEW_LOGGED_IN_USER_KEY, "true");
    return;
  }
  window.sessionStorage.removeItem(PREVIEW_LOGGED_IN_USER_KEY);
}

export function applyRolePreview(profile = {}) {
  const actualRole = normalizeRole(profile?.role);
  if (actualRole) {
    setStoredActualRole(actualRole);
  }

  const previewingLoggedInUser = actualRole === "admin" && isLoggedInUserPreviewEnabled();
  const effectiveRole = previewingLoggedInUser ? "logged_in" : actualRole;

  return {
    ...profile,
    actualRole,
    role: effectiveRole,
    previewingLoggedInUser,
  };
}
