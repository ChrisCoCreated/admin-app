import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const clientsNavLink = document.getElementById("clientsNavLink");
const mappingNavLink = document.getElementById("mappingNavLink");

const API_BASE_URL = (FRONTEND_CONFIG.apiBaseUrl || "").replace(/\/+$/, "");
const ME_ENDPOINT = API_BASE_URL ? `${API_BASE_URL}/api/auth/me` : "/api/auth/me";

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

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

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await fetchCurrentUser();
    const role = String(profile?.role || "").trim().toLowerCase();

    if (role !== "marketing" && role !== "admin") {
      window.location.href = "./clients.html";
      return;
    }
    if (role === "admin") {
      if (clientsNavLink) {
        clientsNavLink.hidden = false;
      }
      if (mappingNavLink) {
        mappingNavLink.hidden = false;
      }
    }

    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");
  } catch (error) {
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
