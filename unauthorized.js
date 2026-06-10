import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { renderTopNavigation } from "./navigation.js?v=20260601";

const signOutBtn = document.getElementById("signOutBtn");
const deniedMessage = document.getElementById("deniedMessage");
const deniedEmail = document.getElementById("deniedEmail");
const retrySignInBtn = document.getElementById("retrySignInBtn");
const differentAccountBtn = document.getElementById("differentAccountBtn");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

const PAGE_LABELS = {
  clients: "Clients",
  clientdata: "Client Report Generator",
  recruitment: "Recruitment",
  kpis: "Weekly KPIs",
  problems: "Problems to Solve",
  mapping: "Time Mapping",
  drivetime: "Our Geography",
  consultant: "Consultant",
  marketing: "Marketing",
  marketingreports: "Marketing Reports",
  photolayout: "Photo Layout",
  qrgenerator: "QR Generator",
  emailtemplates: "Email Templates",
  reports: "Reports",
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

function setDeniedEmail(email, sourceLabel = "Microsoft account") {
  if (!deniedEmail) {
    return;
  }
  const normalizedEmail = String(email || "").trim();
  deniedEmail.textContent = normalizedEmail ? `${sourceLabel}: ${normalizedEmail}` : "";
  deniedEmail.hidden = !normalizedEmail;
}

async function init() {
  try {
    const account = await authController.restoreSession();
    setDeniedEmail(account?.username || authController.getCachedAccount()?.username || "");
    if (!account) {
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    setDeniedEmail(profile?.email || account?.username || "", "Signed in as");
    renderTopNavigation({ role: profile?.role, currentPathname: window.location.pathname });
  } catch (error) {
    console.error(error);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

async function restartSignInFlow({ forceAccountSelection }) {
  const redirectUri = new URL("./index.html", window.location.href);
  if (forceAccountSelection) {
    redirectUri.searchParams.set("reauth", "1");
  }

  try {
    retrySignInBtn && (retrySignInBtn.disabled = true);
    differentAccountBtn && (differentAccountBtn.disabled = true);
    signOutBtn && (signOutBtn.disabled = true);
    await authController.signOut({
      redirectUri: redirectUri.toString(),
      forceAccountSelection,
    });
  } finally {
    window.location.href = redirectUri.toString();
  }
}

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut({
      redirectUri: new URL("./index.html", window.location.href).toString(),
    });
  } finally {
    window.location.href = "./index.html";
  }
});

retrySignInBtn?.addEventListener("click", () => {
  void restartSignInFlow({ forceAccountSelection: false });
});

differentAccountBtn?.addEventListener("click", () => {
  void restartSignInFlow({ forceAccountSelection: true });
});

setDeniedMessage();
void init();
