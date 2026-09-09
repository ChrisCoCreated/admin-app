import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260909";

const DIAGNOSTICS_ADMIN_EMAIL = "chris@planwithcare.co.uk";
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("accessDiagnosticsStatus");
const content = document.getElementById("accessDiagnosticsContent");
const tableBody = document.getElementById("accessDiagnosticsTableBody");
const checkEmailInput = document.getElementById("diagnosticsCheckEmail");
const checkButton = document.getElementById("diagnosticsCheckBtn");
const checkSummary = document.getElementById("diagnosticsCheckSummary");
const authController = createAuthController({ tenantId: FRONTEND_CONFIG.tenantId, clientId: FRONTEND_CONFIG.spaClientId });
const directoryApi = createDirectoryApi(authController);

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function setText(id, value) {
  const node = document.getElementById(id);
  if (node) node.textContent = value || "-";
}

function renderDiagnostics(data) {
  setText("diagnosticsVersion", data.appVersion);
  setText("diagnosticsEnvironment", [data.deployment?.environment, data.deployment?.gitCommit?.slice(0, 7)].filter(Boolean).join(" · "));
  setText("diagnosticsEmail", data.signedInIdentity?.resolvedEmail);
  setText("diagnosticsRole", data.signedInIdentity?.role);
  setText("diagnosticsClaims", (data.signedInIdentity?.tokenEmailCandidates || []).join(" · "));
  const checkedEmail = data.accessCheck?.email || "";
  if (checkEmailInput && document.activeElement !== checkEmailInput) {
    checkEmailInput.value = checkedEmail;
  }
  if (checkSummary) {
    checkSummary.textContent = checkedEmail
      ? `Checking deployed access lists for: ${checkedEmail}`
      : "Enter an email address to check the deployed access lists.";
  }
  tableBody.innerHTML = (data.accessCheck?.configuration || []).map((entry) => `
    <tr>
      <td>${escapeHtml(entry.key)}</td>
      <td>${entry.configured ? "Yes" : "No"}</td>
      <td>${escapeHtml(entry.emailCount)}</td>
      <td><span class="access-match ${entry.matchesSignedInEmail ? "is-match" : ""}">${entry.matchesSignedInEmail ? "Matched" : "No match"}</span></td>
    </tr>`).join("");
}

async function loadDiagnostics(email = "") {
  const data = await directoryApi.getAccessDiagnostics(email);
  renderDiagnostics(data);
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }
    const profile = await directoryApi.getCurrentUser();
    const email = String(profile?.email || "").trim().toLowerCase();
    if (email !== DIAGNOSTICS_ADMIN_EMAIL || !canAccessPage(profile?.role, "accessdiagnostics")) {
      window.location.href = "./unauthorized.html?page=accessdiagnostics";
      return;
    }
    renderTopNavigation({ role: profile.role });
    await loadDiagnostics(email);
    content.hidden = false;
    statusMessage.hidden = true;
  } catch (error) {
    console.error(error);
    statusMessage.textContent = error?.message || "Could not load access diagnostics.";
    statusMessage.classList.add("error");
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

signOutBtn?.addEventListener("click", async () => {
  await authController.signOut({ redirectUri: new URL("./index.html", window.location.href).toString() });
  window.location.href = "./index.html";
});

checkButton?.addEventListener("click", async () => {
  const email = String(checkEmailInput?.value || "").trim();
  if (!email) {
    checkEmailInput?.focus();
    return;
  }
  checkButton.disabled = true;
  try {
    await loadDiagnostics(email);
  } catch (error) {
    console.error(error);
    if (checkSummary) checkSummary.textContent = error?.message || "Could not check this email.";
  } finally {
    checkButton.disabled = false;
  }
});

void init();
