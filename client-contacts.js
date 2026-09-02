import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260601";

const searchInput = document.getElementById("searchInput");
const statusMessage = document.getElementById("statusMessage");
const warningState = document.getElementById("warningState");
const contactsBody = document.getElementById("contactsBody");
const emptyState = document.getElementById("emptyState");
const signOutBtn = document.getElementById("signOutBtn");
let contacts = [];

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function escapeHtml(value) {
  return String(value || "").replace(/[&<>'\"]/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;",
  })[character]);
}

function renderContacts() {
  const query = String(searchInput.value || "").trim().toLowerCase();
  const visible = contacts.filter((contact) => {
    const searchable = `${contact.name} ${contact.email}`.toLowerCase();
    return !query || searchable.includes(query);
  });
  contactsBody.innerHTML = visible.map((contact) => `
    <tr>
      <td>${escapeHtml(contact.name)}</td>
      <td>${contact.email ? escapeHtml(contact.email) : '<span class="muted">No email recorded</span>'}</td>
    </tr>`).join("");
  emptyState.hidden = visible.length > 0;
  setStatus(`${visible.length} of ${contacts.length} active client${contacts.length === 1 ? "" : "s"}.`);
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
    if (!canAccessPage(role, "clientcontacts")) {
      window.location.href = "./unauthorized.html?page=clientcontacts";
      return;
    }
    renderTopNavigation({ role });
    const payload = await directoryApi.listActiveClientContacts();
    contacts = Array.isArray(payload?.contacts) ? payload.contacts : [];
    const warnings = Array.isArray(payload?.warnings) ? payload.warnings.filter(Boolean) : [];
    warningState.hidden = warnings.length === 0;
    warningState.textContent = warnings.join(" ");
    renderContacts();
  } catch (error) {
    if (error?.status === 403) {
      window.location.href = "./unauthorized.html?page=clientcontacts";
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not load client contacts.", true);
  }
}

searchInput.addEventListener("input", renderContacts);
signOutBtn.addEventListener("click", () => authController.signOut());
void init();
