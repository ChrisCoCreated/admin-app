import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const templateSelect = document.getElementById("templateSelect");
const toInput = document.getElementById("toInput");
const subjectInput = document.getElementById("subjectInput");
const bodyInput = document.getElementById("bodyInput");
const draftOutlookBtn = document.getElementById("draftOutlookBtn");
const copyBodyBtn = document.getElementById("copyBodyBtn");
const actionStatus = document.getElementById("actionStatus");

let templates = [];

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setActionStatus(message, isError = false) {
  actionStatus.textContent = message;
  actionStatus.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "emailtemplates").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function applyTemplate(template) {
  if (!template) {
    subjectInput.value = "";
    bodyInput.value = "";
    return;
  }
  subjectInput.value = String(template.subject || "");
  bodyInput.value = String(template.body || "");
}

function renderTemplateOptions() {
  templateSelect.innerHTML = "";

  if (!templates.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "No templates available";
    templateSelect.appendChild(option);
    applyTemplate(null);
    return;
  }

  for (const template of templates) {
    const option = document.createElement("option");
    option.value = String(template.id || template.subject || "");
    option.textContent = String(template.title || template.subject || "Untitled template");
    templateSelect.appendChild(option);
  }

  templateSelect.selectedIndex = 0;
  applyTemplate(templates[0]);
}

async function loadTemplates() {
  const response = await fetch("./data/email-templates.json", {
    cache: "no-store",
  });
  if (!response.ok) {
    throw new Error("Could not load email templates.");
  }

  const payload = await response.json();
  if (!Array.isArray(payload?.templates)) {
    throw new Error("Invalid templates format.");
  }

  templates = payload.templates.filter((template) => typeof template === "object" && template !== null);
  renderTemplateOptions();
}

async function copyBodyText() {
  const body = String(bodyInput.value || "");
  if (!body.trim()) {
    setActionStatus("Body is empty.", true);
    return;
  }

  try {
    await navigator.clipboard.writeText(body);
    setActionStatus("Body copied.");
  } catch (error) {
    console.error(error);
    setActionStatus("Could not copy body text.", true);
  }
}

function openOutlookDraft() {
  const subject = String(subjectInput.value || "").trim();
  const body = String(bodyInput.value || "").trim();
  const to = String(toInput.value || "").trim();

  if (!subject || !body) {
    setActionStatus("Subject and body are required.", true);
    return;
  }

  const url = new URL("https://outlook.office.com/mail/deeplink/compose");
  url.searchParams.set("subject", subject);
  url.searchParams.set("body", body);
  if (to) {
    url.searchParams.set("to", to);
  }

  window.open(url.toString(), "_blank", "noopener,noreferrer");
  setActionStatus("Opened Outlook draft.");
}

async function fetchCurrentUser() {
  return directoryApi.getCurrentUser();
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
    if (!canAccessPage(role, "emailtemplates")) {
      redirectToUnauthorized("emailtemplates");
      return;
    }

    renderTopNavigation({ role });

    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}` : "Signed in");

    await loadTemplates();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("emailtemplates");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize email templates.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

templateSelect?.addEventListener("change", () => {
  const selectedId = String(templateSelect.value || "");
  const selectedTemplate = templates.find((template) => String(template.id || template.subject || "") === selectedId);
  applyTemplate(selectedTemplate || templates[0] || null);
  setActionStatus("");
});

copyBodyBtn?.addEventListener("click", () => {
  void copyBodyText();
});

draftOutlookBtn?.addEventListener("click", () => {
  openOutlookDraft();
});

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } finally {
    window.location.href = "./index.html";
  }
});

void init();
