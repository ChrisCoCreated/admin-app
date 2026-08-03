import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260601";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const templateSelect = document.getElementById("templateSelect");
const toInput = document.getElementById("toInput");
const subjectInput = document.getElementById("subjectInput");
const bodyInput = document.getElementById("bodyInput");
const attachmentPanel = document.getElementById("attachmentPanel");
const attachmentNotice = document.getElementById("attachmentNotice");
const attachmentDownloadLink = document.getElementById("attachmentDownloadLink");
const attachmentConfirmLabel = document.getElementById("attachmentConfirmLabel");
const attachmentConfirmCheckbox = document.getElementById("attachmentConfirmCheckbox");
const draftOutlookBtn = document.getElementById("draftOutlookBtn");
const draftWebBtn = document.getElementById("draftWebBtn");
const copyBodyBtn = document.getElementById("copyBodyBtn");
const actionStatus = document.getElementById("actionStatus");

let templates = [];
let selectedTemplate = null;

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

function requiresAttachmentConfirmation(template) {
  const attachments = Array.isArray(template?.attachments) ? template.attachments.filter(Boolean) : [];
  return attachments.length > 0;
}

function updateDraftButtons() {
  const requiresConfirmation = requiresAttachmentConfirmation(selectedTemplate);
  const confirmationComplete = !requiresConfirmation || Boolean(attachmentConfirmCheckbox?.checked);

  if (draftOutlookBtn) {
    draftOutlookBtn.disabled = !confirmationComplete;
  }
  if (draftWebBtn) {
    draftWebBtn.disabled = !confirmationComplete;
  }
}

function applyTemplate(template) {
  selectedTemplate = template || null;
  if (!template) {
    subjectInput.value = "";
    bodyInput.value = "";
    if (attachmentPanel) {
      attachmentPanel.hidden = true;
    }
    if (attachmentNotice) {
      attachmentNotice.textContent = "";
    }
    if (attachmentDownloadLink) {
      attachmentDownloadLink.hidden = true;
      attachmentDownloadLink.href = "#";
      attachmentDownloadLink.removeAttribute("download");
    }
    if (attachmentConfirmCheckbox) {
      attachmentConfirmCheckbox.checked = false;
    }
    if (attachmentConfirmLabel) {
      attachmentConfirmLabel.hidden = true;
    }
    updateDraftButtons();
    return;
  }
  subjectInput.value = String(template.subject || "");
  bodyInput.value = String(template.body || "");

  if (attachmentConfirmCheckbox) {
    attachmentConfirmCheckbox.checked = false;
  }

  if (attachmentPanel && attachmentNotice) {
    const attachments = Array.isArray(template.attachments) ? template.attachments.filter(Boolean) : [];
    if (attachments.length) {
      attachmentPanel.hidden = false;
      attachmentNotice.textContent = `Remember to attach: ${attachments.join(", ")}`;
      if (attachmentDownloadLink) {
        const downloadHref = String(template.attachmentDownloadHref || "").trim();
        if (downloadHref) {
          attachmentDownloadLink.href = downloadHref;
          attachmentDownloadLink.hidden = false;
          attachmentDownloadLink.setAttribute("download", attachments[0]);
        } else {
          attachmentDownloadLink.hidden = true;
          attachmentDownloadLink.href = "#";
          attachmentDownloadLink.removeAttribute("download");
        }
      }
      if (attachmentConfirmLabel) {
        attachmentConfirmLabel.hidden = false;
      }
    } else {
      attachmentPanel.hidden = true;
      attachmentNotice.textContent = "";
      if (attachmentDownloadLink) {
        attachmentDownloadLink.hidden = true;
        attachmentDownloadLink.href = "#";
        attachmentDownloadLink.removeAttribute("download");
      }
      if (attachmentConfirmLabel) {
        attachmentConfirmLabel.hidden = true;
      }
    }
  }

  updateDraftButtons();
}

function renderTemplateOptions() {
  templateSelect.innerHTML = "";

  const placeholderOption = document.createElement("option");
  placeholderOption.value = "";
  placeholderOption.textContent = "Selected Template";
  templateSelect.appendChild(placeholderOption);

  if (!templates.length) {
    applyTemplate(null);
    return;
  }

  for (const template of templates) {
    const option = document.createElement("option");
    option.value = String(template.id || template.subject || "");
    option.textContent = String(template.title || template.subject || "Untitled template");
    templateSelect.appendChild(option);
  }

  templateSelect.value = "";
  applyTemplate(null);
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
  if (requiresAttachmentConfirmation(selectedTemplate) && !attachmentConfirmCheckbox?.checked) {
    setActionStatus("Please confirm the required file(s) have been downloaded or attached first.", true);
    return;
  }

  const mailtoParts = [`subject=${encodeURIComponent(subject)}`, `body=${encodeURIComponent(body)}`];
  const mailtoUrl = `mailto:${encodeURIComponent(to)}?${mailtoParts.join("&")}`;
  window.location.href = mailtoUrl;
  setActionStatus("Tried opening your mail app. If it did not open, use Open in Web (Fallback).");
}

function openWebDraft() {
  const subject = String(subjectInput.value || "").trim();
  const body = String(bodyInput.value || "").trim();
  const to = String(toInput.value || "").trim();

  if (!subject || !body) {
    setActionStatus("Subject and body are required.", true);
    return;
  }
  if (requiresAttachmentConfirmation(selectedTemplate) && !attachmentConfirmCheckbox?.checked) {
    setActionStatus("Please confirm the required file(s) have been downloaded or attached first.", true);
    return;
  }

  const params = [`subject=${encodeURIComponent(subject)}`, `body=${encodeURIComponent(body)}`];
  if (to) {
    params.push(`to=${encodeURIComponent(to)}`);
  }
  const webUrl = `https://outlook.office.com/mail/deeplink/compose?${params.join("&")}`;
  window.open(webUrl, "_blank", "noopener,noreferrer");
  setActionStatus("Opened Outlook on the web.");
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
  const matchingTemplate = templates.find((template) => String(template.id || template.subject || "") === selectedId);
  applyTemplate(matchingTemplate || null);
  setActionStatus("");
});

attachmentConfirmCheckbox?.addEventListener("change", () => {
  updateDraftButtons();
  if (attachmentConfirmCheckbox.checked) {
    setActionStatus("");
  }
});

copyBodyBtn?.addEventListener("click", () => {
  void copyBodyText();
});

draftOutlookBtn?.addEventListener("click", () => {
  openOutlookDraft();
});

draftWebBtn?.addEventListener("click", () => {
  openWebDraft();
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
