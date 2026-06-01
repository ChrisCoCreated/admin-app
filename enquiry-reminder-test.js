import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260601";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const userMessage = document.getElementById("userMessage");
const dryRunBtn = document.getElementById("dryRunBtn");
const sendTrialBtn = document.getElementById("sendTrialBtn");
const responseOutput = document.getElementById("responseOutput");
const emailPreviewFrame = document.getElementById("emailPreviewFrame");
const summaryList = document.getElementById("summaryList");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});

const directoryApi = createDirectoryApi(authController);

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setBusy(value) {
  dryRunBtn.disabled = value;
  sendTrialBtn.disabled = value;
}

function pretty(value) {
  return JSON.stringify(value, null, 2);
}

function setSummaryField(field, value) {
  const target = summaryList?.querySelector(`[data-field="${field}"]`);
  if (target) {
    target.textContent = value;
  }
}

function renderSummary(payload = {}) {
  setSummaryField("activeCount", String(payload.activeCount ?? "-"));
  setSummaryField("onHoldCount", String(payload.onHoldCount ?? "-"));
  setSummaryField("includeOnHold", payload.includeOnHold === true ? "Yes" : payload.includeOnHold === false ? "No" : "-");
  setSummaryField("recipients", Array.isArray(payload.recipients) && payload.recipients.length ? payload.recipients.join(", ") : "-");
  setSummaryField("subject", payload.email?.subject || payload.subject || "-");
}

function renderEmailPreview(html) {
  if (!emailPreviewFrame) {
    return;
  }
  if (!html) {
    emailPreviewFrame.removeAttribute("srcdoc");
    return;
  }
  emailPreviewFrame.srcdoc = html;
}

function renderResponse(payload = {}) {
  responseOutput.textContent = pretty(payload);
  renderSummary(payload);
  renderEmailPreview(payload.email?.html || "");
}

function renderError(error) {
  const payload = {
    error: {
      message: error?.message || String(error),
      code: error?.code || "",
      detail: error?.detail || "",
      status: error?.status || 0,
      correlationId: error?.correlationId || "",
    },
  };
  responseOutput.textContent = pretty(payload);
  renderEmailPreview("");
}

async function runDryRun() {
  setBusy(true);
  setStatus("Running dry run against live enquiries...");
  responseOutput.textContent = "Waiting for dry-run response...";
  try {
    const payload = await directoryApi.dryRunEnquiryReminders();
    renderResponse(payload);
    if (payload.reason === "NO_MATCHING_ENQUIRIES") {
      setStatus("Dry run complete. No matching overdue enquiries.");
      return;
    }
    setStatus("Dry run complete. Preview updated below.");
  } catch (error) {
    renderError(error);
    setStatus(error?.message || "Dry run failed.", true);
  } finally {
    setBusy(false);
  }
}

async function sendTrialEmail() {
  const confirmed = window.confirm(
    "Send the trial reminder email now? This will use ENQUIRY_REMINDER_RECIPIENT_OVERRIDE, which should currently be chris@planwithcare.co.uk."
  );
  if (!confirmed) {
    return;
  }

  setBusy(true);
  setStatus("Sending trial reminder email...");
  responseOutput.textContent = "Waiting for send response...";
  try {
    const payload = await directoryApi.sendEnquiryReminderTrial();
    renderResponse(payload);
    if (payload.sent) {
      setStatus(`Trial email sent to ${(payload.recipients || []).join(", ") || "configured recipients"}.`);
      return;
    }
    setStatus("Send completed. No email was sent because there were no matching overdue enquiries.");
  } catch (error) {
    renderError(error);
    setStatus(error?.message || "Trial send failed.", true);
  } finally {
    setBusy(false);
  }
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
    if (!canAccessPage(role, "enquiryremindertest")) {
      window.location.href = "./unauthorized.html?page=enquiryremindertest";
      return;
    }

    renderTopNavigation({ role });
    document.body.classList.remove("auth-pending");
    userMessage.textContent = `Signed in as ${profile?.email || "unknown"} (${role || "unknown role"}).`;
    setStatus("Ready. Run a dry preview before sending a trial email.");
  } catch (error) {
    console.error("[enquiry-reminder-test] Init failed", error);
    setStatus(error?.message || "Could not initialise reminder test page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

dryRunBtn?.addEventListener("click", runDryRun);
sendTrialBtn?.addEventListener("click", sendTrialEmail);
signOutBtn?.addEventListener("click", async () => {
  await authController.signOut();
});

init();
