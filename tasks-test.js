import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260317";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const userMessage = document.getElementById("userMessage");
const liveTaskSetRoot = document.getElementById("liveTaskSetRoot");
const reloadBtn = document.getElementById("reloadBtn");
const dryRunBtn = document.getElementById("dryRunBtn");
const createBtn = document.getElementById("createBtn");
const payloadOutput = document.getElementById("payloadOutput");
const responseOutput = document.getElementById("responseOutput");

const DEFAULT_TASK_SET = "Test";
const DEFAULT_AREA = "Colleague";
const DEFAULT_TARGET_USER = "chris@planwithcare.co.uk";

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
  reloadBtn.disabled = value;
  dryRunBtn.disabled = value;
  createBtn.disabled = value;
}

function pretty(value) {
  return JSON.stringify(value, null, 2);
}

function buildPayload({ dryRun }) {
  const payload = {
    taskSet: DEFAULT_TASK_SET,
    area: DEFAULT_AREA,
    dryRun,
    targetUser: DEFAULT_TARGET_USER,
  };

  payloadOutput.textContent = pretty(payload);
  return payload;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function renderLiveTemplates(templates = []) {
  if (!liveTaskSetRoot) {
    return;
  }

  if (!Array.isArray(templates) || !templates.length) {
    liveTaskSetRoot.innerHTML = '<p class="muted">No matching live rows found.</p>';
    return;
  }

  liveTaskSetRoot.innerHTML = templates
    .map((template) => {
      return `
        <article class="task-test-live-card">
          <h3>${escapeHtml(template.title || "Untitled task")}</h3>
          <p>${escapeHtml(template.description || "-")}</p>
          <dl class="detail-list">
            <div>
              <dt>Responsible Person</dt>
              <dd>${escapeHtml(template.responsiblePerson || "-")}</dd>
            </div>
            <div>
              <dt>Due Date Delay</dt>
              <dd>${escapeHtml(template.dueDateDelay ?? "-")}</dd>
            </div>
            <div>
              <dt>Item ID</dt>
              <dd>${escapeHtml(template.itemId || "-")}</dd>
            </div>
          </dl>
        </article>
      `;
    })
    .join("");
}

async function loadLiveTemplates() {
  if (liveTaskSetRoot) {
    liveTaskSetRoot.textContent = "Loading live task set content...";
  }

  const payload = await directoryApi.listTaskSetTemplates({
    taskSet: DEFAULT_TASK_SET,
    area: DEFAULT_AREA,
  });

  renderLiveTemplates(payload?.templates || []);
  return payload;
}

async function submitRequest(dryRun) {
  const payload = buildPayload({ dryRun });
  setBusy(true);
  setStatus(dryRun ? "Running dry run..." : "Creating task...");
  responseOutput.textContent = "Waiting for response...";

  try {
    const response = await directoryApi.assignTasks(payload);
    responseOutput.textContent = pretty(response);
    setStatus(dryRun ? "Dry run complete." : "Task creation request complete.");
  } catch (error) {
    responseOutput.textContent = pretty({
      error: {
        message: error?.message || String(error),
        code: error?.code || "",
        detail: error?.detail || "",
        status: error?.status || 0,
        correlationId: error?.correlationId || "",
      },
    });
    setStatus(error?.message || "Task request failed.", true);
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
    if (!canAccessPage(role, "taskstest")) {
      window.location.href = "./unauthorized.html?page=taskstest";
      return;
    }

    renderTopNavigation({ role });
    document.body.classList.remove("auth-pending");
    userMessage.textContent = `Signed in as ${profile?.email || "unknown"} (${role || "unknown role"}).`;

    buildPayload({ dryRun: true });
    await loadLiveTemplates();
    setStatus("Live task set row loaded. Click a button to run the backend.");
  } catch (error) {
    console.error("[tasks-test] Init failed", error);
    setStatus(error?.message || "Could not initialise test page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

reloadBtn?.addEventListener("click", async () => {
  setBusy(true);
  setStatus("Reloading live task set row...");
  try {
    const payload = await loadLiveTemplates();
    responseOutput.textContent = pretty(payload);
    setStatus("Live task set row reloaded.");
  } catch (error) {
    responseOutput.textContent = pretty({
      error: {
        message: error?.message || String(error),
      },
    });
    setStatus(error?.message || "Could not reload live task set row.", true);
  } finally {
    setBusy(false);
  }
});

dryRunBtn?.addEventListener("click", async () => {
  await submitRequest(true);
});

createBtn?.addEventListener("click", async () => {
  await submitRequest(false);
});
signOutBtn?.addEventListener("click", async () => {
  await authController.signOut();
});

init();
