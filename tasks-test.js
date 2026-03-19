import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260317";

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const userMessage = document.getElementById("userMessage");
const taskAssignForm = document.getElementById("taskAssignForm");
const targetUserInput = document.getElementById("targetUserInput");
const anchorDateInput = document.getElementById("anchorDateInput");
const areaInput = document.getElementById("areaInput");
const taskSetInput = document.getElementById("taskSetInput");
const dryRunBtn = document.getElementById("dryRunBtn");
const createBtn = document.getElementById("createBtn");
const payloadOutput = document.getElementById("payloadOutput");
const responseOutput = document.getElementById("responseOutput");

const DEFAULT_TARGET_USER = "chris@planwithcare.co.uk";
const DEFAULT_TASK_SET = "Test";
const DEFAULT_AREA = "Colleague";

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
  createBtn.disabled = value;
  targetUserInput.disabled = value;
  anchorDateInput.disabled = value;
  areaInput.disabled = value;
  taskSetInput.disabled = value;
}

function pretty(value) {
  return JSON.stringify(value, null, 2);
}

function todayDateValue() {
  const now = new Date();
  const year = String(now.getUTCFullYear()).padStart(4, "0");
  const month = String(now.getUTCMonth() + 1).padStart(2, "0");
  const day = String(now.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function buildPayload({ dryRun }) {
  const targetUser = String(targetUserInput?.value || "").trim().toLowerCase();
  const taskSet = String(taskSetInput?.value || DEFAULT_TASK_SET).trim();
  const anchorDate = String(anchorDateInput?.value || "").trim();
  const area = String(areaInput?.value || DEFAULT_AREA).trim();

  const payload = {
    taskSet,
    area,
    dryRun,
  };

  if (targetUser) {
    payload.targetUser = targetUser;
  }
  if (anchorDate) {
    payload.anchorDate = anchorDate;
  }

  payloadOutput.textContent = pretty(payload);
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

    if (anchorDateInput && !anchorDateInput.value) {
      anchorDateInput.value = todayDateValue();
    }
    if (taskSetInput && !taskSetInput.value) {
      taskSetInput.value = DEFAULT_TASK_SET;
    }
    if (areaInput && !areaInput.value) {
      areaInput.value = DEFAULT_AREA;
    }
    if (targetUserInput) {
      targetUserInput.placeholder = `Defaults to ${DEFAULT_TARGET_USER}`;
    }

    buildPayload({ dryRun: true });
    setStatus("Ready to test /api/tasks/assign.");
  } catch (error) {
    console.error("[tasks-test] Init failed", error);
    setStatus(error?.message || "Could not initialise test page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

taskAssignForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  await submitRequest(false);
});

dryRunBtn?.addEventListener("click", async () => {
  await submitRequest(true);
});

targetUserInput?.addEventListener("input", () => buildPayload({ dryRun: true }));
anchorDateInput?.addEventListener("input", () => buildPayload({ dryRun: true }));
areaInput?.addEventListener("input", () => buildPayload({ dryRun: true }));
taskSetInput?.addEventListener("input", () => buildPayload({ dryRun: true }));

signOutBtn?.addEventListener("click", async () => {
  await authController.signOut();
});

init();
