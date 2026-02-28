import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js";

const searchInput = document.getElementById("searchInput");
const refreshBtn = document.getElementById("refreshBtn");
const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const pillList = document.getElementById("pillList");
const emptyState = document.getElementById("emptyState");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let allTasks = [];
let busy = false;

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setBusy(value) {
  busy = value;
  refreshBtn.disabled = value;
  const pinButtons = pillList?.querySelectorAll?.(".task-pill-pin") || [];
  for (const button of pinButtons) {
    button.disabled = value;
  }
}

function formatDate(value) {
  if (!value) {
    return "No due date";
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return String(value);
  }
  return date.toLocaleDateString();
}

function taskKey(task) {
  return `${String(task?.provider || "").toLowerCase()}|${String(task?.externalTaskId || "")}`;
}

function isCreatedInLastMonth(task) {
  const created = task?.createdDateTimeUtc;
  if (!created) {
    return false;
  }
  const createdDate = new Date(created);
  if (Number.isNaN(createdDate.getTime())) {
    return false;
  }

  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - 1);
  return createdDate >= cutoff;
}

function includesBySimpleRule(task) {
  return Boolean(task?.dueDateTimeUtc) || isCreatedInLastMonth(task);
}

function filteredTasks() {
  const query = String(searchInput?.value || "").trim().toLowerCase();
  const base = allTasks.filter((task) => includesBySimpleRule(task));

  if (!query) {
    return base;
  }

  return base.filter((task) => String(task?.title || "").toLowerCase().includes(query));
}

function sortedTasks(tasks) {
  return [...tasks].sort((a, b) => {
    const aPinned = a?.overlay?.pinned === true;
    const bPinned = b?.overlay?.pinned === true;
    if (aPinned !== bPinned) {
      return aPinned ? -1 : 1;
    }

    const aDue = a?.dueDateTimeUtc || null;
    const bDue = b?.dueDateTimeUtc || null;
    if (!aDue && bDue) {
      return 1;
    }
    if (aDue && !bDue) {
      return -1;
    }
    if (aDue && bDue) {
      const aTime = Date.parse(aDue);
      const bTime = Date.parse(bDue);
      if (Number.isFinite(aTime) && Number.isFinite(bTime) && aTime !== bTime) {
        return aTime - bTime;
      }
    }

    return String(a?.title || "").localeCompare(String(b?.title || ""), undefined, {
      sensitivity: "base",
    });
  });
}

function render() {
  const tasks = sortedTasks(filteredTasks());
  pillList.innerHTML = "";

  if (!tasks.length) {
    emptyState.hidden = false;
    return;
  }
  emptyState.hidden = true;

  for (const task of tasks) {
    const wrapper = document.createElement("div");
    wrapper.className = "task-pill";

    const title = document.createElement("span");
    title.className = "task-pill-title";
    title.textContent = task.title || "Untitled task";

    const due = document.createElement("span");
    due.className = "task-pill-due";
    due.textContent = formatDate(task.dueDateTimeUtc);

    const pinButton = document.createElement("button");
    pinButton.type = "button";
    pinButton.className = "secondary task-pill-pin";
    pinButton.disabled = busy;
    pinButton.textContent = task?.overlay?.pinned === true ? "Unpin" : "Pin";
    pinButton.addEventListener("click", async () => {
      if (busy) {
        return;
      }

      setBusy(true);
      try {
        const result = await directoryApi.upsertTaskOverlay({
          provider: task.provider,
          externalTaskId: task.externalTaskId,
          patch: {
            title: String(task?.title || "").trim(),
            pinned: !(task?.overlay?.pinned === true),
          },
        });

        const key = taskKey(task);
        const target = allTasks.find((entry) => taskKey(entry) === key);
        if (target) {
          target.overlay = {
            ...(target.overlay || {}),
            ...(result?.overlay || {}),
          };
        }

        setStatus("Pin updated.");
        render();
      } catch (error) {
        console.error("[simple-tasks] Pin update failed", error);
        setStatus(error?.message || "Could not update pin.", true);
      } finally {
        setBusy(false);
      }
    });

    wrapper.appendChild(title);
    wrapper.appendChild(due);
    wrapper.appendChild(pinButton);
    pillList.appendChild(wrapper);
  }
}

async function refresh() {
  setBusy(true);
  setStatus("Loading tasks...");
  try {
    const payload = await directoryApi.getUnifiedTasks();
    allTasks = Array.isArray(payload?.tasks) ? payload.tasks : [];
    setStatus(`Loaded ${allTasks.length} task(s).`);
    render();
  } catch (error) {
    console.error("[simple-tasks] Refresh failed", error);
    setStatus(error?.message || "Could not load tasks.", true);
    pillList.innerHTML = "";
    emptyState.hidden = false;
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
    if (!canAccessPage(role, "simpletasks")) {
      window.location.href = "./unauthorized.html?page=simpletasks";
      return;
    }

    renderTopNavigation({ role });
    await refresh();
  } catch (error) {
    console.error("[simple-tasks] Init failed", error);
    setStatus(error?.message || "Could not initialize page.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

searchInput?.addEventListener("input", () => {
  render();
});

refreshBtn?.addEventListener("click", async () => {
  if (busy) {
    return;
  }
  await refresh();
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
