import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260314a";

const ORG_PEOPLE = [
  { name: "Nathan", email: "nathan@planwithcare.co.uk" },
  { name: "Rebecca", email: "rebecca@planwithcare.co.uk" },
  { name: "Peter", email: "peter@planwithcare.co.uk" },
  { name: "Agota", email: "agota@planwithcare.co.uk" },
  { name: "Miska", email: "michalina@thrivehomecare.co.uk" },
  { name: "Claire", email: "claire@planwithcare.co.uk" },
  { name: "Alise", email: "alise@planwithcare.co.uk" },
  { name: "Paul", email: "paul@planwithcare.co.uk" },
];

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const actionStatus = document.getElementById("actionStatus");
const toggleCreatePanelBtn = document.getElementById("toggleCreatePanelBtn");
const agendaCreatePanel = document.getElementById("agendaCreatePanel");
const agendaPeoplePicker = document.getElementById("agendaPeoplePicker");
const agendaCreateForm = document.getElementById("agendaCreateForm");
const agendaTitleInput = document.getElementById("agendaTitleInput");
const agendaTypeSelect = document.getElementById("agendaTypeSelect");
const agendaPrivateInput = document.getElementById("agendaPrivateInput");
const agendaSearchInput = document.getElementById("agendaSearchInput");
const agendaList = document.getElementById("agendaList");
const agendaListEmpty = document.getElementById("agendaListEmpty");
const agendaEmptyState = document.getElementById("agendaEmptyState");
const agendaDetailWrap = document.getElementById("agendaDetailWrap");
const agendaTitle = document.getElementById("agendaTitle");
const agendaMeta = document.getElementById("agendaMeta");
const agendaAttendees = document.getElementById("agendaAttendees");
const agendaSettingsForm = document.getElementById("agendaSettingsForm");
const agendaTitleEditInput = document.getElementById("agendaTitleEditInput");
const agendaPeopleSummaryInput = document.getElementById("agendaPeopleSummaryInput");
const agendaPrivateEditInput = document.getElementById("agendaPrivateEditInput");
const saveAgendaSettingsBtn = document.getElementById("saveAgendaSettingsBtn");
const cancelAgendaSettingsBtn = document.getElementById("cancelAgendaSettingsBtn");
const toggleItemSearchBtn = document.getElementById("toggleItemSearchBtn");
const itemSearchField = document.getElementById("itemSearchField");
const itemSearchInput = document.getElementById("itemSearchInput");
const itemStageFilter = document.getElementById("itemStageFilter");
const agendaItemsList = document.getElementById("agendaItemsList");
const agendaItemsEmpty = document.getElementById("agendaItemsEmpty");
const agendaItemForm = document.getElementById("agendaItemForm");
const agendaItemAdvanced = document.getElementById("agendaItemAdvanced");
const itemTitleInput = document.getElementById("itemTitleInput");
const itemDetailEditor = document.getElementById("itemDetailEditor");
const itemStageSelect = document.getElementById("itemStageSelect");
const itemPrivateInput = document.getElementById("itemPrivateInput");
const itemUrgentInput = document.getElementById("itemUrgentInput");
const itemImportantInput = document.getElementById("itemImportantInput");
const saveAgendaItemBtn = document.getElementById("saveAgendaItemBtn");
const insertLinkBtn = document.getElementById("insertLinkBtn");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);
const AGENDA_SUMMARY_CACHE_PREFIX = "thrive.agendas.summary.v1";
const AGENDA_DEBUG = true;

let currentUser = null;
let agendas = [];
let agendaDetailsById = new Map();
let selectedAgendaId = "";
let selectedItemId = "";
let busy = false;
let dragItemId = "";
let createPanelExpanded = false;
let agendaSettingsExpanded = false;
let agendaItemComposerExpanded = false;
let itemSearchExpanded = false;
let itemDetailsExpanded = false;
let creatingTaskItemId = "";
let completingTaskKey = "";
let taskComposerItemId = "";
const taskDraftsByItemId = new Map();
const loadingAgendaDetails = new Set();

function logAgendaDebug(message, details) {
  if (!AGENDA_DEBUG) {
    return;
  }
  if (details === undefined) {
    console.log(`[agendas] ${message}`);
    return;
  }
  console.log(`[agendas] ${message}`, details);
}

function setBusy(value) {
  busy = value;
  saveAgendaItemBtn.disabled = value;
  saveAgendaSettingsBtn.disabled = value;
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setCreatePanelExpanded(value) {
  createPanelExpanded = value === true;
  agendaCreatePanel.hidden = !createPanelExpanded;
  if (toggleCreatePanelBtn) {
    toggleCreatePanelBtn.textContent = createPanelExpanded ? "Minimise" : "New meeting agenda";
    toggleCreatePanelBtn.setAttribute("aria-expanded", createPanelExpanded ? "true" : "false");
  }
}

function setAgendaSettingsExpanded(value) {
  agendaSettingsExpanded = value === true;
  agendaSettingsForm.hidden = !agendaSettingsExpanded;
}

function setItemSearchExpanded(value) {
  itemSearchExpanded = value === true;
  if (itemSearchField) {
    itemSearchField.hidden = !itemSearchExpanded;
  }
  if (toggleItemSearchBtn) {
    toggleItemSearchBtn.setAttribute("aria-expanded", itemSearchExpanded ? "true" : "false");
  }
  if (itemSearchExpanded) {
    itemSearchInput?.focus();
  } else {
    if (itemSearchInput) {
      itemSearchInput.value = "";
    }
    if (itemStageFilter) {
      itemStageFilter.value = "";
    }
  }
}

function hasComposerContent() {
  const title = String(itemTitleInput?.value || "").trim();
  const detailText = String(itemDetailEditor?.textContent || "").replace(/\u00a0/g, " ").trim();
  return Boolean(title || detailText);
}

function setAgendaItemComposerExpanded(value) {
  agendaItemComposerExpanded = value === true;
  agendaItemAdvanced.hidden = !agendaItemComposerExpanded;
}

function setActionStatus(message, isError = false) {
  actionStatus.textContent = message;
  actionStatus.classList.toggle("error", isError);
}

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function displayNameForEmail(email) {
  const known = ORG_PEOPLE.find((person) => normalizeEmail(person.email) === normalizeEmail(email));
  if (known) {
    return known.name;
  }
  const local = String(email || "").split("@")[0].replace(/[._-]+/g, " ").trim();
  return local
    .split(/\s+/)
    .filter(Boolean)
    .map((part) => part.slice(0, 1).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");
}

function selectedAgenda() {
  return agendaDetailsById.get(selectedAgendaId) || agendas.find((agenda) => agenda.id === selectedAgendaId) || null;
}

function selectedItem() {
  const agenda = selectedAgenda();
  if (!agenda) {
    return null;
  }
  return agenda.items.find((item) => item.id === selectedItemId) || null;
}

function selectedPeopleFromCreateForm() {
  return Array.from(agendaPeoplePicker.querySelectorAll('input[type="checkbox"]:checked'))
    .map((input) => {
      const email = String(input.value || "").trim();
      const person = ORG_PEOPLE.find((entry) => normalizeEmail(entry.email) === normalizeEmail(email));
      return person || { email, name: displayNameForEmail(email) };
    });
}

function renderPeoplePicker() {
  agendaPeoplePicker.innerHTML = "";
  ORG_PEOPLE.forEach((person) => {
    const label = document.createElement("label");
    label.className = "agenda-person-option";
    label.innerHTML = `
      <span class="agenda-person-check">
        <input type="checkbox" value="${escapeHtml(person.email)}" />
      </span>
      <span class="agenda-person-copy">
        <strong>${escapeHtml(person.name)}</strong>
      </span>
    `;
    agendaPeoplePicker.appendChild(label);
  });
}

function filteredAgendas() {
  const query = String(agendaSearchInput.value || "").trim().toLowerCase();
  if (!query) {
    return agendas;
  }
  return agendas.filter((agenda) => {
    const haystack = [
      agenda.title,
      agenda.ownerEmail,
      ...(Array.isArray(agenda.participantNames) ? agenda.participantNames : []),
      ...(Array.isArray(agenda.participantEmails) ? agenda.participantEmails : []),
    ]
      .join(" ")
      .toLowerCase();
    return haystack.includes(query);
  });
}

function filteredItems(agenda) {
  const allItems = Array.isArray(agenda?.items) ? agenda.items : [];
  const query = String(itemSearchInput.value || "").trim().toLowerCase();
  const stage = String(itemStageFilter.value || "").trim().toLowerCase();
  return allItems.filter((item) => {
    if (stage && String(item.stageTag || "").trim().toLowerCase() !== stage) {
      return false;
    }
    if (!query) {
      return true;
    }
    const detailText = String(item.detailHtml || "").replace(/<[^>]+>/g, " ").toLowerCase();
    return `${String(item.title || "").toLowerCase()} ${detailText}`.includes(query);
  });
}

function agendaTasksForItem(item) {
  return Array.isArray(item?.taskMaps) ? item.taskMaps : [];
}

function hasLinkedTask(item) {
  return agendaTasksForItem(item).length > 0;
}

function taskLinkKey(task) {
  return `${String(task?.provider || "").trim().toLowerCase()}|${String(task?.externalTaskId || "").trim()}`;
}

function formatTaskDueDate(value) {
  if (!value) {
    return "No due date";
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return String(value);
  }
  return parsed.toLocaleDateString();
}

function padDatePart(value) {
  return String(value).padStart(2, "0");
}

function formatDateInputValue(date) {
  return `${date.getFullYear()}-${padDatePart(date.getMonth() + 1)}-${padDatePart(date.getDate())}`;
}

function todayDateInputValue() {
  return formatDateInputValue(new Date());
}

function nextWeekdayDateInputValue(targetDay) {
  const date = new Date();
  const currentDay = date.getDay();
  let offset = (targetDay - currentDay + 7) % 7;
  if (offset === 0) {
    offset = 7;
  }
  date.setDate(date.getDate() + offset);
  return formatDateInputValue(date);
}

function taskDraftForItem(item) {
  const itemId = String(item?.id || "").trim();
  if (!itemId) {
    return { title: "", dueDate: todayDateInputValue() };
  }
  const existing = taskDraftsByItemId.get(itemId);
  if (existing) {
    return existing;
  }
  const initial = {
    title: String(item?.title || "").trim(),
    dueDate: todayDateInputValue(),
  };
  taskDraftsByItemId.set(itemId, initial);
  return initial;
}

function updateTaskDraft(itemId, patch) {
  const current = taskDraftsByItemId.get(itemId) || { title: "", dueDate: todayDateInputValue() };
  taskDraftsByItemId.set(itemId, {
    ...current,
    ...patch,
  });
}

function dueDateTimeUtcFromDateInput(value) {
  const date = String(value || "").trim();
  return date ? `${date}T12:00:00.000Z` : null;
}

function agendaDisplayTitle(agenda) {
  const detailedAgenda = agendaDetailsById.get(agenda?.id);
  const title = String(agenda?.title || detailedAgenda?.title || "").trim();
  if (title) {
    logAgendaDebug("Resolved agenda title.", {
      agendaId: agenda?.id || "",
      summaryTitle: String(agenda?.title || ""),
      detailTitle: String(detailedAgenda?.title || ""),
      chosenTitle: title,
    });
    return title;
  }
  const members = Array.isArray(agenda?.members) ? agenda.members : Array.isArray(detailedAgenda?.members) ? detailedAgenda.members : [];
  const memberSummary = members
    .map((member) => member.displayName || displayNameForEmail(member.userEmail))
    .filter(Boolean)
    .join(", ");
  if (memberSummary) {
    logAgendaDebug("Fell back to member summary for agenda title.", {
      agendaId: agenda?.id || "",
      summaryTitle: String(agenda?.title || ""),
      detailTitle: String(detailedAgenda?.title || ""),
      memberSummary,
    });
    return memberSummary;
  }
  logAgendaDebug("Agenda title missing entirely.", {
    agendaId: agenda?.id || "",
    summaryTitle: String(agenda?.title || ""),
    detailTitle: String(detailedAgenda?.title || ""),
    agenda,
    detailedAgenda,
  });
  return "Untitled agenda";
}

function renderAgendaList() {
  const list = filteredAgendas();
  agendaList.innerHTML = "";
  agendaListEmpty.hidden = list.length > 0;

  list.forEach((agenda) => {
    const isOwner = normalizeEmail(agenda.ownerEmail) === normalizeEmail(currentUser?.email);
    const card = document.createElement("article");
    card.className = "agenda-list-card";
    if (agenda.id === selectedAgendaId) {
      card.classList.add("is-selected");
    }
    const detailedAgenda = agendaDetailsById.get(agenda?.id);
    const members = Array.isArray(agenda?.members) && agenda.members.length
      ? agenda.members
      : Array.isArray(detailedAgenda?.members)
        ? detailedAgenda.members
        : [];
    const people = members.map((member) => member.displayName || displayNameForEmail(member.userEmail));
    const title = agendaDisplayTitle(agenda);
    const top = document.createElement("div");
    top.className = "agenda-list-card-top";

    const mainButton = document.createElement("button");
    mainButton.type = "button";
    mainButton.className = "agenda-list-card-main";

    const titleEl = document.createElement("strong");
    titleEl.textContent = title;
    mainButton.appendChild(titleEl);

    if (agenda.isPrivate) {
      const badge = document.createElement("span");
      badge.className = "agenda-list-badge";
      badge.textContent = "Private";
      mainButton.appendChild(badge);
    }

    const peopleEl = document.createElement("small");
    peopleEl.textContent = people.join(", ") || "Just you";
    mainButton.appendChild(peopleEl);
    top.appendChild(mainButton);

    if (isOwner) {
      const editButton = document.createElement("button");
      editButton.type = "button";
      editButton.className = "ghost subtle icon-only agenda-card-edit-btn";
      editButton.setAttribute("aria-label", "Edit agenda");
      const icon = document.createElement("span");
      icon.setAttribute("aria-hidden", "true");
      icon.textContent = "✎";
      editButton.appendChild(icon);
      editButton.addEventListener("click", (event) => {
        event.stopPropagation();
        void openAgenda(agenda.id, { editSettings: true });
      });
      top.appendChild(editButton);
    }

    card.appendChild(top);
    mainButton.addEventListener("click", () => {
      void openAgenda(agenda.id);
    });
    agendaList.appendChild(card);
  });
}

async function openAgenda(agendaId, options = {}) {
  selectedAgendaId = agendaId;
  selectedItemId = "";
  setAgendaSettingsExpanded(options.editSettings === true);
  renderAgendaList();
  renderAgendaDetail();
  await loadAgendaDetail(agendaId);
  if (options.editSettings === true) {
    setAgendaSettingsExpanded(true);
  }
}

function resetItemForm() {
  selectedItemId = "";
  itemTitleInput.value = "";
  itemDetailEditor.innerHTML = "<p></p>";
  itemStageSelect.value = "";
  itemPrivateInput.checked = false;
  itemUrgentInput.checked = false;
  itemImportantInput.checked = false;
  setAgendaItemComposerExpanded(false);
}

function populateItemForm(item) {
  if (!item) {
    resetItemForm();
    return;
  }
  selectedItemId = item.id;
  itemTitleInput.value = item.title || "";
  itemDetailEditor.innerHTML = item.detailHtml || "<p></p>";
  itemStageSelect.value = item.stageTag || "";
  itemPrivateInput.checked = item.isPrivate === true;
  itemUrgentInput.checked = item.isUrgent === true;
  itemImportantInput.checked = item.isImportant === true;
  setAgendaItemComposerExpanded(true);
}

function participantSummary(agenda) {
  return agenda.members
    .map((member) => member.displayName || displayNameForEmail(member.userEmail))
    .filter(Boolean)
    .join(", ");
}

function renderAgendaAttendees(agenda) {
  if (!agendaAttendees) {
    return;
  }
  const members = Array.isArray(agenda?.members) ? agenda.members : [];
  agendaAttendees.innerHTML = "";
  agendaAttendees.hidden = members.length === 0;

  members.forEach((member) => {
    const pill = document.createElement("span");
    pill.className = "agenda-attendee-pill";
    pill.textContent = member.displayName || displayNameForEmail(member.userEmail);
    agendaAttendees.appendChild(pill);
  });
}

function renderAgendaItems(agenda) {
  const items = filteredItems(agenda);
  agendaItemsList.innerHTML = "";
  agendaItemsEmpty.hidden = items.length > 0;

  function buildDropZone(insertIndex) {
    const zone = document.createElement("div");
    zone.className = "agenda-drop-zone";
    zone.setAttribute("aria-hidden", "true");
    zone.addEventListener("dragover", (event) => {
      event.preventDefault();
      zone.classList.add("is-active");
    });
    zone.addEventListener("dragleave", () => {
      zone.classList.remove("is-active");
    });
    zone.addEventListener("drop", async (event) => {
      event.preventDefault();
      zone.classList.remove("is-active");
      if (!dragItemId) {
        return;
      }
      await reorderItemToIndex(agenda, dragItemId, insertIndex);
      dragItemId = "";
    });
    return zone;
  }

  if (items.length) {
    agendaItemsList.appendChild(buildDropZone(0));
  }

  items.forEach((item, index) => {
    const card = document.createElement("article");
    card.className = "agenda-item-card";
    card.draggable = true;
    if (item.id === selectedItemId) {
      card.classList.add("is-selected");
    }

    const linkedTasks = agendaTasksForItem(item);
    const badges = [
      item.stageTag ? `<span class="agenda-badge">${escapeHtml(item.stageTag)}</span>` : "",
      item.isUrgent ? '<span class="agenda-badge is-urgent">Urgent</span>' : "",
      item.isImportant ? '<span class="agenda-badge is-important">Important</span>' : "",
      item.isPrivate ? '<span class="agenda-badge">Private</span>' : "",
      linkedTasks.length ? `<span class="agenda-badge is-task">${linkedTasks.length} task${linkedTasks.length === 1 ? "" : "s"}</span>` : "",
    ].join("");
    const previewHtml = itemDetailsExpanded ? item.detailHtml || "<p></p>" : "";
    const draft = taskDraftForItem(item);
    const linkedTasksHtml = itemDetailsExpanded && linkedTasks.length
      ? `
        <div class="agenda-linked-task-list">
          ${linkedTasks
            .map((task) => {
              const isCompleted = task?.isCompleted === true;
              const key = taskLinkKey(task);
              return `
                <div class="agenda-linked-task-row">
                  <div class="agenda-linked-task-copy">
                    <strong>${escapeHtml(task.title || item.title || "Task")}</strong>
                    <span>${escapeHtml(formatTaskDueDate(task.dueDateTimeUtc))} · ${escapeHtml(
                      task.ownerEmail ? displayNameForEmail(task.ownerEmail) : "You"
                    )}</span>
                  </div>
                  ${
                    isCompleted
                      ? '<span class="agenda-task-check is-complete" aria-label="Completed"><span aria-hidden="true">✓</span></span>'
                      : `<button
                          type="button"
                          class="ghost subtle icon-only agenda-task-complete-btn"
                          data-task-key="${escapeHtml(key)}"
                          aria-label="Mark task complete"
                          title="Mark task complete"
                          ${completingTaskKey === key ? "disabled" : ""}
                        ><span aria-hidden="true">${completingTaskKey === key ? "…" : "✓"}</span></button>`
                  }
                </div>
              `;
            })
            .join("")}
        </div>
      `
      : "";
    const taskComposerHtml =
      taskComposerItemId === item.id
        ? `
          <form class="agenda-task-create-form" data-item-id="${escapeHtml(item.id)}">
            <label class="field" for="agendaTaskTitle-${escapeHtml(item.id)}">
              Task title
              <input
                id="agendaTaskTitle-${escapeHtml(item.id)}"
                name="taskTitle"
                type="text"
                value="${escapeHtml(draft.title)}"
                placeholder="Task title"
              />
            </label>
            <label class="field" for="agendaTaskDue-${escapeHtml(item.id)}">
              Due date
              <input
                id="agendaTaskDue-${escapeHtml(item.id)}"
                name="taskDueDate"
                type="date"
                value="${escapeHtml(draft.dueDate)}"
              />
            </label>
            <div class="agenda-task-date-presets" role="group" aria-label="Quick due date options">
              <button type="button" class="ghost subtle agenda-task-preset-btn" data-preset="today">Today</button>
              <button type="button" class="ghost subtle agenda-task-preset-btn" data-preset="tomorrow">Tomorrow</button>
              <button type="button" class="ghost subtle agenda-task-preset-btn" data-preset="monday">Monday</button>
              <button type="button" class="ghost subtle agenda-task-preset-btn" data-preset="wednesday">Wednesday</button>
            </div>
            <div class="agenda-task-form-actions">
              <button type="submit" ${creatingTaskItemId === item.id ? "disabled" : ""}>
                ${creatingTaskItemId === item.id ? "Creating..." : "Create task"}
              </button>
              <button type="button" class="ghost agenda-task-cancel-btn">Cancel</button>
            </div>
          </form>
        `
        : "";
    const showAddTaskIcon = taskComposerItemId !== item.id;

    card.innerHTML = `
      <div class="agenda-item-card-head">
        <strong>${escapeHtml(item.title)}</strong>
        <div class="agenda-item-card-head-actions">
          ${
            showAddTaskIcon
              ? `<button
                  type="button"
                  class="ghost subtle icon-only agenda-task-link-icon"
                  aria-label="Add task"
                  title="Add task"
                >
                  <span aria-hidden="true">+</span>
                </button>`
              : ""
          }
          <button
            type="button"
            class="agenda-expand-toggle"
            aria-label="${itemDetailsExpanded ? "Hide item details" : "Show item details"}"
            aria-expanded="${itemDetailsExpanded ? "true" : "false"}"
          >
            <span aria-hidden="true">${itemDetailsExpanded ? "⌃" : "⌄"}</span>
          </button>
        </div>
      </div>
      <div class="agenda-item-badges">${badges}</div>
      ${itemDetailsExpanded ? `<div class="agenda-item-preview">${previewHtml}</div>` : ""}
      ${linkedTasksHtml}
      <div class="agenda-item-actions">${taskComposerHtml}</div>
    `;

    card.addEventListener("click", () => {
      populateItemForm(item);
      renderAgendaItems(agenda);
    });
    card.querySelector(".agenda-expand-toggle")?.addEventListener("click", (event) => {
      event.stopPropagation();
      itemDetailsExpanded = !itemDetailsExpanded;
      renderAgendaItems(agenda);
    });
    card.addEventListener("dragstart", () => {
      dragItemId = item.id;
      card.classList.add("is-dragging");
    });
    card.addEventListener("dragend", () => {
      dragItemId = "";
      card.classList.remove("is-dragging");
    });
    card.querySelector(".agenda-item-actions")?.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    card.querySelector(".agenda-linked-task-list")?.addEventListener("click", (event) => {
      event.stopPropagation();
    });
    card.querySelector(".agenda-task-link-icon")?.addEventListener("click", (event) => {
      event.stopPropagation();
      taskComposerItemId = item.id;
      renderAgendaItems(agenda);
    });
    const taskForm = card.querySelector(".agenda-task-create-form");
    taskForm?.addEventListener("submit", async (event) => {
      event.preventDefault();
      event.stopPropagation();
      const formData = new FormData(taskForm);
      updateTaskDraft(item.id, {
        title: String(formData.get("taskTitle") || "").trim(),
        dueDate: String(formData.get("taskDueDate") || "").trim(),
      });
      await createTaskForAgendaItem(agenda, item);
    });
    taskForm?.querySelector('[name="taskTitle"]')?.addEventListener("input", (event) => {
      updateTaskDraft(item.id, { title: String(event.target?.value || "") });
    });
    taskForm?.querySelector('[name="taskDueDate"]')?.addEventListener("input", (event) => {
      updateTaskDraft(item.id, { dueDate: String(event.target?.value || "") });
    });
    taskForm?.querySelectorAll(".agenda-task-preset-btn").forEach((button) => {
      button.addEventListener("click", (event) => {
        event.stopPropagation();
        const preset = String(button.getAttribute("data-preset") || "").trim().toLowerCase();
        let dueDate = todayDateInputValue();
        if (preset === "tomorrow") {
          const date = new Date();
          date.setDate(date.getDate() + 1);
          dueDate = formatDateInputValue(date);
        } else if (preset === "monday") {
          dueDate = nextWeekdayDateInputValue(1);
        } else if (preset === "wednesday") {
          dueDate = nextWeekdayDateInputValue(3);
        }
        updateTaskDraft(item.id, { dueDate });
        const dueInput = taskForm.querySelector('[name="taskDueDate"]');
        if (dueInput) {
          dueInput.value = dueDate;
        }
      });
    });
    card.querySelector(".agenda-task-cancel-btn")?.addEventListener("click", (event) => {
      event.stopPropagation();
      taskComposerItemId = "";
      renderAgendaItems(agenda);
    });
    card.querySelectorAll(".agenda-task-complete-btn").forEach((button) => {
      button.addEventListener("click", async (event) => {
        event.stopPropagation();
        const key = String(button.getAttribute("data-task-key") || "").trim();
        const task = linkedTasks.find((entry) => taskLinkKey(entry) === key);
        if (!task) {
          return;
        }
        await markAgendaTaskComplete(agenda, item, task);
      });
    });

    agendaItemsList.appendChild(card);
    agendaItemsList.appendChild(buildDropZone(index + 1));
  });
}

async function createTaskForAgendaItem(agenda, item) {
  if (!agenda?.id || !item?.id) {
    return;
  }

  const draft = taskDraftForItem(item);
  const title = String(draft.title || item.title || "").trim();
  if (!title) {
    setActionStatus("Item title is required before adding a task.", true);
    return;
  }

  const dueDateTimeUtc = dueDateTimeUtcFromDateInput(draft.dueDate);
  const nextTaskMaps = [
    ...agendaTasksForItem(item),
  ];

  try {
    creatingTaskItemId = item.id;
    renderAgendaItems(agenda);
    setActionStatus("Creating task...");
    const created = await directoryApi.createTask({
      title,
      ...(dueDateTimeUtc ? { dueDateTimeUtc } : {}),
    });
    const createdTask = created?.task || null;
    if (!createdTask?.provider || !createdTask?.externalTaskId) {
      throw new Error("Task was created but the link details were missing.");
    }

    nextTaskMaps.push({
      provider: createdTask.provider,
      externalTaskId: createdTask.externalTaskId,
      externalContainerId: createdTask.externalContainerId || "",
      title: createdTask.title || title,
      ownerEmail: currentUser?.email || "",
      linkedAt: new Date().toISOString(),
      dueDateTimeUtc: createdTask.dueDateTimeUtc || dueDateTimeUtc || null,
      isCompleted: false,
      source: "agenda_item",
    });

    await directoryApi.updateAgendaItem({
      itemId: item.id,
      taskMaps: nextTaskMaps,
    });
    taskComposerItemId = "";
    await refreshAgendas("Task added from agenda item.", { selectedAgendaId: agenda.id });
  } catch (error) {
    console.error(error);
    setActionStatus(error?.message || "Could not add task from agenda item.", true);
  } finally {
    creatingTaskItemId = "";
    const latestAgenda = selectedAgenda();
    if (latestAgenda?.id === agenda.id) {
      renderAgendaItems(latestAgenda);
    }
  }
}

async function markAgendaTaskComplete(agenda, item, task) {
  const key = taskLinkKey(task);
  if (!agenda?.id || !item?.id || !key) {
    return;
  }

  try {
    completingTaskKey = key;
    renderAgendaItems(agenda);
    setActionStatus("Completing task...");
    const updated = await directoryApi.updateTask({
      provider: task.provider,
      externalTaskId: task.externalTaskId,
      externalContainerId: task.externalContainerId || "",
      isCompleted: true,
    });
    const updatedTask = updated?.task || null;
    const nextTaskMaps = agendaTasksForItem(item).map((entry) => {
      if (taskLinkKey(entry) !== key) {
        return entry;
      }
      return {
        ...entry,
        isCompleted: true,
        title: updatedTask?.title || entry.title,
        dueDateTimeUtc: updatedTask?.dueDateTimeUtc || entry.dueDateTimeUtc || null,
        completedDateTimeUtc: updatedTask?.completedDateTimeUtc || new Date().toISOString(),
      };
    });
    await directoryApi.updateAgendaItem({
      itemId: item.id,
      taskMaps: nextTaskMaps,
    });
    await refreshAgendas("Task marked complete.", { selectedAgendaId: agenda.id });
  } catch (error) {
    console.error(error);
    setActionStatus(error?.message || "Could not mark task complete.", true);
  } finally {
    completingTaskKey = "";
    const latestAgenda = selectedAgenda();
    if (latestAgenda?.id === agenda.id) {
      renderAgendaItems(latestAgenda);
    }
  }
}

function renderAgendaDetail() {
  const agenda = selectedAgenda();
  agendaEmptyState.hidden = Boolean(agenda);
  agendaDetailWrap.hidden = !agenda;

  if (!agenda) {
    return;
  }

  const isOwner = normalizeEmail(agenda.ownerEmail) === normalizeEmail(currentUser?.email);
  agendaTitle.textContent = agenda.title || "Agenda";
  agendaMeta.textContent = agenda.isPrivate ? "Private" : "";
  agendaMeta.hidden = !agendaMeta.textContent;
  renderAgendaAttendees(agenda);
  agendaTitleEditInput.value = agenda.title || "";
  agendaPeopleSummaryInput.value = participantSummary(agenda);
  agendaPrivateEditInput.checked = agenda.isPrivate === true;
  agendaTitleEditInput.disabled = !isOwner;
  agendaPrivateEditInput.disabled = !isOwner;
  saveAgendaSettingsBtn.disabled = !isOwner;

  const detailLoaded = agendaDetailsById.has(agenda.id);
  const agendaItems = Array.isArray(agenda.items) ? agenda.items : [];

  if (selectedItemId && !agendaItems.some((item) => item.id === selectedItemId)) {
    selectedItemId = "";
  }
  populateItemForm(detailLoaded ? selectedItem() : null);

  if (!detailLoaded) {
    agendaItemsList.innerHTML = '<p class="muted scorecard-empty-state">Loading agenda items...</p>';
    agendaItemsEmpty.hidden = true;
    return;
  }

  renderAgendaItems(agenda);
}

function agendaSummaryCacheKey(email) {
  return `${AGENDA_SUMMARY_CACHE_PREFIX}:${normalizeEmail(email)}`;
}

function readCachedAgendaSummaries() {
  try {
    const raw = window.localStorage.getItem(agendaSummaryCacheKey(currentUser?.email || ""));
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed?.agendas) ? parsed.agendas : [];
  } catch {
    return [];
  }
}

function writeCachedAgendaSummaries() {
  try {
    window.localStorage.setItem(
      agendaSummaryCacheKey(currentUser?.email || ""),
      JSON.stringify({
        agendas,
        cachedAt: new Date().toISOString(),
      })
    );
  } catch {
    // Ignore cache write errors.
  }
}

function mergeAgendaSummaryWithDetail(agenda) {
  const detailed = agendaDetailsById.get(agenda?.id);
  if (!agenda) {
    return null;
  }
  if (!detailed) {
    logAgendaDebug("No detail available while merging agenda summary.", {
      agendaId: agenda?.id || "",
      summaryTitle: String(agenda?.title || ""),
    });
    return agenda;
  }
  const merged = {
    ...agenda,
    title: String(agenda.title || detailed.title || "").trim(),
    members: Array.isArray(agenda.members) && agenda.members.length ? agenda.members : detailed.members || [],
    participantEmails:
      Array.isArray(agenda.participantEmails) && agenda.participantEmails.length
        ? agenda.participantEmails
        : detailed.participantEmails || [],
    participantNames:
      Array.isArray(agenda.participantNames) && agenda.participantNames.length
        ? agenda.participantNames
        : detailed.participantNames || [],
  };
  logAgendaDebug("Merged agenda summary with detail.", {
    agendaId: agenda?.id || "",
    summaryTitle: String(agenda?.title || ""),
    detailTitle: String(detailed?.title || ""),
    mergedTitle: String(merged?.title || ""),
  });
  return merged;
}

function applyAgendaSummaries(nextAgendas) {
  agendas = (Array.isArray(nextAgendas) ? nextAgendas : []).map((agenda) => mergeAgendaSummaryWithDetail(agenda));
  if (!selectedAgendaId && agendas.length) {
    selectedAgendaId = agendas[0].id;
  }
  if (selectedAgendaId && !agendas.some((agenda) => agenda.id === selectedAgendaId)) {
    selectedAgendaId = agendas[0]?.id || "";
    selectedItemId = "";
  }
  const visibleIds = new Set(agendas.map((agenda) => agenda.id));
  agendaDetailsById = new Map([...agendaDetailsById.entries()].filter(([agendaId]) => visibleIds.has(agendaId)));
}

async function loadAgendaDetail(agendaId, options = {}) {
  const normalizedAgendaId = String(agendaId || "").trim();
  if (!normalizedAgendaId) {
    return null;
  }
  if (loadingAgendaDetails.has(normalizedAgendaId)) {
    return agendaDetailsById.get(normalizedAgendaId) || null;
  }
  if (!options.force && agendaDetailsById.has(normalizedAgendaId)) {
    return agendaDetailsById.get(normalizedAgendaId) || null;
  }

  loadingAgendaDetails.add(normalizedAgendaId);
  try {
    const payload = await directoryApi.getAgendaDetail({ agendaId: normalizedAgendaId });
    const agenda = payload?.agenda || null;
    logAgendaDebug("Loaded agenda detail payload.", {
      agendaId: normalizedAgendaId,
      title: String(agenda?.title || ""),
      members: Array.isArray(agenda?.members) ? agenda.members.map((member) => member.displayName || member.userEmail) : [],
    });
    if (agenda?.id) {
      agendaDetailsById.set(agenda.id, agenda);
      agendas = agendas.map((entry) => {
        if (entry.id !== agenda.id) {
          return entry;
        }
        return {
          ...entry,
          ...agenda,
          title: String(agenda.title || entry.title || "").trim(),
        };
      });
      writeCachedAgendaSummaries();
      renderAgendaList();
      if (agenda.id === selectedAgendaId) {
        renderAgendaDetail();
      }
    }
    return agenda;
  } catch (error) {
    console.error(error);
    if (normalizedAgendaId === selectedAgendaId) {
      setActionStatus(error?.message || "Could not load agenda details.", true);
    }
    return null;
  } finally {
    loadingAgendaDetails.delete(normalizedAgendaId);
  }
}

async function loadAgendas(status = "", options = {}) {
  const payload = await directoryApi.listAgendas({ summaryOnly: "true" });
  const nextAgendas = Array.isArray(payload?.agendas) ? payload.agendas : [];
  logAgendaDebug("Loaded agenda summaries payload.", nextAgendas.map((agenda) => ({
    id: agenda?.id || "",
    title: String(agenda?.title || ""),
    members: Array.isArray(agenda?.members) ? agenda.members.map((member) => member.displayName || member.userEmail) : [],
  })));
  if (options.selectedAgendaId) {
    selectedAgendaId = options.selectedAgendaId;
  }
  applyAgendaSummaries(nextAgendas);
  writeCachedAgendaSummaries();
  renderAgendaList();
  renderAgendaDetail();
  if (selectedAgendaId) {
    void loadAgendaDetail(selectedAgendaId, { force: options.forceSelectedDetail === true });
  }
  agendas.forEach((agenda) => {
    if (agenda.id && agenda.id !== selectedAgendaId) {
      void loadAgendaDetail(agenda.id);
    }
  });
  setStatus(status || `Loaded ${agendas.length} agenda${agendas.length === 1 ? "" : "s"}.`);
}

function hydrateAgendasFromCache() {
  const cachedAgendas = readCachedAgendaSummaries();
  if (!cachedAgendas.length) {
    return false;
  }
  applyAgendaSummaries(cachedAgendas);
  renderAgendaList();
  renderAgendaDetail();
  setStatus(`Loaded ${cachedAgendas.length} cached agenda${cachedAgendas.length === 1 ? "" : "s"}. Refreshing...`);
  agendas.forEach((agenda) => {
    if (agenda.id) {
      void loadAgendaDetail(agenda.id);
    }
  });
  return true;
}

async function refreshAgendas(status = "", options = {}) {
  await loadAgendas(status, options);
  if (selectedAgendaId) {
    await loadAgendaDetail(selectedAgendaId, { force: true });
  }
}

function editorCommand(command, value = null) {
  itemDetailEditor.focus();
  document.execCommand(command, false, value);
}

async function reorderItemToIndex(agenda, sourceItemId, insertIndex) {
  const items = filteredItems(agenda);
  const sourceIndex = items.findIndex((item) => item.id === sourceItemId);
  if (sourceIndex === -1) {
    return;
  }
  const remaining = items.filter((item) => item.id !== sourceItemId);
  const normalizedInsertIndex = Math.max(0, Math.min(insertIndex, remaining.length));
  if (normalizedInsertIndex === sourceIndex || normalizedInsertIndex === sourceIndex + 1) {
    return;
  }

  const previous = remaining[normalizedInsertIndex - 1] || null;
  const next = remaining[normalizedInsertIndex] || null;
  let sortOrder = 0;
  if (!previous && !next) {
    sortOrder = 100;
  } else if (!previous && next) {
    sortOrder = Number(next.sortOrder || 0) - 100;
  } else if (previous && next) {
    sortOrder = (Number(previous.sortOrder || 0) + Number(next.sortOrder || 0)) / 2;
  } else if (previous) {
    sortOrder = Number(previous.sortOrder || 0) + 100;
  }

  try {
    setActionStatus("Reordering item...");
    await directoryApi.updateAgendaItem({ itemId: sourceItemId, sortOrder });
    await refreshAgendas("Agenda order updated.", { selectedAgendaId: agenda.id });
  } catch (error) {
    console.error(error);
    setActionStatus(error?.message || "Could not reorder item.", true);
  }
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    currentUser = await directoryApi.getCurrentUser();
    const role = String(currentUser?.role || "").trim().toLowerCase();
    if (!canAccessPage(role, "agendas")) {
      window.location.href = "./unauthorized.html?page=agendas";
      return;
    }

    renderTopNavigation({ role });
    renderPeoplePicker();
    setCreatePanelExpanded(false);
    setAgendaSettingsExpanded(false);
    setAgendaItemComposerExpanded(false);
    setItemSearchExpanded(false);
    hydrateAgendasFromCache();
    await loadAgendas();
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not initialize agendas.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

agendaCreateForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const title = String(agendaTitleInput.value || "").trim();
  if (!title) {
    setStatus("Agenda title is required.", true);
    return;
  }

  const people = selectedPeopleFromCreateForm();
  try {
    setBusy(true);
    await directoryApi.createAgenda({
      title,
      agendaType: agendaTypeSelect.value || "meeting",
      participantEmails: people.map((person) => person.email),
      participantNames: people.map((person) => person.name),
      isPrivate: agendaPrivateInput.checked,
    });
    agendaCreateForm.reset();
    setCreatePanelExpanded(false);
    await refreshAgendas("Agenda created.");
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not create agenda.", true);
  } finally {
    setBusy(false);
  }
});

agendaSearchInput?.addEventListener("input", () => {
  renderAgendaList();
});

agendaSettingsForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const agenda = selectedAgenda();
  if (!agenda) {
    return;
  }

  try {
    setBusy(true);
    await directoryApi.updateAgenda({
      agendaId: agenda.id,
      title: String(agendaTitleEditInput.value || "").trim(),
      agendaType: agenda.agendaType,
      participantEmails: agenda.participantEmails,
      participantNames: agenda.participantNames,
      isPrivate: agendaPrivateEditInput.checked,
    });
    setAgendaSettingsExpanded(false);
    await refreshAgendas("Agenda settings saved.", { selectedAgendaId: agenda.id });
  } catch (error) {
    console.error(error);
    setActionStatus(error?.message || "Could not save agenda settings.", true);
  } finally {
    setBusy(false);
  }
});

cancelAgendaSettingsBtn?.addEventListener("click", () => {
  const agenda = selectedAgenda();
  if (!agenda) {
    return;
  }
  agendaTitleEditInput.value = agenda.title || "";
  agendaPeopleSummaryInput.value = participantSummary(agenda);
  agendaPrivateEditInput.checked = agenda.isPrivate === true;
  setAgendaSettingsExpanded(false);
});

agendaItemForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  const agenda = selectedAgenda();
  if (!agenda) {
    return;
  }

  const payload = {
    title: String(itemTitleInput.value || "").trim(),
    detailHtml: itemDetailEditor.innerHTML,
    stageTag: itemStageSelect.value || "",
    isPrivate: itemPrivateInput.checked,
    isUrgent: itemUrgentInput.checked,
    isImportant: itemImportantInput.checked,
  };

  if (!payload.title) {
    setActionStatus("Item title is required.", true);
    return;
  }

  try {
    setBusy(true);
    if (selectedItemId) {
      await directoryApi.updateAgendaItem({ itemId: selectedItemId, ...payload });
      await refreshAgendas("Agenda item updated.", { selectedAgendaId: agenda.id });
    } else {
      await directoryApi.createAgendaItem({ agendaId: agenda.id, ...payload });
      await refreshAgendas("Agenda item created.", { selectedAgendaId: agenda.id });
    }
    resetItemForm();
  } catch (error) {
    console.error(error);
    setActionStatus(error?.message || "Could not save agenda item.", true);
  } finally {
    setBusy(false);
  }
});

itemTitleInput?.addEventListener("input", () => {
  if (!agendaItemComposerExpanded && String(itemTitleInput.value || "").trim()) {
    setAgendaItemComposerExpanded(true);
  }
  if (agendaItemComposerExpanded && !selectedItemId && !hasComposerContent()) {
    setAgendaItemComposerExpanded(false);
  }
});

itemTitleInput?.addEventListener("focus", () => {
  setAgendaItemComposerExpanded(true);
});

itemDetailEditor?.addEventListener("focus", () => {
  setAgendaItemComposerExpanded(true);
});

itemSearchInput?.addEventListener("input", () => {
  renderAgendaDetail();
});

toggleItemSearchBtn?.addEventListener("click", () => {
  const nextExpanded = !itemSearchExpanded;
  setItemSearchExpanded(nextExpanded);
  if (!nextExpanded) {
    renderAgendaDetail();
  }
});

itemStageFilter?.addEventListener("change", () => {
  renderAgendaDetail();
});

document.querySelectorAll(".editor-btn[data-cmd]").forEach((button) => {
  button.addEventListener("click", () => {
    editorCommand(String(button.getAttribute("data-cmd") || ""));
  });
});

insertLinkBtn?.addEventListener("click", () => {
  const url = window.prompt("Enter a full link URL");
  if (url) {
    editorCommand("createLink", url);
  }
});

toggleCreatePanelBtn?.addEventListener("click", () => {
  setCreatePanelExpanded(!createPanelExpanded);
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
