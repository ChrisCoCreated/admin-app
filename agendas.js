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
];

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const actionStatus = document.getElementById("actionStatus");
const refreshBtn = document.getElementById("refreshBtn");
const quickAgendaButtons = document.getElementById("quickAgendaButtons");
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
const agendaSettingsForm = document.getElementById("agendaSettingsForm");
const agendaTitleEditInput = document.getElementById("agendaTitleEditInput");
const agendaPeopleSummaryInput = document.getElementById("agendaPeopleSummaryInput");
const agendaPrivateEditInput = document.getElementById("agendaPrivateEditInput");
const saveAgendaSettingsBtn = document.getElementById("saveAgendaSettingsBtn");
const itemSearchInput = document.getElementById("itemSearchInput");
const itemStageFilter = document.getElementById("itemStageFilter");
const agendaItemsList = document.getElementById("agendaItemsList");
const agendaItemsEmpty = document.getElementById("agendaItemsEmpty");
const agendaItemForm = document.getElementById("agendaItemForm");
const itemFormHeading = document.getElementById("itemFormHeading");
const itemTitleInput = document.getElementById("itemTitleInput");
const itemDetailEditor = document.getElementById("itemDetailEditor");
const itemStageSelect = document.getElementById("itemStageSelect");
const itemPrivateInput = document.getElementById("itemPrivateInput");
const itemUrgentInput = document.getElementById("itemUrgentInput");
const itemImportantInput = document.getElementById("itemImportantInput");
const saveAgendaItemBtn = document.getElementById("saveAgendaItemBtn");
const newAgendaItemBtn = document.getElementById("newAgendaItemBtn");
const insertLinkBtn = document.getElementById("insertLinkBtn");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let currentUser = null;
let agendas = [];
let selectedAgendaId = "";
let selectedItemId = "";
let busy = false;
let dragItemId = "";

function setBusy(value) {
  busy = value;
  refreshBtn.disabled = value;
  saveAgendaItemBtn.disabled = value;
  saveAgendaSettingsBtn.disabled = value;
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
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
  return agendas.find((agenda) => agenda.id === selectedAgendaId) || null;
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

function renderQuickCreate() {
  quickAgendaButtons.innerHTML = "";
  ORG_PEOPLE.forEach((person) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "secondary agenda-quick-card";
    button.innerHTML = `<strong>1:1 with ${escapeHtml(person.name)}</strong><span>${escapeHtml(person.email)}</span>`;
    button.addEventListener("click", async () => {
      try {
        setBusy(true);
        await directoryApi.createAgenda({
          title: `1:1 with ${person.name}`,
          agendaType: "one_to_one",
          participantEmails: [person.email],
          participantNames: [person.name],
          isPrivate: false,
        });
        await loadAgendas(`1:1 with ${person.name} created.`);
      } catch (error) {
        console.error(error);
        setStatus(error?.message || "Could not create 1:1 agenda.", true);
      } finally {
        setBusy(false);
      }
    });
    quickAgendaButtons.appendChild(button);
  });
}

function renderPeoplePicker() {
  agendaPeoplePicker.innerHTML = "";
  ORG_PEOPLE.forEach((person) => {
    const label = document.createElement("label");
    label.className = "agenda-person-option";
    label.innerHTML = `
      <input type="checkbox" value="${escapeHtml(person.email)}" />
      <span>${escapeHtml(person.name)}</span>
      <small>${escapeHtml(person.email)}</small>
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
  const query = String(itemSearchInput.value || "").trim().toLowerCase();
  const stage = String(itemStageFilter.value || "").trim().toLowerCase();
  return agenda.items.filter((item) => {
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

function renderAgendaList() {
  const list = filteredAgendas();
  agendaList.innerHTML = "";
  agendaListEmpty.hidden = list.length > 0;

  list.forEach((agenda) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "agenda-list-card";
    if (agenda.id === selectedAgendaId) {
      button.classList.add("is-selected");
    }
    const people = agenda.members.map((member) => member.displayName || displayNameForEmail(member.userEmail));
    button.innerHTML = `
      <strong>${escapeHtml(agenda.title)}</strong>
      <span>${escapeHtml(agenda.agendaType === "one_to_one" ? "1:1" : "Meeting")}${agenda.isPrivate ? " • Private" : ""}</span>
      <small>${escapeHtml(people.join(", ") || "Just you")}</small>
    `;
    button.addEventListener("click", () => {
      selectedAgendaId = agenda.id;
      selectedItemId = "";
      renderAgendaList();
      renderAgendaDetail();
    });
    agendaList.appendChild(button);
  });
}

function resetItemForm() {
  selectedItemId = "";
  itemFormHeading.textContent = "New item";
  itemTitleInput.value = "";
  itemDetailEditor.innerHTML = "<p></p>";
  itemStageSelect.value = "";
  itemPrivateInput.checked = false;
  itemUrgentInput.checked = false;
  itemImportantInput.checked = false;
}

function populateItemForm(item) {
  if (!item) {
    resetItemForm();
    return;
  }
  selectedItemId = item.id;
  itemFormHeading.textContent = "Edit item";
  itemTitleInput.value = item.title || "";
  itemDetailEditor.innerHTML = item.detailHtml || "<p></p>";
  itemStageSelect.value = item.stageTag || "";
  itemPrivateInput.checked = item.isPrivate === true;
  itemUrgentInput.checked = item.isUrgent === true;
  itemImportantInput.checked = item.isImportant === true;
}

function participantSummary(agenda) {
  return agenda.members
    .map((member) => member.displayName || displayNameForEmail(member.userEmail))
    .filter(Boolean)
    .join(", ");
}

function renderAgendaItems(agenda) {
  const items = filteredItems(agenda);
  agendaItemsList.innerHTML = "";
  agendaItemsEmpty.hidden = items.length > 0;

  items.forEach((item) => {
    const card = document.createElement("article");
    card.className = "agenda-item-card";
    card.draggable = true;
    if (item.id === selectedItemId) {
      card.classList.add("is-selected");
    }

    const badges = [
      item.stageTag ? `<span class="agenda-badge">${escapeHtml(item.stageTag)}</span>` : "",
      item.isUrgent ? '<span class="agenda-badge is-urgent">Urgent</span>' : "",
      item.isImportant ? '<span class="agenda-badge is-important">Important</span>' : "",
      item.isPrivate ? '<span class="agenda-badge">Private</span>' : "",
    ].join("");

    card.innerHTML = `
      <div class="agenda-item-card-head">
        <strong>${escapeHtml(item.title)}</strong>
        <span class="agenda-drag-handle">Drag</span>
      </div>
      <div class="agenda-item-badges">${badges}</div>
      <div class="agenda-item-preview">${item.detailHtml || "<p></p>"}</div>
    `;

    card.addEventListener("click", () => {
      populateItemForm(item);
      renderAgendaItems(agenda);
    });
    card.addEventListener("dragstart", () => {
      dragItemId = item.id;
    });
    card.addEventListener("dragover", (event) => {
      event.preventDefault();
      card.classList.add("is-drop-target");
    });
    card.addEventListener("dragleave", () => {
      card.classList.remove("is-drop-target");
    });
    card.addEventListener("drop", async (event) => {
      event.preventDefault();
      card.classList.remove("is-drop-target");
      if (!dragItemId || dragItemId === item.id) {
        return;
      }
      await reorderItem(agenda, dragItemId, item.id);
      dragItemId = "";
    });

    agendaItemsList.appendChild(card);
  });
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
  agendaMeta.textContent = `${agenda.agendaType === "one_to_one" ? "1:1" : "Meeting"} with ${participantSummary(agenda)}${agenda.isPrivate ? " • Private" : ""}`;
  agendaTitleEditInput.value = agenda.title || "";
  agendaPeopleSummaryInput.value = participantSummary(agenda);
  agendaPrivateEditInput.checked = agenda.isPrivate === true;
  agendaTitleEditInput.disabled = !isOwner;
  agendaPrivateEditInput.disabled = !isOwner;
  saveAgendaSettingsBtn.disabled = !isOwner;

  if (selectedItemId && !agenda.items.some((item) => item.id === selectedItemId)) {
    selectedItemId = "";
  }
  populateItemForm(selectedItem());
  renderAgendaItems(agenda);
}

async function loadAgendas(status = "") {
  const payload = await directoryApi.listAgendas();
  agendas = Array.isArray(payload?.agendas) ? payload.agendas : [];
  if (!selectedAgendaId && agendas.length) {
    selectedAgendaId = agendas[0].id;
  }
  if (selectedAgendaId && !agendas.some((agenda) => agenda.id === selectedAgendaId)) {
    selectedAgendaId = agendas[0]?.id || "";
  }
  renderAgendaList();
  renderAgendaDetail();
  setStatus(status || `Loaded ${agendas.length} agenda${agendas.length === 1 ? "" : "s"}.`);
}

function editorCommand(command, value = null) {
  itemDetailEditor.focus();
  document.execCommand(command, false, value);
}

async function reorderItem(agenda, sourceItemId, targetItemId) {
  const items = filteredItems(agenda);
  const remaining = items.filter((item) => item.id !== sourceItemId);
  const targetIndex = remaining.findIndex((item) => item.id === targetItemId);
  if (targetIndex === -1) {
    return;
  }
  const previous = remaining[targetIndex - 1] || null;
  const next = remaining[targetIndex] || null;
  let sortOrder = 0;
  if (!previous && next) {
    sortOrder = Number(next.sortOrder || 0) - 1;
  } else if (previous && next) {
    sortOrder = (Number(previous.sortOrder || 0) + Number(next.sortOrder || 0)) / 2;
  } else if (previous) {
    sortOrder = Number(previous.sortOrder || 0) + 1;
  }

  try {
    setActionStatus("Reordering item...");
    await directoryApi.updateAgendaItem({ itemId: sourceItemId, sortOrder });
    await loadAgendas("Agenda order updated.");
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
    renderQuickCreate();
    renderPeoplePicker();
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
    await loadAgendas("Agenda created.");
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
    await loadAgendas("Agenda settings saved.");
  } catch (error) {
    console.error(error);
    setActionStatus(error?.message || "Could not save agenda settings.", true);
  } finally {
    setBusy(false);
  }
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
      await loadAgendas("Agenda item updated.");
    } else {
      await directoryApi.createAgendaItem({ agendaId: agenda.id, ...payload });
      await loadAgendas("Agenda item created.");
    }
    resetItemForm();
  } catch (error) {
    console.error(error);
    setActionStatus(error?.message || "Could not save agenda item.", true);
  } finally {
    setBusy(false);
  }
});

newAgendaItemBtn?.addEventListener("click", () => {
  resetItemForm();
  renderAgendaDetail();
});

itemSearchInput?.addEventListener("input", () => {
  renderAgendaDetail();
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

refreshBtn?.addEventListener("click", async () => {
  try {
    setBusy(true);
    await loadAgendas("Agendas refreshed.");
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not refresh agendas.", true);
  } finally {
    setBusy(false);
  }
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
