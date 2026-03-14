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
const loadingAgendaDetails = new Set();

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

function agendaDisplayTitle(agenda) {
  const detailedAgenda = agendaDetailsById.get(agenda?.id);
  const title = String(detailedAgenda?.title || agenda?.title || "").trim();
  if (title) {
    return title;
  }
  const people = Array.isArray(agenda?.members)
    ? agenda.members.map((member) => member.displayName || displayNameForEmail(member.userEmail)).filter(Boolean)
    : [];
  return people.join(", ") || "Agenda";
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
    const people = agenda.members.map((member) => member.displayName || displayNameForEmail(member.userEmail));
    card.innerHTML = `
      <div class="agenda-list-card-top">
        <button type="button" class="agenda-list-card-main">
          <strong>${escapeHtml(agendaDisplayTitle(agenda))}</strong>
          <span>${agenda.isPrivate ? "Private" : ""}</span>
          <small>${escapeHtml(people.join(", ") || "Just you")}</small>
        </button>
        ${isOwner ? '<button type="button" class="ghost subtle agenda-card-edit-btn">Edit</button>' : ""}
      </div>
    `;
    card.querySelector(".agenda-list-card-main")?.addEventListener("click", () => {
      void openAgenda(agenda.id);
    });
    card.querySelector(".agenda-card-edit-btn")?.addEventListener("click", (event) => {
      event.stopPropagation();
      void openAgenda(agenda.id, { editSettings: true });
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

    const badges = [
      item.stageTag ? `<span class="agenda-badge">${escapeHtml(item.stageTag)}</span>` : "",
      item.isUrgent ? '<span class="agenda-badge is-urgent">Urgent</span>' : "",
      item.isImportant ? '<span class="agenda-badge is-important">Important</span>' : "",
      item.isPrivate ? '<span class="agenda-badge">Private</span>' : "",
    ].join("");
    const previewHtml = itemDetailsExpanded ? item.detailHtml || "<p></p>" : "";

    card.innerHTML = `
      <div class="agenda-item-card-head">
        <strong>${escapeHtml(item.title)}</strong>
        <button
          type="button"
          class="agenda-expand-toggle"
          aria-label="${itemDetailsExpanded ? "Hide item details" : "Show item details"}"
          aria-expanded="${itemDetailsExpanded ? "true" : "false"}"
        >
          <span aria-hidden="true">${itemDetailsExpanded ? "⌃" : "⌄"}</span>
        </button>
      </div>
      <div class="agenda-item-badges">${badges}</div>
      ${itemDetailsExpanded ? `<div class="agenda-item-preview">${previewHtml}</div>` : ""}
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

    agendaItemsList.appendChild(card);
    agendaItemsList.appendChild(buildDropZone(index + 1));
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

function applyAgendaSummaries(nextAgendas) {
  agendas = Array.isArray(nextAgendas) ? nextAgendas : [];
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
    if (agenda?.id) {
      agendaDetailsById.set(agenda.id, agenda);
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
  if (selectedAgendaId) {
    void loadAgendaDetail(selectedAgendaId);
  }
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
