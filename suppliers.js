import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260317";
import {
  buildFieldMaps,
  createSharePointApi,
  fieldByTitle as sharedFieldByTitle,
  fieldByTitleWritable as sharedFieldByTitleWritable,
} from "./sharepoint-list-utils.js";

const FIXED_TENANT_ID = FRONTEND_CONFIG.tenantId;
const FIXED_CLIENT_ID = FRONTEND_CONFIG.spaClientId;
const FIXED_SITE_URL =
  FRONTEND_CONFIG.sharePoint?.wellbeingSiteUrl || "https://planwithcare.sharepoint.com/sites/Wellbeing";
const SUPPLIERS_LIST_PATH =
  FRONTEND_CONFIG.sharePoint?.suppliersListPath || "/sites/Wellbeing/Lists/Suppliers Database";
const SHAREPOINT_LIST_URL = `${FIXED_SITE_URL}/Lists/Suppliers%20Database?env=WebViewList`;
const DEFAULT_TAG_OPTIONS = ["sport", "travel", "fashion", "identity", "fun", "play", "music", "religion"];

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const standardModeBtn = document.getElementById("standardModeBtn");
const eeModeBtn = document.getElementById("eeModeBtn");
const refreshSuppliersBtn = document.getElementById("refreshSuppliersBtn");
const supplierSearchInput = document.getElementById("supplierSearchInput");
const supplierTypeSelect = document.getElementById("supplierTypeSelect");
const supplierCountySelect = document.getElementById("supplierCountySelect");
const supplierTownSelect = document.getElementById("supplierTownSelect");
const supplierRatingSelect = document.getElementById("supplierRatingSelect");
const supplierTagSelect = document.getElementById("supplierTagSelect");
const supplierWebsiteOnly = document.getElementById("supplierWebsiteOnly");
const supplierQuickFilters = document.getElementById("supplierQuickFilters");
const supplierStatsGrid = document.getElementById("supplierStatsGrid");
const resultsHeading = document.getElementById("resultsHeading");
const resultsSummary = document.getElementById("resultsSummary");
const suppliersResultsGrid = document.getElementById("suppliersResultsGrid");
const clearFiltersBtn = document.getElementById("clearFiltersBtn");
const emptyState = document.getElementById("emptyState");
const quickEntryIntro = document.getElementById("quickEntryIntro");
const quickEntryStatus = document.getElementById("quickEntryStatus");
const quickEntryModeBadge = document.getElementById("quickEntryModeBadge");
const quickEntryPreview = document.getElementById("quickEntryPreview");
const quickEntryForm = document.getElementById("quickEntryForm");
const entryTitleInput = document.getElementById("entryTitleInput");
const entryTypeInput = document.getElementById("entryTypeInput");
const entryTownInput = document.getElementById("entryTownInput");
const entryCountyInput = document.getElementById("entryCountyInput");
const entryWebsiteInput = document.getElementById("entryWebsiteInput");
const entryRatingInput = document.getElementById("entryRatingInput");
const entryTagInput = document.getElementById("entryTagInput");
const entryContactInput = document.getElementById("entryContactInput");
const entryNotesInput = document.getElementById("entryNotesInput");
const saveEntryBtn = document.getElementById("saveEntryBtn");
const copyEntrySummaryBtn = document.getElementById("copyEntrySummaryBtn");
const copyEntryCsvBtn = document.getElementById("copyEntryCsvBtn");
const resetEntryBtn = document.getElementById("resetEntryBtn");
const supplierTypeSuggestions = document.getElementById("supplierTypeSuggestions");
const supplierTagSuggestions = document.getElementById("supplierTagSuggestions");

const authController = createAuthController({
  tenantId: FIXED_TENANT_ID,
  clientId: FIXED_CLIENT_ID,
  onSignedOut: () => {
    spApi?.clearCaches();
    spApi = null;
    listInfo = null;
    fieldMap = null;
    allSuppliers = [];
  },
});
const directoryApi = createDirectoryApi(authController);

let spApi = null;
let listInfo = null;
let fieldMap = null;
let currentMode = "ee";
let activeCategoryFilter = "";
let allSuppliers = [];
let currentRole = "";
let listChoices = {
  supplierTypes: [],
  tags: [...DEFAULT_TAG_OPTIONS],
};
let suppliersLoaded = false;
let loadingSuppliers = false;
let savingSupplier = false;

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "suppliers").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function getConfig() {
  const siteUrl = new URL(FIXED_SITE_URL);
  return {
    siteUrl: siteUrl.origin + siteUrl.pathname.replace(/\/$/, ""),
    siteHost: siteUrl.origin,
  };
}

async function getSharePointToken(config) {
  return authController.acquireSharePointToken(config.siteHost);
}

function getSpApi(config) {
  if (!spApi) {
    spApi = createSharePointApi({
      siteUrl: config.siteUrl,
      getToken: () => getSharePointToken(config),
    });
  }
  return spApi;
}

function fieldByTitle(candidates) {
  return sharedFieldByTitle(fieldMap, candidates);
}

function fieldByTitleWritable(candidates) {
  return sharedFieldByTitleWritable(fieldMap, candidates);
}

function fieldInternalName(candidates, options = {}) {
  const field = options.writable ? fieldByTitleWritable(candidates) : fieldByTitle(candidates);
  return field?.InternalName || null;
}

function resolveFieldAliases() {
  return {
    title: fieldInternalName(["Title"], { writable: true }) || "Title",
    supplierType: fieldInternalName(["Supplier Type"], { writable: true }),
    notes: fieldInternalName(["Notes"], { writable: true }),
    website: fieldInternalName(["Website"], { writable: true }),
    town: fieldInternalName(["Town"], { writable: true }),
    county: fieldInternalName(["County"], { writable: true }),
    contactDetails: fieldInternalName(["Contact Details"], { writable: true }),
    rating: fieldInternalName(["Rating"], { writable: true }),
    ee: fieldInternalName(["EE"], { writable: true }),
    tags: fieldInternalName(["Tags"], { writable: true }),
  };
}

function normalizeValue(value) {
  return String(value || "").trim();
}

function normalizeKey(value) {
  return normalizeValue(value).toLowerCase();
}

function normalizeTagList(value) {
  if (Array.isArray(value)) {
    return Array.from(
      new Set(
        value
          .map((tag) => normalizeValue(tag))
          .filter(Boolean)
      )
    );
  }

  if (value && typeof value === "object" && Array.isArray(value.results)) {
    return normalizeTagList(value.results);
  }

  const normalized = normalizeValue(value);
  if (!normalized) {
    return [];
  }

  return Array.from(
    new Set(
      normalized
        .split(/[;,]/)
        .map((tag) => normalizeValue(tag))
        .filter(Boolean)
    )
  );
}

function formatLabel(value, fallback = "Unspecified") {
  return normalizeValue(value) || fallback;
}

function formatTagLabel(value) {
  return normalizeValue(value) || "No tag";
}

function optionValues(items, fieldName) {
  return Array.from(
    new Set(
      items
        .map((item) => normalizeValue(item[fieldName]))
        .filter(Boolean)
    )
  ).sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));
}

function getAllTagOptions(items = allSuppliers) {
  const tags = new Set(listChoices.tags || DEFAULT_TAG_OPTIONS);
  for (const item of items) {
    for (const tag of normalizeTagList(item.tags)) {
      tags.add(tag);
    }
  }
  return Array.from(tags).sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));
}

function getModeLabel(mode = currentMode) {
  return mode === "ee" ? "EE experiences" : "standard suppliers";
}

function getModeItems(mode = currentMode) {
  return allSuppliers.filter((item) => Boolean(item.isEe) === (mode === "ee"));
}

function categoryCounts(items) {
  const counts = new Map();
  for (const item of items) {
    const key = normalizeValue(item.supplierType) || "Unspecified";
    counts.set(key, (counts.get(key) || 0) + 1);
  }
  return Array.from(counts.entries()).sort((left, right) => {
    if (right[1] !== left[1]) {
      return right[1] - left[1];
    }
    return left[0].localeCompare(right[0], undefined, { sensitivity: "base" });
  });
}

function populateSelect(select, values, label) {
  if (!select) {
    return;
  }

  const previousValue = select.value;
  select.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "";
  allOption.textContent = `All ${label}`;
  select.appendChild(allOption);

  for (const value of values) {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  }

  if (values.includes(previousValue)) {
    select.value = previousValue;
  }
}

function updateFilterOptions() {
  const modeItems = getModeItems();
  const supplierTypes = Array.from(
    new Set([...(listChoices.supplierTypes || []), ...optionValues(modeItems, "supplierType")])
  ).sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" }));

  populateSelect(supplierTypeSelect, supplierTypes, "categories");
  populateSelect(supplierCountySelect, optionValues(modeItems, "county"), "counties");
  populateSelect(supplierTownSelect, optionValues(modeItems, "town"), "towns");
  populateSelect(supplierTagSelect, getAllTagOptions(modeItems), "tags");

  const availableCategories = new Set(supplierTypes);
  if (activeCategoryFilter && !availableCategories.has(activeCategoryFilter)) {
    activeCategoryFilter = "";
  }

  supplierTypeSuggestions.innerHTML = "";
  for (const category of supplierTypes) {
    const option = document.createElement("option");
    option.value = category;
    supplierTypeSuggestions.appendChild(option);
  }

  supplierTagSuggestions.innerHTML = "";
  for (const tag of getAllTagOptions(modeItems)) {
    const option = document.createElement("option");
    option.value = tag;
    supplierTagSuggestions.appendChild(option);
  }
}

function normalizeWebsiteValue(value) {
  if (!value) {
    return "";
  }
  if (typeof value === "string") {
    return normalizeValue(value);
  }
  if (typeof value === "object") {
    return normalizeValue(value.Url || value.url || value.Description || value.description);
  }
  return "";
}

function normalizeBooleanValue(value) {
  if (typeof value === "boolean") {
    return value;
  }
  const normalized = normalizeValue(value).toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes";
}

function normalizeSupplierItem(item, aliases) {
  return {
    id: Number(item?.Id || 0),
    title: normalizeValue(item?.[aliases.title]),
    supplierType: normalizeValue(item?.[aliases.supplierType]),
    notes: normalizeValue(item?.[aliases.notes]),
    website: normalizeWebsiteValue(item?.[aliases.website]),
    town: normalizeValue(item?.[aliases.town]),
    county: normalizeValue(item?.[aliases.county]),
    contactDetails: normalizeValue(item?.[aliases.contactDetails]),
    rating: normalizeValue(item?.[aliases.rating]),
    isEe: normalizeBooleanValue(item?.[aliases.ee]),
    tags: normalizeTagList(item?.[aliases.tags]),
    modified: normalizeValue(item?.Modified),
  };
}

function matchesFilters(item) {
  const searchTerm = normalizeKey(supplierSearchInput?.value);
  const typeValue = normalizeValue(supplierTypeSelect?.value);
  const countyValue = normalizeValue(supplierCountySelect?.value);
  const townValue = normalizeValue(supplierTownSelect?.value);
  const ratingValue = normalizeValue(supplierRatingSelect?.value);
  const tagValue = normalizeValue(supplierTagSelect?.value);
  const websiteOnly = Boolean(supplierWebsiteOnly?.checked);
  const categoryValue = activeCategoryFilter || typeValue;

  if (categoryValue && normalizeValue(item.supplierType) !== categoryValue) {
    return false;
  }
  if (countyValue && normalizeValue(item.county) !== countyValue) {
    return false;
  }
  if (townValue && normalizeValue(item.town) !== townValue) {
    return false;
  }
  if (ratingValue === "unrated" && normalizeValue(item.rating)) {
    return false;
  }
  if (ratingValue && ratingValue !== "unrated" && normalizeValue(item.rating) !== ratingValue) {
    return false;
  }
  if (tagValue && !normalizeTagList(item.tags).includes(tagValue)) {
    return false;
  }
  if (websiteOnly && !normalizeValue(item.website)) {
    return false;
  }
  if (!searchTerm) {
    return true;
  }

  const haystack = [
    item.title,
    item.supplierType,
    item.town,
    item.county,
    item.notes,
    item.contactDetails,
    item.website,
    normalizeTagList(item.tags).join("\n"),
  ]
    .map((value) => normalizeKey(value))
    .join("\n");

  return haystack.includes(searchTerm);
}

function filteredItems() {
  return getModeItems().filter(matchesFilters);
}

function buildStatCard(label, value, accentClass = "") {
  const card = document.createElement("article");
  card.className = `suppliers-stat-card ${accentClass}`.trim();
  const heading = document.createElement("h3");
  heading.textContent = label;
  const amount = document.createElement("p");
  amount.className = "suppliers-stat-value";
  amount.textContent = String(value);
  card.append(heading, amount);
  return card;
}

function renderStats(items) {
  if (!supplierStatsGrid) {
    return;
  }

  const modeItems = getModeItems();
  const recommendedCount = items.filter((item) => item.rating === "👍").length;
  const websiteCount = items.filter((item) => normalizeValue(item.website)).length;
  const topCategory = categoryCounts(modeItems)[0];
  const taggedCount = items.filter((item) => normalizeTagList(item.tags).length > 0).length;

  supplierStatsGrid.innerHTML = "";
  supplierStatsGrid.append(
    buildStatCard("In this mode", modeItems.length, currentMode === "ee" ? "is-ee" : ""),
    buildStatCard("Matching filters", items.length),
    buildStatCard("With websites", websiteCount),
    buildStatCard("Tagged", taggedCount),
    buildStatCard("Recommended", recommendedCount, "is-positive"),
    buildStatCard("Top category", topCategory ? `${topCategory[0]} (${topCategory[1]})` : "None")
  );
}

function renderQuickFilters() {
  if (!supplierQuickFilters) {
    return;
  }

  const topCategories = categoryCounts(getModeItems()).slice(0, 6);
  supplierQuickFilters.innerHTML = "";

  const allButton = document.createElement("button");
  allButton.type = "button";
  allButton.className = "status-filter-btn";
  if (!activeCategoryFilter) {
    allButton.classList.add("active");
  }
  allButton.textContent = "All";
  allButton.addEventListener("click", () => {
    activeCategoryFilter = "";
    supplierTypeSelect.value = "";
    renderPage();
  });
  supplierQuickFilters.appendChild(allButton);

  for (const [category, count] of topCategories) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "status-filter-btn";
    if (activeCategoryFilter === category) {
      button.classList.add("active");
    }
    button.textContent = `${category} (${count})`;
    button.addEventListener("click", () => {
      activeCategoryFilter = activeCategoryFilter === category ? "" : category;
      supplierTypeSelect.value = activeCategoryFilter;
      renderPage();
    });
    supplierQuickFilters.appendChild(button);
  }
}

function locationLabel(item) {
  const town = normalizeValue(item.town);
  const county = normalizeValue(item.county);
  if (town && county) {
    return `${town}, ${county}`;
  }
  return town || county || "Location not added";
}

function excerpt(value, fallback = "No notes added yet.") {
  const trimmed = normalizeValue(value);
  if (!trimmed) {
    return fallback;
  }
  if (trimmed.length <= 180) {
    return trimmed;
  }
  return `${trimmed.slice(0, 177).trimEnd()}...`;
}

function getSupplierEditUrl(itemId) {
  if (!Number.isFinite(Number(itemId)) || Number(itemId) <= 0) {
    return SHAREPOINT_LIST_URL;
  }
  const relative = `${SUPPLIERS_LIST_PATH}/EditForm.aspx?ID=${encodeURIComponent(String(itemId))}`;
  return new URL(relative, FIXED_SITE_URL).toString();
}

async function copyText(text, successMessage) {
  if (!text) {
    setStatus("Nothing to copy yet.", true);
    return;
  }

  try {
    if (navigator?.clipboard?.writeText) {
      await navigator.clipboard.writeText(text);
    } else {
      const input = document.createElement("textarea");
      input.value = text;
      input.setAttribute("readonly", "");
      input.style.position = "fixed";
      input.style.opacity = "0";
      document.body.appendChild(input);
      input.select();
      document.execCommand("copy");
      document.body.removeChild(input);
    }
    setStatus(successMessage);
  } catch (error) {
    setStatus(error?.message || "Copy failed.", true);
  }
}

function supplierCard(item) {
  const card = document.createElement("article");
  card.className = `supplier-card ${item.isEe ? "is-ee" : "is-standard"}`;

  const header = document.createElement("div");
  header.className = "supplier-card-head";

  const titleWrap = document.createElement("div");
  const title = document.createElement("h3");
  title.textContent = item.title;
  const meta = document.createElement("p");
  meta.className = "muted";
  meta.textContent = locationLabel(item);
  titleWrap.append(title, meta);

  const rating = document.createElement("span");
  rating.className = `supplier-rating-pill ${normalizeValue(item.rating) ? "" : "is-unrated"}`.trim();
  rating.textContent = normalizeValue(item.rating) || "Unrated";

  header.append(titleWrap, rating);

  const chips = document.createElement("div");
  chips.className = "supplier-chip-row";

  const typeChip = document.createElement("span");
  typeChip.className = "supplier-chip";
  typeChip.textContent = formatLabel(item.supplierType);
  chips.appendChild(typeChip);

  const modeChip = document.createElement("span");
  modeChip.className = `supplier-chip ${item.isEe ? "ee-flag" : "standard-flag"}`;
  modeChip.textContent = item.isEe ? "Enhance & Exceed" : "Standard";
  chips.appendChild(modeChip);

  for (const tag of normalizeTagList(item.tags)) {
    const tagChip = document.createElement("span");
    tagChip.className = "supplier-chip supplier-tag-chip";
    tagChip.textContent = `#${tag}`;
    chips.appendChild(tagChip);
  }

  const notes = document.createElement("p");
  notes.className = "supplier-copy";
  notes.textContent = excerpt(item.notes);

  const contact = document.createElement("p");
  contact.className = "supplier-contact";
  contact.textContent = excerpt(item.contactDetails, "No contact details added.");

  const actions = document.createElement("div");
  actions.className = "supplier-actions";

  if (normalizeValue(item.website)) {
    const websiteLink = document.createElement("a");
    websiteLink.className = "button-link";
    websiteLink.href = item.website;
    websiteLink.target = "_blank";
    websiteLink.rel = "noopener noreferrer";
    websiteLink.textContent = item.isEe ? "Open experience" : "Open website";
    actions.appendChild(websiteLink);
  }

  const copyContactBtn = document.createElement("button");
  copyContactBtn.className = "secondary";
  copyContactBtn.type = "button";
  copyContactBtn.textContent = "Copy contact";
  copyContactBtn.addEventListener("click", () => {
    void copyText(
      [item.title, item.contactDetails, item.website].filter(Boolean).join("\n"),
      `Copied details for ${item.title}.`
    );
  });
  actions.appendChild(copyContactBtn);

  const sharepointLink = document.createElement("a");
  sharepointLink.className = "secondary supplier-inline-link";
  sharepointLink.href = getSupplierEditUrl(item.id);
  sharepointLink.target = "_blank";
  sharepointLink.rel = "noopener noreferrer";
  sharepointLink.textContent = "Edit in SharePoint";
  actions.appendChild(sharepointLink);

  card.append(header, chips, notes, contact, actions);
  return card;
}

function renderResults(items) {
  suppliersResultsGrid.innerHTML = "";
  for (const item of items) {
    suppliersResultsGrid.appendChild(supplierCard(item));
  }
  emptyState.hidden = items.length > 0;
}

function renderHeading(items) {
  const label = getModeLabel();
  resultsHeading.textContent = currentMode === "ee" ? "EE Directory" : "Supplier Directory";
  if (!suppliersLoaded) {
    resultsSummary.textContent = "Loading live SharePoint records...";
    return;
  }
  resultsSummary.textContent = `${items.length} ${label} shown from ${getModeItems().length} live SharePoint records.`;
}

function quickEntryPayload() {
  return {
    title: normalizeValue(entryTitleInput.value),
    supplierType: normalizeValue(entryTypeInput.value),
    town: normalizeValue(entryTownInput.value),
    county: normalizeValue(entryCountyInput.value),
    website: normalizeValue(entryWebsiteInput.value),
    rating: normalizeValue(entryRatingInput.value),
    tag: normalizeValue(entryTagInput.value),
    contactDetails: normalizeValue(entryContactInput.value),
    notes: normalizeValue(entryNotesInput.value),
    isEe: currentMode === "ee",
  };
}

function buildQuickEntryPreviewText() {
  const payload = quickEntryPayload();
  return [
    `Name: ${payload.title || "-"}`,
    `Category: ${payload.supplierType || "-"}`,
    `Tag: ${formatTagLabel(payload.tag)}`,
    `Town: ${payload.town || "-"}`,
    `County: ${payload.county || "-"}`,
    `Website: ${payload.website || "-"}`,
    `Rating: ${payload.rating || "-"}`,
    `EE: ${payload.isEe ? "Yes" : "No"}`,
    "Contact Details:",
    payload.contactDetails || "-",
    "Notes:",
    payload.notes || "-",
  ].join("\n");
}

function escapeCsvValue(value) {
  return `"${String(value || "").replace(/"/g, '""')}"`;
}

function buildQuickEntryCsv() {
  const payload = quickEntryPayload();
  return [
    escapeCsvValue(payload.title),
    escapeCsvValue(payload.supplierType),
    escapeCsvValue(payload.notes),
    escapeCsvValue(payload.website),
    escapeCsvValue(payload.town),
    escapeCsvValue(payload.county),
    escapeCsvValue(payload.contactDetails),
    escapeCsvValue(payload.rating),
    escapeCsvValue(payload.isEe ? "True" : "False"),
    escapeCsvValue(payload.tag),
  ].join(",");
}

function renderQuickEntryPreview() {
  const isEeMode = currentMode === "ee";
  document.body.classList.toggle("is-ee-mode", isEeMode);
  quickEntryModeBadge.textContent = isEeMode ? "Enhance & Exceed" : "Standard supplier";
  quickEntryModeBadge.className = `supplier-chip ${isEeMode ? "ee-flag" : "standard-flag"}`;
  quickEntryIntro.textContent = isEeMode
    ? "Create a new EE experience or resource in SharePoint from here."
    : "Create a new standard supplier entry in SharePoint from here.";
  quickEntryStatus.textContent = isEeMode
    ? savingSupplier
      ? "Saving this EE entry to SharePoint..."
      : "This entry will be saved with EE set to True."
    : savingSupplier
      ? "Saving this standard supplier to SharePoint..."
      : "This entry will be saved with EE set to False.";
  quickEntryPreview.textContent = buildQuickEntryPreviewText();
}

function renderModeButtons() {
  const eeMode = currentMode === "ee";
  standardModeBtn.classList.toggle("active", !eeMode);
  eeModeBtn.classList.toggle("active", eeMode);
}

function renderPage() {
  renderModeButtons();
  updateFilterOptions();
  renderQuickFilters();
  renderQuickEntryPreview();
  const items = filteredItems();
  renderHeading(items);
  renderStats(items);
  renderResults(items);
}

function setBusyState() {
  if (refreshSuppliersBtn) {
    refreshSuppliersBtn.disabled = loadingSuppliers || savingSupplier;
  }
  if (saveEntryBtn) {
    saveEntryBtn.disabled = savingSupplier || loadingSuppliers;
  }
}

function setMode(mode) {
  currentMode = mode === "ee" ? "ee" : "standard";
  activeCategoryFilter = "";
  supplierTypeSelect.value = "";
  supplierCountySelect.value = "";
  supplierTownSelect.value = "";
  supplierRatingSelect.value = "";
  supplierTagSelect.value = "";
  supplierWebsiteOnly.checked = false;
  renderPage();
}

function resetFilters() {
  supplierSearchInput.value = "";
  supplierTypeSelect.value = "";
  supplierCountySelect.value = "";
  supplierTownSelect.value = "";
  supplierRatingSelect.value = "";
  supplierTagSelect.value = "";
  supplierWebsiteOnly.checked = false;
  activeCategoryFilter = "";
  renderPage();
}

async function fetchCurrentUser() {
  return directoryApi.getCurrentUser();
}

async function loadListInfo() {
  const config = getConfig();
  const api = getSpApi(config);
  listInfo = await api.resolveListByPath(SUPPLIERS_LIST_PATH);
  const fields = await api.getListFields(listInfo.Id);
  fieldMap = buildFieldMaps(fields);

  const supplierTypeField = fieldByTitle(["Supplier Type"]);
  const tagsField = fieldByTitle(["Tags"]);
  listChoices = {
    supplierTypes: Array.isArray(supplierTypeField?.Choices?.results)
      ? supplierTypeField.Choices.results.map((entry) => normalizeValue(entry)).filter(Boolean)
      : [],
    tags: Array.from(
      new Set([
        ...DEFAULT_TAG_OPTIONS,
        ...(Array.isArray(tagsField?.Choices?.results)
          ? tagsField.Choices.results.map((entry) => normalizeValue(entry)).filter(Boolean)
          : []),
      ])
    ).sort((left, right) => left.localeCompare(right, undefined, { sensitivity: "base" })),
  };
}

async function loadSuppliers() {
  loadingSuppliers = true;
  setBusyState();
  setStatus("Loading live suppliers from SharePoint...");

  try {
    if (!listInfo || !fieldMap) {
      await loadListInfo();
    }

    const config = getConfig();
    const api = getSpApi(config);
    const aliases = resolveFieldAliases();
    const selectFields = [
      "Id",
      "Modified",
      aliases.title,
      aliases.supplierType,
      aliases.notes,
      aliases.website,
      aliases.town,
      aliases.county,
      aliases.contactDetails,
      aliases.rating,
      aliases.ee,
      aliases.tags,
    ].filter(Boolean);

    const query = Array.from(new Set(selectFields)).join(",");
    const data = await api.request(
      `/_api/web/lists(guid'${listInfo.Id}')/items?$top=5000&$orderby=Modified desc&$select=${encodeURIComponent(query)}`
    );

    allSuppliers = Array.isArray(data?.d?.results)
      ? data.d.results
          .map((item) => normalizeSupplierItem(item, aliases))
          .filter((item) => item.title)
      : [];

    suppliersLoaded = true;
    setStatus(`Loaded ${allSuppliers.length} supplier records from SharePoint.`);
    renderPage();
  } catch (error) {
    console.error(error);
    suppliersLoaded = false;
    allSuppliers = [];
    renderPage();
    setStatus(error?.message || "Could not load suppliers from SharePoint.", true);
  } finally {
    loadingSuppliers = false;
    setBusyState();
  }
}

function buildWebsitePayload(url, title) {
  const normalizedUrl = normalizeValue(url);
  if (!normalizedUrl) {
    return null;
  }

  return {
    __metadata: { type: "SP.FieldUrlValue" },
    Url: normalizedUrl,
    Description: normalizeValue(title) || normalizedUrl,
  };
}

async function createSupplierInSharePoint() {
  const payload = quickEntryPayload();
  if (!payload.title) {
    entryTitleInput.focus();
    throw new Error("Name is required before saving.");
  }

  if (!listInfo || !fieldMap) {
    await loadListInfo();
  }

  const aliases = resolveFieldAliases();
  const requestBody = {
    __metadata: { type: listInfo.ListItemEntityTypeFullName },
  };

  requestBody[aliases.title] = payload.title;
  if (aliases.supplierType && payload.supplierType) {
    requestBody[aliases.supplierType] = payload.supplierType;
  }
  if (aliases.notes && payload.notes) {
    requestBody[aliases.notes] = payload.notes;
  }
  if (aliases.website && payload.website) {
    requestBody[aliases.website] = buildWebsitePayload(payload.website, payload.title);
  }
  if (aliases.town && payload.town) {
    requestBody[aliases.town] = payload.town;
  }
  if (aliases.county && payload.county) {
    requestBody[aliases.county] = payload.county;
  }
  if (aliases.contactDetails && payload.contactDetails) {
    requestBody[aliases.contactDetails] = payload.contactDetails;
  }
  if (aliases.rating && payload.rating) {
    requestBody[aliases.rating] = payload.rating;
  }
  if (aliases.ee) {
    requestBody[aliases.ee] = Boolean(payload.isEe);
  }
  if (aliases.tags && payload.tag) {
    requestBody[aliases.tags] = payload.tag;
  }

  const config = getConfig();
  const api = getSpApi(config);
  const result = await api.createListItem(listInfo.Id, requestBody);
  return Number(result?.d?.Id || 0);
}

function resetQuickEntry() {
  quickEntryForm.reset();
  renderQuickEntryPreview();
}

async function handleSaveSupplier() {
  savingSupplier = true;
  setBusyState();
  renderQuickEntryPreview();

  try {
    setStatus("Saving supplier to SharePoint...");
    const itemId = await createSupplierInSharePoint();
    await loadSuppliers();
    resetQuickEntry();
    setStatus(itemId ? `Saved to SharePoint as item #${itemId}.` : "Saved to SharePoint.");
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Could not save supplier to SharePoint.", true);
  } finally {
    savingSupplier = false;
    setBusyState();
    renderQuickEntryPreview();
  }
}

async function init() {
  try {
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await fetchCurrentUser();
    currentRole = String(profile?.role || "").trim().toLowerCase();
    if (!canAccessPage(currentRole, "suppliers")) {
      redirectToUnauthorized("suppliers");
      return;
    }

    renderTopNavigation({ role: currentRole });
    const email = String(profile?.email || "").trim();
    setStatus(email ? `Signed in as ${email}. Loading live suppliers...` : "Signed in. Loading live suppliers...");
    renderPage();
    await loadSuppliers();
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("suppliers");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize the suppliers directory.", true);
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

standardModeBtn?.addEventListener("click", () => {
  setMode("standard");
});

eeModeBtn?.addEventListener("click", () => {
  setMode("ee");
});

refreshSuppliersBtn?.addEventListener("click", () => {
  void loadSuppliers();
});

supplierSearchInput?.addEventListener("input", () => {
  renderPage();
});

supplierTypeSelect?.addEventListener("change", () => {
  activeCategoryFilter = normalizeValue(supplierTypeSelect.value);
  renderPage();
});

supplierCountySelect?.addEventListener("change", () => {
  renderPage();
});

supplierTownSelect?.addEventListener("change", () => {
  renderPage();
});

supplierRatingSelect?.addEventListener("change", () => {
  renderPage();
});

supplierTagSelect?.addEventListener("change", () => {
  renderPage();
});

supplierWebsiteOnly?.addEventListener("change", () => {
  renderPage();
});

clearFiltersBtn?.addEventListener("click", () => {
  resetFilters();
});

quickEntryForm?.addEventListener("input", () => {
  renderQuickEntryPreview();
});

saveEntryBtn?.addEventListener("click", () => {
  void handleSaveSupplier();
});

copyEntrySummaryBtn?.addEventListener("click", () => {
  void copyText(buildQuickEntryPreviewText(), "Copied quick entry summary.");
});

copyEntryCsvBtn?.addEventListener("click", () => {
  void copyText(buildQuickEntryCsv(), "Copied quick entry CSV row.");
});

resetEntryBtn?.addEventListener("click", () => {
  resetQuickEntry();
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
