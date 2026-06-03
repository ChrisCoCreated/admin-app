import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260603";
import { buildFieldMaps, createSharePointApi, fieldByTitle } from "./sharepoint-list-utils.js";

const MARKETING_SITE_URL =
  FRONTEND_CONFIG.sharePoint?.marketingSiteUrl || "https://planwithcare.sharepoint.com/sites/Marketing";
const MASTER_CONTACTS_LIST_PATH =
  FRONTEND_CONFIG.sharePoint?.masterContactsListPath || "/sites/Marketing/Lists/Master Contacts";
const SHAREPOINT_LIST_URL =
  "https://planwithcare.sharepoint.com/sites/Marketing/Lists/Master%20Contacts/AllItems.aspx?env=WebViewList";
const SUPPORTED_FIELD_TYPES = new Set(["Text", "Note", "Choice", "MultiChoice", "Boolean", "DateTime", "Number", "Currency", "URL"]);
const LINKEDIN_BATCH_SIZE = 10;
const SKIPPED_INTERNAL_NAMES = new Set([
  "id",
  "contenttype",
  "contenttypeid",
  "created",
  "modified",
  "author",
  "editor",
  "attachments",
  "guid",
  "complianceassetid",
  "_moderationstatus",
  "_moderationcomments",
]);
const SKIPPED_FIELD_TITLES = new Set([
  "enquiry date time",
  "invite to east kent events",
  "response 15/01/26",
]);
const ADDED_BY_FIELD_HINTS = [
  "addedby",
  "userwhofirstsentanemailtothisperson",
  "whofirstsentanemail",
  "firstsentanemail",
  "firstsentemail",
];

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const openListLink = document.getElementById("openListLink");
const contactForm = document.getElementById("contactForm");
const formMeta = document.getElementById("formMeta");
const saveBtn = document.getElementById("saveBtn");
const resetDraftBtn = document.getElementById("resetDraftBtn");
const saveStatus = document.getElementById("saveStatus");
const linkedinMeta = document.getElementById("linkedinMeta");
const linkedinStatus = document.getElementById("linkedinStatus");
const openLinkedInBatchBtn = document.getElementById("openLinkedInBatchBtn");
const nextLinkedInBatchBtn = document.getElementById("nextLinkedInBatchBtn");
const resetLinkedInBatchBtn = document.getElementById("resetLinkedInBatchBtn");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let currentRole = "";
let listInfo = null;
let supportedFields = [];
let spApi = null;
let saving = false;
let currentUserEmail = "";
let loadingLinkedInLinks = false;
let linkedinLinks = [];
let linkedinBatchIndex = 0;

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setSaveStatus(message, isError = false) {
  saveStatus.textContent = message;
  saveStatus.classList.toggle("error", isError);
}

function setLinkedInStatus(message, isError = false) {
  linkedinStatus.textContent = message;
  linkedinStatus.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "mastercontacts").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeKey(value) {
  return normalizeText(value).toLowerCase();
}

function normalizeIdentifier(value) {
  return normalizeKey(value)
    .replace(/_x[0-9a-f]{4}_/gi, "")
    .replace(/[^a-z0-9]/g, "");
}

function getAccountUsername(account) {
  return normalizeText(
    account?.username ||
      account?.idTokenClaims?.preferred_username ||
      account?.idTokenClaims?.upn ||
      account?.idTokenClaims?.email,
  );
}

function getSiteConfig() {
  const siteUrl = new URL(MARKETING_SITE_URL);
  return {
    siteUrl: siteUrl.origin + siteUrl.pathname.replace(/\/$/, ""),
    siteHost: siteUrl.origin,
  };
}

function getSpApi() {
  if (!spApi) {
    const config = getSiteConfig();
    spApi = createSharePointApi({
      siteUrl: config.siteUrl,
      getToken: () => authController.acquireSharePointToken(config.siteHost),
    });
  }
  return spApi;
}

function isSkippedField(field) {
  const internalName = normalizeKey(field?.InternalName);
  const title = normalizeKey(field?.Title);
  return !internalName || SKIPPED_INTERNAL_NAMES.has(internalName) || SKIPPED_FIELD_TITLES.has(title);
}

function isAddedByField(field) {
  const candidates = [
    normalizeIdentifier(field?.Title || field?.title),
    normalizeIdentifier(field?.InternalName || field?.internalName),
    normalizeIdentifier(field?.Description || field?.description),
  ].filter(Boolean);
  return candidates.some((candidate) => ADDED_BY_FIELD_HINTS.some((hint) => candidate.includes(hint)));
}

function isSupportedField(field) {
  if (!field || isSkippedField(field) || field.Hidden || field.ReadOnlyField) {
    return false;
  }
  return SUPPORTED_FIELD_TYPES.has(normalizeText(field.TypeAsString));
}

function getFieldPriority(field) {
  const key = normalizeKey(field?.Title);
  if (isAddedByField(field)) {
    return 1000;
  }

  const priorities = new Map([
    ["first name", 0],
    ["title", 1],
    ["surname", 1],
    ["last name", 1],
    ["name", 3],
    ["organisation", 4],
    ["organization", 4],
    ["company", 5],
    ["job title", 6],
    ["email", 7],
    ["email address", 7],
    ["phone", 8],
    ["phone number", 8],
    ["mobile", 9],
    ["website", 10],
    ["notes", 99],
  ]);
  return priorities.get(key) ?? 50;
}

function sortFields(fields) {
  return [...fields].sort((left, right) => {
    const priorityDelta = getFieldPriority(left) - getFieldPriority(right);
    if (priorityDelta !== 0) {
      return priorityDelta;
    }
    return normalizeText(left?.Title).localeCompare(normalizeText(right?.Title), undefined, { sensitivity: "base" });
  });
}

function getContactLabel(item) {
  const parts = [
    normalizeText(item?.FirstName || item?.First_x0020_Name),
    normalizeText(item?.Title || item?.Surname || item?.LastName || item?.Last_x0020_Name),
    normalizeText(item?.Organisation || item?.Organization || item?.Company),
  ].filter(Boolean);
  return parts.length ? parts.join(" ") : `SharePoint item #${item?.Id || ""}`.trim();
}

function normalizeUrlValue(value) {
  if (!value) {
    return "";
  }

  if (typeof value === "string") {
    const trimmed = normalizeText(value);
    if (!trimmed) {
      return "";
    }
    const [maybeUrl] = trimmed.split(/,\s*/);
    return maybeUrl || trimmed;
  }

  return normalizeText(value.Url || value.url || value.Description || value.description);
}

function normalizeExternalUrl(value) {
  const url = normalizeUrlValue(value);
  if (!url) {
    return "";
  }
  if (/^https?:\/\//i.test(url)) {
    return url;
  }
  if (/^www\./i.test(url) || /^[a-z0-9.-]+\.[a-z]{2,}(\/|$)/i.test(url)) {
    return `https://${url}`;
  }
  return "";
}

function getLinkedInLinkSummary() {
  const total = linkedinLinks.length;
  if (!total) {
    return "No LinkedIn links were found in the live list.";
  }
  const opened = Math.min(linkedinBatchIndex, total);
  const remaining = Math.max(total - opened, 0);
  return `${total} LinkedIn link${total === 1 ? "" : "s"} loaded. ${opened} opened, ${remaining} remaining.`;
}

function setLinkedInBusyState() {
  const hasLinks = linkedinLinks.length > 0;
  const canOpenNext = hasLinks && linkedinBatchIndex < linkedinLinks.length && !loadingLinkedInLinks;
  const canReset = hasLinks && linkedinBatchIndex > 0 && !loadingLinkedInLinks;
  openLinkedInBatchBtn.disabled = !canOpenNext;
  nextLinkedInBatchBtn.disabled = !canOpenNext;
  resetLinkedInBatchBtn.disabled = !canReset;
  openLinkedInBatchBtn.textContent = linkedinBatchIndex > 0 ? "Open next 10" : "Open first 10";
  linkedinMeta.textContent = loadingLinkedInLinks ? "Loading LinkedIn links from the live list..." : getLinkedInLinkSummary();
}

function toFieldDefinition(field) {
  const title = normalizeText(field?.Title);
  const displayTitle = isAddedByField(field) ? "Added by" : normalizeKey(title) === "title" ? "Surname" : title;
  return {
    internalName: normalizeText(field?.InternalName),
    title,
    displayTitle,
    type: normalizeText(field?.TypeAsString),
    required: field?.Required === true,
    description: normalizeText(field?.Description),
    choices: Array.isArray(field?.Choices?.results)
      ? field.Choices.results.map((entry) => normalizeText(entry)).filter(Boolean)
      : [],
  };
}

function buildInputId(field) {
  return `master-contact-field-${field.internalName}`;
}

function buildFieldHelpText(field) {
  if (isAddedByField(field)) {
    return "";
  }

  const parts = [];
  if (field.required) {
    parts.push("Required");
  }
  if ((field.type === "Choice" || field.type === "MultiChoice") && field.choices.length) {
    parts.push(`Choices: ${field.choices.join(", ")}`);
  }
  if (field.description) {
    parts.push(field.description);
  }
  return parts.join(" • ");
}

function createFieldControl(field) {
  const inputId = buildInputId(field);
  let control;

  if (field.type === "Note" || field.type === "MultiChoice") {
    control = document.createElement("textarea");
    control.rows = field.type === "MultiChoice" ? 3 : 4;
    if (field.type === "MultiChoice") {
      control.placeholder = "Separate multiple choices with commas";
    }
  } else if (field.type === "Choice") {
    control = document.createElement("select");
    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "Select an option";
    control.appendChild(emptyOption);
    for (const choice of field.choices) {
      const option = document.createElement("option");
      option.value = choice;
      option.textContent = choice;
      control.appendChild(option);
    }
  } else if (field.type === "Boolean") {
    control = document.createElement("select");
    const emptyOption = document.createElement("option");
    emptyOption.value = "";
    emptyOption.textContent = "Select yes or no";
    const yesOption = document.createElement("option");
    yesOption.value = "true";
    yesOption.textContent = "Yes";
    const noOption = document.createElement("option");
    noOption.value = "false";
    noOption.textContent = "No";
    control.append(emptyOption, yesOption, noOption);
  } else {
    control = document.createElement("input");
    control.type = field.type === "DateTime" ? "date" : field.type === "Number" || field.type === "Currency" ? "number" : "text";
    if (field.type === "Number" || field.type === "Currency") {
      control.step = "any";
    }
    if (field.type === "URL" || normalizeKey(field.title).includes("website")) {
      control.placeholder = "https://";
    }
    if (normalizeKey(field.title).includes("email")) {
      control.type = "email";
    }
    if (normalizeKey(field.title).includes("phone") || normalizeKey(field.title).includes("mobile")) {
      control.type = "tel";
    }
  }

  control.id = inputId;
  control.dataset.internalName = field.internalName;
  control.dataset.fieldType = field.type;
  control.required = field.required;
  if (isAddedByField(field)) {
    console.info("[Add Contact] Added by field detected", {
      title: field.title,
      internalName: field.internalName,
      expectedUsername: currentUserEmail,
    });
    if (control.tagName === "SELECT" && currentUserEmail && !field.choices.includes(currentUserEmail)) {
      const userOption = document.createElement("option");
      userOption.value = currentUserEmail;
      userOption.textContent = currentUserEmail;
      control.appendChild(userOption);
    }
    control.value = currentUserEmail;
    control.readOnly = true;
    control.classList.add("master-contacts-readonly-field");
    if (control.tagName === "SELECT") {
      control.disabled = true;
    }
    control.setAttribute("aria-readonly", "true");
    console.info("[Add Contact] Added by control populated", {
      internalName: field.internalName,
      value: control.value,
    });
  }
  control.classList.add("wellbeing-field-control");
  return control;
}

function renderContactForm() {
  contactForm.innerHTML = "";

  for (const field of supportedFields) {
    const label = document.createElement("label");
    label.className = "field";
    label.htmlFor = buildInputId(field);
    label.textContent = field.displayTitle || field.title;

    const control = createFieldControl(field);
    label.appendChild(control);

    const helpText = buildFieldHelpText(field);
    if (helpText) {
      const help = document.createElement("span");
      help.className = "muted wellbeing-field-help";
      help.textContent = helpText;
      label.appendChild(help);
    }

    contactForm.appendChild(label);
  }

  formMeta.textContent = supportedFields.length
    ? `${supportedFields.length} editable SharePoint fields loaded from the live list.`
    : "No editable SharePoint fields were found.";
}

function getFieldControl(internalName) {
  return contactForm.querySelector(`[data-internal-name="${internalName}"]`);
}

function setAddedByValue() {
  for (const field of supportedFields) {
    if (!isAddedByField(field)) {
      continue;
    }
    const control = getFieldControl(field.internalName);
    if (control) {
      control.value = currentUserEmail;
      console.info("[Add Contact] Added by value reset", {
        internalName: field.internalName,
        expectedUsername: currentUserEmail,
        controlValue: control.value,
      });
    }
  }
}

function coerceFieldValue(field, rawValue) {
  const value = normalizeText(rawValue);
  if (!value) {
    return null;
  }

  if (field.type === "Boolean") {
    return value.toLowerCase() === "true";
  }

  if (field.type === "Number" || field.type === "Currency") {
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) {
      throw new Error(`${field.displayTitle || field.title} must be a valid number.`);
    }
    return numeric;
  }

  if (field.type === "MultiChoice") {
    const values = value
      .split(/[;,]/)
      .map((entry) => normalizeText(entry))
      .filter(Boolean);
    return values.length ? { results: values } : null;
  }

  if (field.type === "URL") {
    return {
      __metadata: { type: "SP.FieldUrlValue" },
      Url: value,
      Description: value,
    };
  }

  return value;
}

function collectPayload() {
  setAddedByValue();
  const payload = {
    __metadata: { type: listInfo?.ListItemEntityTypeFullName || "" },
  };

  for (const field of supportedFields) {
    const control = getFieldControl(field.internalName);
    const rawValue = isAddedByField(field) ? currentUserEmail : control?.value;
    if (isAddedByField(field)) {
      console.info("[Add Contact] Added by payload value", {
        internalName: field.internalName,
        expectedUsername: currentUserEmail,
        controlValue: control?.value || "",
        payloadValue: rawValue,
      });
    }
    if (field.required && !normalizeText(rawValue)) {
      throw new Error(`${field.displayTitle || field.title} is required.`);
    }
    const coerced = coerceFieldValue(field, rawValue);
    if (coerced != null) {
      payload[field.internalName] = coerced;
    }
  }

  return payload;
}

function resetForm() {
  contactForm.reset();
  setAddedByValue();
  setSaveStatus("Form cleared.");
}

function setBusyState() {
  saveBtn.disabled = saving || !supportedFields.length;
  resetDraftBtn.disabled = saving || !supportedFields.length;
  saveBtn.textContent = saving ? "Saving..." : "Save to SharePoint";
  setLinkedInBusyState();
}

async function loadListInfo() {
  const api = getSpApi();
  listInfo = await api.resolveListByPath(MASTER_CONTACTS_LIST_PATH);
  const fields = await api.getListFields(listInfo.Id);
  console.info(
    "[Add Contact] SharePoint fields loaded",
    fields.map((field) => ({
      title: field?.Title || "",
      internalName: field?.InternalName || "",
      type: field?.TypeAsString || "",
      description: field?.Description || "",
      isAddedBy: isAddedByField(field),
      isSupported: isSupportedField(field),
      hidden: field?.Hidden === true,
      readOnly: field?.ReadOnlyField === true,
    })),
  );
  supportedFields = sortFields(fields.filter(isSupportedField)).map(toFieldDefinition);
  if (!supportedFields.some(isAddedByField)) {
    console.warn("[Add Contact] Added by field was not found in supported fields.", {
      expectedUsername: currentUserEmail,
    });
  }
  renderContactForm();
  setAddedByValue();
  await loadLinkedInLinks(fields);
}

async function fetchAllListItems(path) {
  const api = getSpApi();
  const items = [];
  let nextPath = path;

  while (nextPath) {
    const data = await api.request(nextPath);
    if (Array.isArray(data?.d?.results)) {
      items.push(...data.d.results);
    }
    nextPath = data?.d?.__next || "";
  }

  return items;
}

async function loadLinkedInLinks(fields) {
  loadingLinkedInLinks = true;
  linkedinLinks = [];
  linkedinBatchIndex = 0;
  setLinkedInStatus("Loading LinkedIn links from SharePoint...");
  setLinkedInBusyState();

  try {
    const fieldMap = buildFieldMaps(fields);
    const linkedInField = fieldByTitle(fieldMap, ["LinkedIn", "Linked In", "LinkedIn URL", "LinkedIn Profile"]);
    if (!linkedInField?.InternalName) {
      throw new Error("Could not find a LinkedIn column in Master Contacts.");
    }

    const selectFields = Array.from(new Set(["Id", "Title", linkedInField.InternalName]));
    const query = selectFields.join(",");
    const items = await fetchAllListItems(
      `/_api/web/lists(guid'${listInfo.Id}')/items?$top=5000&$orderby=Id asc&$select=${encodeURIComponent(query)}`
    );
    const seenUrls = new Set();

    linkedinLinks = items
      .map((item) => {
        const url = normalizeExternalUrl(item?.[linkedInField.InternalName]);
        return {
          id: Number(item?.Id || 0),
          label: getContactLabel(item),
          url,
        };
      })
      .filter((entry) => {
        if (!entry.url || seenUrls.has(entry.url)) {
          return false;
        }
        seenUrls.add(entry.url);
        return true;
      });

    setLinkedInStatus(linkedinLinks.length ? "Ready to open LinkedIn links in batches of 10." : "No LinkedIn links were found.");
  } catch (error) {
    console.error(error);
    setLinkedInStatus(error?.message || "Could not load LinkedIn links.", true);
  } finally {
    loadingLinkedInLinks = false;
    setLinkedInBusyState();
  }
}

function openLinkedInBatch() {
  if (!linkedinLinks.length || linkedinBatchIndex >= linkedinLinks.length) {
    setLinkedInStatus("All LinkedIn links have been opened.");
    setLinkedInBusyState();
    return;
  }

  const startIndex = linkedinBatchIndex;
  const batch = linkedinLinks.slice(startIndex, startIndex + LINKEDIN_BATCH_SIZE);
  let openedCount = 0;

  for (const entry of batch) {
    const openedWindow = window.open(entry.url, "_blank");
    if (openedWindow) {
      openedWindow.opener = null;
      openedCount += 1;
    }
  }

  linkedinBatchIndex = startIndex + batch.length;
  const firstNumber = startIndex + 1;
  const lastNumber = startIndex + batch.length;
  const blockedCount = batch.length - openedCount;
  const blockedNote = blockedCount ? ` ${blockedCount} tab${blockedCount === 1 ? "" : "s"} may have been blocked by the browser.` : "";
  setLinkedInStatus(`Opened LinkedIn links ${firstNumber}-${lastNumber} of ${linkedinLinks.length}.${blockedNote}`, Boolean(blockedCount));
  setLinkedInBusyState();
}

function resetLinkedInBatch() {
  linkedinBatchIndex = 0;
  setLinkedInStatus("Reset to the first LinkedIn batch.");
  setLinkedInBusyState();
}

async function handleSave() {
  saving = true;
  setSaveStatus("Saving contact to SharePoint...");
  setBusyState();

  try {
    const payload = collectPayload();
    const result = await getSpApi().createListItem(listInfo.Id, payload);
    const itemId = Number(result?.d?.Id || 0);
    contactForm.reset();
    setAddedByValue();
    setSaveStatus(itemId ? `Saved to SharePoint as item #${itemId}.` : "Saved to SharePoint.");
  } catch (error) {
    console.error(error);
    setSaveStatus(error?.message || "Could not save the contact.", true);
  } finally {
    saving = false;
    setBusyState();
  }
}

async function init() {
  try {
    openListLink.href = SHAREPOINT_LIST_URL;
    const account = await authController.restoreSession();
    if (!account) {
      window.location.href = "./index.html";
      return;
    }

    const profile = await directoryApi.getCurrentUser();
    currentRole = normalizeKey(profile?.role);
    if (!canAccessPage(currentRole, "mastercontacts")) {
      redirectToUnauthorized("mastercontacts");
      return;
    }

    renderTopNavigation({ role: currentRole });
    currentUserEmail = getAccountUsername(account);
    const email = normalizeText(profile?.email);
    if (!currentUserEmail) {
      currentUserEmail = email;
    }
    console.info("[Add Contact] Current user resolved", {
      accountUsername: getAccountUsername(account),
      profileEmail: email,
      expectedAddedBy: currentUserEmail,
    });
    setStatus(email ? `Signed in as ${email}. Loading live SharePoint fields...` : "Loading live SharePoint fields...");
    await loadListInfo();
    setStatus(`Ready to add a contact to ${normalizeText(listInfo?.Title) || "Master Contacts"}.`);
    setSaveStatus("Nothing has been saved yet.");
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("mastercontacts");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize the Add Contact form.", true);
    setSaveStatus("Saving is unavailable until the SharePoint form loads.", true);
  } finally {
    setBusyState();
    document.body.classList.remove("auth-pending");
  }
}

saveBtn?.addEventListener("click", () => {
  void handleSave();
});

resetDraftBtn?.addEventListener("click", () => {
  resetForm();
});

openLinkedInBatchBtn?.addEventListener("click", () => {
  openLinkedInBatch();
});

nextLinkedInBatchBtn?.addEventListener("click", () => {
  openLinkedInBatch();
});

resetLinkedInBatchBtn?.addEventListener("click", () => {
  resetLinkedInBatch();
});

signOutBtn?.addEventListener("click", () => {
  authController.signOut().catch((error) => {
    console.error(error);
  });
});

void init();
