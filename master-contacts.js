import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260603";
import { createSharePointApi } from "./sharepoint-list-utils.js";

const MARKETING_SITE_URL =
  FRONTEND_CONFIG.sharePoint?.marketingSiteUrl || "https://planwithcare.sharepoint.com/sites/Marketing";
const MASTER_CONTACTS_LIST_PATH =
  FRONTEND_CONFIG.sharePoint?.masterContactsListPath || "/sites/Marketing/Lists/Master Contacts";
const SHAREPOINT_LIST_URL =
  "https://planwithcare.sharepoint.com/sites/Marketing/Lists/Master%20Contacts/AllItems.aspx?viewid=d24bd223%2D977a%2D4584%2Da766%2Df5852044b758&env=WebViewList";
const SUPPORTED_FIELD_TYPES = new Set(["Text", "Note", "Choice", "MultiChoice", "Boolean", "DateTime", "Number", "Currency", "URL"]);
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

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const openListLink = document.getElementById("openListLink");
const contactForm = document.getElementById("contactForm");
const formMeta = document.getElementById("formMeta");
const unsupportedFieldsWrap = document.getElementById("unsupportedFieldsWrap");
const unsupportedFieldsText = document.getElementById("unsupportedFieldsText");
const saveBtn = document.getElementById("saveBtn");
const resetDraftBtn = document.getElementById("resetDraftBtn");
const saveStatus = document.getElementById("saveStatus");

const authController = createAuthController({
  tenantId: FRONTEND_CONFIG.tenantId,
  clientId: FRONTEND_CONFIG.spaClientId,
});
const directoryApi = createDirectoryApi(authController);

let currentRole = "";
let listInfo = null;
let supportedFields = [];
let unsupportedFields = [];
let spApi = null;
let saving = false;
let currentUserEmail = "";

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setSaveStatus(message, isError = false) {
  saveStatus.textContent = message;
  saveStatus.classList.toggle("error", isError);
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
  return normalizeKey(field?.Title) === "added by" || normalizeKey(field?.InternalName) === "addedby";
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
    ["title", 0],
    ["first name", 1],
    ["last name", 2],
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

function toFieldDefinition(field) {
  return {
    internalName: normalizeText(field?.InternalName),
    title: normalizeText(field?.Title),
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
  const parts = [];
  if (isAddedByField(field)) {
    parts.push("Set automatically");
  }
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
    if (control.tagName === "SELECT" && currentUserEmail && !field.choices.includes(currentUserEmail)) {
      const userOption = document.createElement("option");
      userOption.value = currentUserEmail;
      userOption.textContent = currentUserEmail;
      control.appendChild(userOption);
    }
    control.value = currentUserEmail;
    control.readOnly = true;
    if (control.tagName === "SELECT") {
      control.disabled = true;
    }
    control.setAttribute("aria-readonly", "true");
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
    label.textContent = field.title;

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

  unsupportedFieldsWrap.hidden = unsupportedFields.length === 0;
  unsupportedFieldsText.textContent = unsupportedFields.length
    ? unsupportedFields.map((field) => field.Title).join(", ")
    : "";
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
      throw new Error(`${field.title} must be a valid number.`);
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
    const rawValue = control?.value;
    if (field.required && !normalizeText(rawValue)) {
      throw new Error(`${field.title} is required.`);
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
}

async function loadListInfo() {
  const api = getSpApi();
  listInfo = await api.resolveListByPath(MASTER_CONTACTS_LIST_PATH);
  const fields = await api.getListFields(listInfo.Id);
  supportedFields = sortFields(fields.filter(isSupportedField)).map(toFieldDefinition);
  unsupportedFields = sortFields(fields.filter((field) => !isSupportedField(field) && !isSkippedField(field)));
  renderContactForm();
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
    const email = normalizeText(profile?.email);
    currentUserEmail = email || normalizeText(account?.username);
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

signOutBtn?.addEventListener("click", () => {
  authController.signOut().catch((error) => {
    console.error(error);
  });
});

void init();
