import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import { createDirectoryApi } from "./directory-api.js";
import { canAccessPage, renderTopNavigation } from "./navigation.js?v=20260411";
import { createSharePointApi } from "./sharepoint-list-utils.js";

const WELLBEING_SITE_URL =
  FRONTEND_CONFIG.sharePoint?.wellbeingSiteUrl || "https://planwithcare.sharepoint.com/sites/Wellbeing";
const SUPPLIERS_LIST_PATH =
  FRONTEND_CONFIG.sharePoint?.suppliersListPath || "/sites/Wellbeing/Lists/Suppliers Database";
const SHAREPOINT_LIST_URL = `${WELLBEING_SITE_URL}/Lists/Suppliers%20Database/AllItems.aspx`;
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

const signOutBtn = document.getElementById("signOutBtn");
const statusMessage = document.getElementById("statusMessage");
const openListLink = document.getElementById("openListLink");
const sourceTextInput = document.getElementById("sourceTextInput");
const extractBtn = document.getElementById("extractBtn");
const clearSourceBtn = document.getElementById("clearSourceBtn");
const extractStatus = document.getElementById("extractStatus");
const reviewForm = document.getElementById("reviewForm");
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
let extracting = false;
let saving = false;

function setStatus(message, isError = false) {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
}

function setExtractStatus(message, isError = false) {
  extractStatus.textContent = message;
  extractStatus.classList.toggle("error", isError);
}

function setSaveStatus(message, isError = false) {
  saveStatus.textContent = message;
  saveStatus.classList.toggle("error", isError);
}

function redirectToUnauthorized(pageKey) {
  const page = encodeURIComponent(String(pageKey || "wellbeingintake").trim().toLowerCase());
  window.location.href = `./unauthorized.html?page=${page}`;
}

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeKey(value) {
  return normalizeText(value).toLowerCase();
}

function getSiteConfig() {
  const siteUrl = new URL(WELLBEING_SITE_URL);
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
  return !internalName || SKIPPED_INTERNAL_NAMES.has(internalName);
}

function isSupportedField(field) {
  if (!field || isSkippedField(field) || field.Hidden || field.ReadOnlyField) {
    return false;
  }
  return SUPPORTED_FIELD_TYPES.has(normalizeText(field.TypeAsString));
}

function sortFields(fields) {
  return [...fields].sort((left, right) => {
    const leftTitle = normalizeText(left?.Title);
    const rightTitle = normalizeText(right?.Title);
    if (leftTitle === "Title") {
      return -1;
    }
    if (rightTitle === "Title") {
      return 1;
    }
    return leftTitle.localeCompare(rightTitle, undefined, { sensitivity: "base" });
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

function buildFieldHelpText(field) {
  const parts = [];
  if (field.required) {
    parts.push("Required");
  }
  if (field.type === "Choice" && field.choices.length && normalizeKey(field.title) !== "supplier type") {
    parts.push(`Choices: ${field.choices.join(", ")}`);
  }
  if (normalizeKey(field.title) === "supplier type" && field.choices.length) {
    parts.push(`Suggestions: ${field.choices.join(", ")}`);
  }
  if (field.description) {
    parts.push(field.description);
  }
  return parts.join(" • ");
}

function buildInputId(field) {
  return `wellbeing-field-${field.internalName}`;
}

function createFieldControl(field) {
  const inputId = buildInputId(field);
  let control;
  const isSupplierTypeField = normalizeKey(field.title) === "supplier type";

  if (field.type === "Note") {
    control = document.createElement("textarea");
    control.rows = 4;
  } else if (field.type === "Choice" && !isSupplierTypeField) {
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
    if (field.type === "URL") {
      control.placeholder = "https://";
    }
    if (isSupplierTypeField && field.choices.length) {
      const listId = `${inputId}-suggestions`;
      const datalist = document.createElement("datalist");
      datalist.id = listId;
      for (const choice of field.choices) {
        const option = document.createElement("option");
        option.value = choice;
        datalist.appendChild(option);
      }
      reviewForm.appendChild(datalist);
      control.setAttribute("list", listId);
      control.placeholder = "Type a supplier category";
    }
  }

  control.id = inputId;
  control.dataset.internalName = field.internalName;
  control.dataset.fieldType = field.type;
  control.required = field.required;
  control.classList.add("wellbeing-field-control");
  return control;
}

function renderReviewForm() {
  reviewForm.innerHTML = "";

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

    reviewForm.appendChild(label);
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
  return reviewForm.querySelector(`[data-internal-name="${internalName}"]`);
}

function setControlValue(field, value) {
  const control = getFieldControl(field.internalName);
  if (!control) {
    return;
  }
  control.value = value == null ? "" : String(value);
}

function resetDraft() {
  for (const field of supportedFields) {
    setControlValue(field, "");
  }
  setSaveStatus("Draft cleared.");
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

function collectDraftPayload() {
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

function setBusyState() {
  extractBtn.disabled = extracting || saving || !supportedFields.length;
  saveBtn.disabled = saving || extracting || !supportedFields.length;
  resetDraftBtn.disabled = saving || extracting || !supportedFields.length;
  clearSourceBtn.disabled = extracting || saving;
  extractBtn.textContent = extracting ? "Extracting..." : "Extract with AI";
  saveBtn.textContent = saving ? "Saving..." : "Save to SharePoint";
}

async function loadListInfo() {
  const api = getSpApi();
  listInfo = await api.resolveListByPath(SUPPLIERS_LIST_PATH);
  const fields = await api.getListFields(listInfo.Id);
  supportedFields = sortFields(fields.filter(isSupportedField)).map(toFieldDefinition);
  unsupportedFields = sortFields(fields.filter((field) => !isSupportedField(field) && !isSkippedField(field)));
  renderReviewForm();
}

async function handleExtract() {
  const sourceText = normalizeText(sourceTextInput.value);
  if (!sourceText) {
    setExtractStatus("Paste some text first.", true);
    sourceTextInput.focus();
    return;
  }

  extracting = true;
  setExtractStatus("Extracting a draft from the pasted text...");
  setBusyState();

  try {
    const payload = await directoryApi.parseWellbeingIntake({
      sourceText,
      fields: supportedFields,
    });
    const values = payload?.values && typeof payload.values === "object" ? payload.values : {};
    for (const field of supportedFields) {
      setControlValue(field, values[field.internalName]);
    }
    setExtractStatus("Draft extracted. Review every field before saving.");
    setSaveStatus("The form has been filled with the AI draft. Edit anything that needs changing.");
  } catch (error) {
    console.error(error);
    setExtractStatus(error?.message || "Could not extract a draft from the pasted text.", true);
  } finally {
    extracting = false;
    setBusyState();
  }
}

async function handleSave() {
  saving = true;
  setSaveStatus("Saving item to SharePoint...");
  setBusyState();

  try {
    const payload = collectDraftPayload();
    const result = await getSpApi().createListItem(listInfo.Id, payload);
    const itemId = Number(result?.d?.Id || 0);
    setSaveStatus(itemId ? `Saved to SharePoint as item #${itemId}.` : "Saved to SharePoint.");
  } catch (error) {
    console.error(error);
    setSaveStatus(error?.message || "Could not save the SharePoint entry.", true);
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
    if (!canAccessPage(currentRole, "wellbeingintake")) {
      redirectToUnauthorized("wellbeingintake");
      return;
    }

    renderTopNavigation({ role: currentRole });
    const email = normalizeText(profile?.email);
    setStatus(email ? `Signed in as ${email}. Loading live SharePoint fields...` : "Loading live SharePoint fields...");
    await loadListInfo();
    setStatus(`Ready to create a new item in ${normalizeText(listInfo?.Title) || "the SharePoint list"}.`);
    setExtractStatus("Paste text and extract a draft when you are ready.");
    setSaveStatus("Nothing has been saved yet.");
  } catch (error) {
    if (error?.status === 403) {
      redirectToUnauthorized("wellbeingintake");
      return;
    }
    console.error(error);
    setStatus(error?.message || "Could not initialize the Wellbeing intake page.", true);
    setExtractStatus("The form is not ready yet.", true);
    setSaveStatus("Saving is unavailable until the SharePoint form loads.", true);
  } finally {
    setBusyState();
    document.body.classList.remove("auth-pending");
  }
}

extractBtn?.addEventListener("click", () => {
  void handleExtract();
});

saveBtn?.addEventListener("click", () => {
  void handleSave();
});

resetDraftBtn?.addEventListener("click", () => {
  resetDraft();
});

clearSourceBtn?.addEventListener("click", () => {
  sourceTextInput.value = "";
  setExtractStatus("Source text cleared.");
});

signOutBtn?.addEventListener("click", () => {
  authController.signOut().catch((error) => {
    console.error(error);
  });
});

void init();
