import { createAuthController } from "./auth-common.js";
import { FRONTEND_CONFIG } from "./frontend-config.js";
import {
  buildFieldMaps,
  createSharePointApi,
  fieldByInternalName as sharedFieldByInternalName,
  fieldByTitle as sharedFieldByTitle,
  fieldByTitleWritable as sharedFieldByTitleWritable,
} from "./sharepoint-list-utils.js";

const $ = (id) => document.getElementById(id);

const FIXED_TENANT_ID = FRONTEND_CONFIG.tenantId;
const FIXED_CLIENT_ID = FRONTEND_CONFIG.sharePointSpaClientId || FRONTEND_CONFIG.spaClientId;
const FIXED_SITE_URL =
  FRONTEND_CONFIG.sharePoint?.thriveCallsSiteUrl || "https://planwithcare.sharepoint.com/sites/ThriveCalls";
const ENQUIRIES_LIST_TITLE = FRONTEND_CONFIG.sharePoint?.enquiriesListTitle || "Enquiries Log";
const ENQUIRIES_LIST_PATH = `${new URL(FIXED_SITE_URL).pathname}/Lists/${ENQUIRIES_LIST_TITLE}`;

const signInBtn = $("signInBtn");
const signOutBtn = $("signOutBtn");
const authState = $("authState");
const createEnquiryBtn = $("createEnquiryBtn");
const saveUpdateBtn = $("saveUpdateBtn");
const listWrap = $("listWrap");
const searchInput = $("searchInput");
const toggleFiltersBtn = $("toggleFiltersBtn");
const filterPanel = $("filterPanel");
const ownerFilter = $("ownerFilter");
const scopeFilter = $("scopeFilter");
const sortFilter = $("sortFilter");
const showBusinessInfoBtn = $("showBusinessInfoBtn");
const updatePanel = $("updatePanel");
const businessFields = $("businessFields");
const selectedTitle = $("selectedTitle");
const selectedMeta = $("selectedMeta");
const selectedEnquiryLink = $("selectedEnquiryLink");
const statusMessage = $("statusMessage");
const statusDetail = $("statusDetail");
const qClientTbcBtn = $("qClientTbcBtn");
const lostFields = $("lostFields");
const quickEntryToggleBtn = $("quickEntryToggleBtn");
const quickEntryContent = $("quickEntryContent");
const qReferralFromWrap = $("qReferralFromWrap");
const qBusDevNotesWrap = $("qBusDevNotesWrap");
const qLocationCustomWrap = $("qLocationCustomWrap");
const uLocationCustomWrap = $("uLocationCustomWrap");
const dataQualityCard = $("dataQualityCard");
const dataQualityCount = $("dataQualityCount");
const dataQualityList = $("dataQualityList");
const dataQualityToggleBtn = $("dataQualityToggleBtn");
const dataQualityContent = $("dataQualityContent");
const authCard = document.querySelector(".auth-card");

let account;
let spApi = null;
let listInfo = null;
let fieldMap = null;
let allEnquiries = [];
let enquiries = [];
let selectedItem = null;
let tbcAutoMode = false;
let currentSpUserId = null;

const LOCATION_OTHER_VALUE = "__other__";
const LOCATION_CHOICES = ["Ashford", "Canterbury", "Deal", "Folkestone", "Maidstone", "London", "West Kent"];
const LOST_STATUSES_REQUIRING_DETAILS = new Set(["lost - post assessment", "lost - pre assessment"]);
const SOURCE_OPTIONS = [
  "Family/Friend",
  "Professional",
  "Legal Representative",
  "Autumna",
  "Lottie",
  "Team Member",
  "carehome.co.uk",
  "Google",
  "Event",
  "BNI",
  "Saw Article",
  "Community Partner",
];
const LOSS_REASON_OPTIONS = [
  "Cost – Too Expensive",
  "Timing – Not Ready Yet",
  "Timing – Urgent Need, We Had No Availability",
  "Location – Outside Service Area",
  "Needs – Required Support Beyond Our Offer",
  "Changed Circumstances – Care No Longer Needed",
  "Changed Circumstances – Client Passed Away",
  "Communication – Couldn’t Reach Client / No Response",
  "Internal – Duplicate Enquiry or Admin Error",
  "Competition - Went with another provider",
];

const authController = createAuthController({
  tenantId: FIXED_TENANT_ID,
  clientId: FIXED_CLIENT_ID,
  onSignedIn: (signedInAccount) => {
    account = signedInAccount;
    setSignedInUi();
  },
  onSignedOut: () => {
    account = null;
    spApi?.clearCaches();
    spApi = null;
    currentSpUserId = null;
    allEnquiries = [];
    enquiries = [];
    selectedItem = null;
    setSignedOutUi();
    setStatus("Signed out.");
  },
});

function setSignedInUi() {
  if (authCard) {
    authCard.hidden = true;
  }
  authState.textContent = account?.username ? `Signed in as ${account.username}` : "Signed in";
  createEnquiryBtn.disabled = false;
  searchInput.disabled = false;
  toggleFiltersBtn.disabled = false;
  ownerFilter.disabled = false;
  scopeFilter.disabled = false;
  sortFilter.disabled = false;
  showBusinessInfoBtn.disabled = true;
  setBusinessInfoVisible(false);
  if (signOutBtn) {
    signOutBtn.hidden = false;
    signOutBtn.disabled = false;
  }
}

function setSignedOutUi() {
  if (authCard) {
    authCard.hidden = false;
  }
  authState.textContent = "Please sign in";
  createEnquiryBtn.disabled = true;
  searchInput.disabled = true;
  toggleFiltersBtn.disabled = true;
  ownerFilter.disabled = true;
  scopeFilter.disabled = true;
  sortFilter.disabled = true;
  showBusinessInfoBtn.disabled = true;
  ownerFilter.value = "mine";
  scopeFilter.value = "active";
  sortFilter.value = "status";
  saveUpdateBtn.disabled = true;
  updatePanel.hidden = true;
  businessFields.hidden = true;
  setBusinessInfoVisible(false);
  setFilterPanelVisible(false);
  listWrap.innerHTML = '<p class="muted">Sign in to load enquiries.</p>';
  selectedEnquiryLink.hidden = true;
  selectedEnquiryLink.removeAttribute("href");
  dataQualityCard.hidden = true;
  dataQualityCount.textContent = "";
  dataQualityList.innerHTML = "";
  if (signOutBtn) {
    signOutBtn.hidden = true;
  }
}

function setStatus(message, isError = false, detail = "") {
  statusMessage.textContent = message;
  statusMessage.classList.toggle("error", isError);
  statusDetail.textContent = detail;
}

function toIsoDate(dateValue) {
  if (!dateValue) {
    return null;
  }

  const date = new Date(`${dateValue}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date.toISOString();
}

function fromIsoDate(dateValue) {
  if (!dateValue) {
    return "";
  }

  const date = new Date(dateValue);
  if (Number.isNaN(date.getTime())) {
    return "";
  }

  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${date.getFullYear()}-${month}-${day}`;
}

function fieldByTitle(candidates) {
  return sharedFieldByTitle(fieldMap, candidates)?.InternalName || null;
}

function fieldByTitleWritable(candidates) {
  return sharedFieldByTitleWritable(fieldMap, candidates)?.InternalName || null;
}

function fieldByInternalName(candidates) {
  return sharedFieldByInternalName(fieldMap, candidates)?.InternalName || null;
}

function resolveFieldAliases() {
  return {
    fullName:
      fieldByTitleWritable(["Clients Full Name"]) ||
      fieldByTitleWritable(["Title"]) ||
      fieldByInternalName(["Title"]),
    callerName: fieldByTitle(["Person Making Enquiry"]),
    relationship: fieldByTitle(["Relationship to proposed Client"]),
    enquiryOwner: fieldByTitle(["Enquiry owned by"]),
    phone: fieldByTitle(["Phone Number"]),
    email: fieldByTitle(["Email of Caller"]),
    source: fieldByTitle(["Source of enquiry"]),
    status: fieldByTitle(["Status"]),
    currentStatus: fieldByTitle(["Current Status"]),
    updates: fieldByTitle(["Updates"]),
    followUp: fieldByTitle(["Follow up", "Followup"]),
    likelihood: fieldByTitle(["Likelihood"]),
    rateHourly: fieldByTitle(["rate_hourly"]),
    reasonForLoss: fieldByTitle(["Reason for loss"]),
    busDevNotes: fieldByTitle(["bus_dev_notes"]),
    referralFrom: fieldByTitle(["Referral From"]),
    location: fieldByTitle(["Location"]),
    callbackDate: fieldByTitle(["Callback arrange for", "Callback arrange for "]),
    assessmentDate: fieldByTitle(["Assessor's meeting arrange for", "Assessor's meeting arrange for "]),
    background: fieldByTitle(["Background", "How can we help you?"]),
  };
}

function isLostStatus(value) {
  return String(value || "").toLowerCase().includes("lost");
}

function isCompletedStatus(value) {
  const status = String(value || "").toLowerCase().trim();
  if (!status) {
    return false;
  }

  return status.includes("lost") || status === "won" || status === "didn't enquire" || status === "not qualified";
}

function syncLostFieldsVisibility() {
  lostFields.hidden = !isLostStatus($("uStatus").value || "");
}

function setQuickEntryVisible(visible) {
  quickEntryContent.hidden = !visible;
  quickEntryToggleBtn.setAttribute("aria-expanded", visible ? "true" : "false");
}

function setDataQualityVisible(visible) {
  dataQualityContent.hidden = !visible;
  dataQualityToggleBtn.setAttribute("aria-expanded", visible ? "true" : "false");
}

function syncReferralDetailsVisibility() {
  const hasSource = Boolean($("qSource").value);
  qReferralFromWrap.hidden = !hasSource;
  qBusDevNotesWrap.hidden = !hasSource;
}

function setInputError(element, hasError) {
  if (!element) {
    return;
  }
  element.classList.toggle("input-error", hasError);
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function isBlank(value) {
  return !String(value || "").trim();
}

function isStatusRequiringLossDetails(statusValue) {
  return LOST_STATUSES_REQUIRING_DETAILS.has(String(statusValue || "").trim().toLowerCase());
}

function buildSelectOptions(options, placeholderLabel) {
  const optionHtml = options.map((option) => `<option value="${escapeHtml(option)}">${escapeHtml(option)}</option>`).join("");
  return `<option value="">${escapeHtml(placeholderLabel)}</option>${optionHtml}`;
}

function getFieldValue(item, internalName) {
  if (!internalName) {
    return "";
  }
  const value = item[internalName];
  return value == null ? "" : String(value);
}

function getPersonFieldId(item, internalName) {
  if (!internalName) {
    return null;
  }

  const idKey = `${internalName}Id`;
  if (item[idKey] != null && item[idKey] !== "") {
    return Number(item[idKey]);
  }

  const raw = item[internalName];
  if (raw && typeof raw === "object" && raw.Id != null) {
    return Number(raw.Id);
  }

  if (raw != null && raw !== "") {
    const parsed = Number(raw);
    return Number.isNaN(parsed) ? null : parsed;
  }

  return null;
}

function findEnquiriesMissingRequiredInfo() {
  const aliases = resolveFieldAliases();
  const missingItems = [];
  const oneWeekAgo = Date.now() - 7 * 24 * 60 * 60 * 1000;
  const oneMonthAgo = Date.now() - 30 * 24 * 60 * 60 * 1000;
  const ownerMode = ownerFilter.value;

  allEnquiries.forEach((item) => {
    if (ownerMode === "mine") {
      const enquiryOwnerId = getPersonFieldId(item, aliases.enquiryOwner);
      if (enquiryOwnerId !== currentSpUserId) {
        return;
      }
    }

    const missing = {
      source: false,
      referralFrom: false,
      busDevNotes: false,
      reasonForLoss: false,
    };
    const tags = [];
    const statusValue = getFieldValue(item, aliases.status) || getFieldValue(item, aliases.currentStatus);
    const busDevNotes = getFieldValue(item, aliases.busDevNotes);

    if (isStatusRequiringLossDetails(statusValue)) {
      let hasLostGap = false;
      if (aliases.reasonForLoss && isBlank(getFieldValue(item, aliases.reasonForLoss))) {
        missing.reasonForLoss = true;
        hasLostGap = true;
      }
      if (aliases.busDevNotes && isBlank(busDevNotes)) {
        missing.busDevNotes = true;
        hasLostGap = true;
      }
      if (hasLostGap) {
        tags.push("Lost enquiry missing outcome details");
      }
    }

    const createdTime = Date.parse(item.Created || "");
    if (!Number.isFinite(createdTime) || createdTime < oneMonthAgo) {
      return;
    }

    const isOlderThanWeek = Number.isFinite(createdTime) && createdTime <= oneWeekAgo;
    if (isOlderThanWeek) {
      let hasReferralGap = false;
      if (aliases.source && isBlank(getFieldValue(item, aliases.source))) {
        missing.source = true;
        hasReferralGap = true;
      }
      if (aliases.referralFrom && isBlank(getFieldValue(item, aliases.referralFrom))) {
        missing.referralFrom = true;
        hasReferralGap = true;
      }
      if (aliases.busDevNotes && isBlank(busDevNotes)) {
        missing.busDevNotes = true;
        hasReferralGap = true;
      }
      if (hasReferralGap) {
        tags.push("Older than 7 days missing referral details");
      }
    }

    if (!Object.values(missing).some(Boolean)) {
      return;
    }

    missingItems.push({
      id: item.Id,
      name: getFieldValue(item, aliases.fullName) || "(No name)",
      status: statusValue || "-",
      created: item.Created ? new Date(item.Created).toLocaleDateString() : "Unknown date",
      tags,
      missing,
    });
  });

  return missingItems;
}

function renderDataQualityList() {
  const missingItems = findEnquiriesMissingRequiredInfo();
  if (!missingItems.length) {
    dataQualityCard.hidden = true;
    dataQualityCount.textContent = "";
    dataQualityList.innerHTML = "";
    return;
  }

  const rows = missingItems
    .map((item) => {
      const fields = [];
      if (item.missing.source) {
        fields.push(`
          <label class="field">Referral source
            <select data-field="source">
              ${buildSelectOptions(SOURCE_OPTIONS, "Select source")}
            </select>
          </label>
        `);
      }
      if (item.missing.referralFrom) {
        fields.push(`
          <label class="field">Referral from
            <input data-field="referralFrom" type="text" placeholder="Who referred this enquiry?" />
          </label>
        `);
      }
      if (item.missing.reasonForLoss) {
        fields.push(`
          <label class="field">Reason for loss
            <select data-field="reasonForLoss">
              ${buildSelectOptions(LOSS_REASON_OPTIONS, "Select reason")}
            </select>
          </label>
        `);
      }
      if (item.missing.busDevNotes) {
        fields.push(`
          <label class="field">Business development notes
            <textarea data-field="busDevNotes" placeholder="Add notes"></textarea>
          </label>
        `);
      }

      const tags = item.tags.map((tag) => `<span class="data-quality-tag">${escapeHtml(tag)}</span>`).join("");

      return `
        <div class="data-quality-row" data-id="${item.id}">
          <p class="data-quality-title">${escapeHtml(item.name)} (#${item.id})</p>
          <p class="data-quality-meta">${escapeHtml(item.status)} • Created ${escapeHtml(item.created)}</p>
          <p class="data-quality-link-wrap">
            <a class="data-quality-link" href="${getEnquiryItemUrl(item.id)}" target="_blank" rel="noopener noreferrer">Open in SharePoint</a>
          </p>
          <div class="data-quality-tags">${tags}</div>
          <div class="data-quality-fields">${fields.join("")}</div>
          <div class="data-quality-actions">
            <button class="data-quality-save-btn" type="button">Save required info</button>
          </div>
        </div>
      `;
    })
    .join("");

  dataQualityCount.textContent = `${missingItems.length} ${missingItems.length === 1 ? "enquiry" : "enquiries"} need updates`;
  dataQualityList.innerHTML = rows;
  dataQualityCard.hidden = false;
}

async function saveDataQualityRow(row) {
  const id = Number(row.dataset.id);
  if (!id) {
    throw new Error("Invalid enquiry ID.");
  }

  const item = allEnquiries.find((entry) => entry.Id === id);
  if (!item) {
    throw new Error("Could not find enquiry to update.");
  }

  const aliases = resolveFieldAliases();
  const payload = {
    __metadata: { type: listInfo.ListItemEntityTypeFullName },
  };
  let hasChanges = false;

  const sourceInput = row.querySelector('[data-field="source"]');
  if (sourceInput && aliases.source) {
    const value = sourceInput.value.trim();
    setInputError(sourceInput, !value);
    if (!value) {
      throw new Error("Referral source is required.");
    }
    payload[aliases.source] = value;
    hasChanges = true;
  }

  const referralFromInput = row.querySelector('[data-field="referralFrom"]');
  if (referralFromInput && aliases.referralFrom) {
    const value = referralFromInput.value.trim();
    setInputError(referralFromInput, !value);
    if (!value) {
      throw new Error("Referral from is required.");
    }
    payload[aliases.referralFrom] = value;
    hasChanges = true;
  }

  const reasonInput = row.querySelector('[data-field="reasonForLoss"]');
  if (reasonInput && aliases.reasonForLoss) {
    const value = reasonInput.value.trim();
    setInputError(reasonInput, !value);
    if (!value) {
      throw new Error("Reason for loss is required.");
    }
    payload[aliases.reasonForLoss] = value;
    hasChanges = true;
  }

  const busDevNotesInput = row.querySelector('[data-field="busDevNotes"]');
  if (busDevNotesInput && aliases.busDevNotes) {
    const value = busDevNotesInput.value.trim();
    setInputError(busDevNotesInput, !value);
    if (!value) {
      throw new Error("Business development notes are required.");
    }
    payload[aliases.busDevNotes] = value;
    hasChanges = true;
  }

  if (!hasChanges) {
    return;
  }

  const config = getConfig();
  const api = getSpApi(config);
  const digest = await getFormDigest(config);
  await api.request(`/_api/web/lists(guid'${listInfo.Id}')/items(${item.Id})`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE",
    },
    body: JSON.stringify(payload),
  });
}

function syncLocationCustomVisibility(prefix) {
  const select = $(`${prefix}LocationSelect`);
  const wrap = prefix === "q" ? qLocationCustomWrap : uLocationCustomWrap;
  wrap.hidden = select.value !== LOCATION_OTHER_VALUE;
}

function getLocationValue(prefix) {
  const select = $(`${prefix}LocationSelect`).value;
  if (!select) {
    return "";
  }
  if (select === LOCATION_OTHER_VALUE) {
    return $(`${prefix}LocationCustom`).value.trim();
  }
  return select;
}

function setLocationValue(prefix, value) {
  const select = $(`${prefix}LocationSelect`);
  const custom = $(`${prefix}LocationCustom`);
  const normalized = String(value || "").trim();

  if (!normalized) {
    select.value = "";
    custom.value = "";
    syncLocationCustomVisibility(prefix);
    return;
  }

  if (LOCATION_CHOICES.includes(normalized)) {
    select.value = normalized;
    custom.value = "";
    syncLocationCustomVisibility(prefix);
    return;
  }

  select.value = LOCATION_OTHER_VALUE;
  custom.value = normalized;
  syncLocationCustomVisibility(prefix);
}

function syncBusinessFieldsVisibility() {
  businessFields.hidden = showBusinessInfoBtn?.getAttribute("aria-pressed") !== "true";
}

function setBusinessInfoVisible(visible) {
  showBusinessInfoBtn.setAttribute("aria-pressed", visible ? "true" : "false");
  showBusinessInfoBtn.textContent = visible ? "Hide business development info" : "Show business development info";
  syncBusinessFieldsVisibility();
}

function setFilterPanelVisible(visible) {
  filterPanel.hidden = !visible;
  toggleFiltersBtn.textContent = visible ? "Hide filters" : "Show filters";
}

function parseOptionalNumber(value) {
  const trimmed = String(value || "").trim();
  if (!trimmed) {
    return null;
  }

  const parsed = Number(trimmed);
  return Number.isNaN(parsed) ? null : parsed;
}

function getConfig() {
  return {
    tenantId: FIXED_TENANT_ID,
    clientId: FIXED_CLIENT_ID,
    siteUrl: FIXED_SITE_URL,
    siteHost: new URL(FIXED_SITE_URL).origin,
  };
}

function getEnquiryItemUrl(itemId) {
  const listPath = `Lists/${encodeURIComponent(ENQUIRIES_LIST_TITLE)}`;
  return `${FIXED_SITE_URL}/${listPath}/DispForm.aspx?ID=${encodeURIComponent(String(itemId))}`;
}

function formatAuthError(error) {
  const code = error?.errorCode || error?.code || "";
  const message = error?.message || String(error);

  if (code.includes("popup_window_error") || code.includes("popup_window_open_error")) {
    return "Popup was blocked. Allow pop-ups for this site and try again.";
  }

  if (code.includes("user_cancelled")) {
    return "Sign-in was cancelled.";
  }

  if (message.includes("redirect_uri")) {
    return `Redirect URI mismatch. Register ${window.location.origin} in the Azure app as an SPA redirect URI.`;
  }

  return message;
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

async function getFormDigest(config) {
  return getSpApi(config).getFormDigest();
}

async function resolveCurrentSharePointUserId(config) {
  const email = account?.username?.trim();
  if (!email) {
    throw new Error("No signed-in user found for Enquiry owned by.");
  }
  currentSpUserId = await getSpApi(config).ensureCurrentUserId(email);
  if (!currentSpUserId) {
    throw new Error("Could not resolve current user in SharePoint.");
  }

  return currentSpUserId;
}

async function loadListInfo(config) {
  const api = getSpApi(config);
  listInfo = await api.resolveListByPath(ENQUIRIES_LIST_PATH);
  const fields = await api.getListFields(listInfo.Id);
  fieldMap = buildFieldMaps(fields);
}

function toSortableTimestamp(value) {
  const parsed = Date.parse(String(value || ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function sortEnquiries(items) {
  const mode = sortFilter.value;
  const aliases = resolveFieldAliases();

  return [...items].sort((first, second) => {
    if (mode === "updated") {
      return toSortableTimestamp(second.Modified) - toSortableTimestamp(first.Modified);
    }

    if (mode === "created") {
      return toSortableTimestamp(second.Created) - toSortableTimestamp(first.Created);
    }

    const firstStatus = (getFieldValue(first, aliases.status) || getFieldValue(first, aliases.currentStatus) || "").toLowerCase();
    const secondStatus = (getFieldValue(second, aliases.status) || getFieldValue(second, aliases.currentStatus) || "").toLowerCase();
    const statusOrder = firstStatus.localeCompare(secondStatus);
    if (statusOrder !== 0) {
      return statusOrder;
    }

    return toSortableTimestamp(second.Created) - toSortableTimestamp(first.Created);
  });
}

function renderEnquiries(items) {
  if (!items.length) {
    const ownerMode = ownerFilter.value;
    const scopeMode = scopeFilter.value;
    const ownerLabel = ownerMode === "mine" ? "owned by you" : "in this list";
    const scopeLabel = scopeMode === "active" ? "active " : "";
    listWrap.innerHTML = `<p class="muted">No ${scopeLabel}enquiries ${ownerLabel}.</p>`;
    return;
  }

  const aliases = resolveFieldAliases();
  const rows = items
    .map((item) => {
      const name = getFieldValue(item, aliases.fullName) || "(No name)";
      const status = getFieldValue(item, aliases.status) || getFieldValue(item, aliases.currentStatus) || "-";
      const caller = getFieldValue(item, aliases.callerName) || "-";
      return `
        <button class="list-row" data-id="${item.Id}" type="button">
          <span class="list-row-name">${escapeHtml(name)}</span>
          <span class="list-row-meta">${escapeHtml(status)} • Caller: ${escapeHtml(caller)}</span>
        </button>
      `;
    })
    .join("");

  listWrap.innerHTML = `<div class="list-grid">${rows}</div>`;

  listWrap.querySelectorAll(".list-row").forEach((button) => {
    button.addEventListener("click", () => {
      const id = Number(button.dataset.id);
      const item = enquiries.find((entry) => entry.Id === id);
      if (item) {
        selectEnquiry(item);
      }
    });
  });
}

function filterEnquiries(term) {
  const aliases = resolveFieldAliases();
  const query = term.trim().toLowerCase();

  if (!query) {
    renderEnquiries(sortEnquiries(enquiries));
    return;
  }

  const filtered = enquiries.filter((item) => {
    const values = [
      getFieldValue(item, aliases.fullName),
      getFieldValue(item, aliases.status),
      getFieldValue(item, aliases.currentStatus),
      getFieldValue(item, aliases.callerName),
      getFieldValue(item, aliases.phone),
    ]
      .join(" ")
      .toLowerCase();

    return values.includes(query);
  });

  renderEnquiries(sortEnquiries(filtered));
}

function applyEnquiryFilters() {
  const aliases = resolveFieldAliases();
  const ownerMode = ownerFilter.value;
  const scopeMode = scopeFilter.value;

  enquiries = allEnquiries.filter((item) => {
    if (ownerMode === "mine") {
      const enquiryOwnerId = getPersonFieldId(item, aliases.enquiryOwner);
      if (enquiryOwnerId !== currentSpUserId) {
        return false;
      }
    }

    if (scopeMode === "active") {
      const status = getFieldValue(item, aliases.status) || getFieldValue(item, aliases.currentStatus);
      if (isCompletedStatus(status)) {
        return false;
      }
    }

    return true;
  });

  if (selectedItem && !enquiries.some((item) => item.Id === selectedItem.Id)) {
    selectedItem = null;
    updatePanel.hidden = true;
    saveUpdateBtn.disabled = true;
    showBusinessInfoBtn.disabled = true;
    setBusinessInfoVisible(false);
    selectedEnquiryLink.hidden = true;
    selectedEnquiryLink.removeAttribute("href");
  }

  renderDataQualityList();
  filterEnquiries(searchInput.value);
}

function selectEnquiry(item) {
  selectedItem = item;
  updatePanel.hidden = false;
  saveUpdateBtn.disabled = false;

  const aliases = resolveFieldAliases();
  const name = getFieldValue(item, aliases.fullName) || "(No name)";
  const created = item.Created ? new Date(item.Created).toLocaleString() : "Unknown";

  selectedTitle.textContent = name;
  selectedMeta.textContent = `Created: ${created} • Item #${item.Id}`;
  selectedEnquiryLink.href = getEnquiryItemUrl(item.Id);
  selectedEnquiryLink.hidden = false;

  $("uStatus").value = getFieldValue(item, aliases.status) || "7. Initial Enquiry";
  $("uCurrentStatus").value = getFieldValue(item, aliases.currentStatus);
  $("uFollowUp").value = getFieldValue(item, aliases.followUp);
  $("uUpdates").value = getFieldValue(item, aliases.updates);
  $("uLikelihood").value = getFieldValue(item, aliases.likelihood);
  $("uRateHourly").value = getFieldValue(item, aliases.rateHourly);
  $("uSource").value = getFieldValue(item, aliases.source);
  setLocationValue("u", getFieldValue(item, aliases.location));
  $("uReferralFrom").value = getFieldValue(item, aliases.referralFrom);
  $("uReasonForLoss").value = getFieldValue(item, aliases.reasonForLoss);
  $("uBusDevNotes").value = getFieldValue(item, aliases.busDevNotes);
  $("uCallbackDate").value = fromIsoDate(getFieldValue(item, aliases.callbackDate));
  $("uAssessmentDate").value = fromIsoDate(getFieldValue(item, aliases.assessmentDate));
  showBusinessInfoBtn.disabled = false;
  syncLostFieldsVisibility();
  syncBusinessFieldsVisibility();
  updatePanel.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function loadEnquiries() {
  const config = getConfig();
  const api = getSpApi(config);

  if (!listInfo || !fieldMap) {
    await loadListInfo(config);
  }

  const aliases = resolveFieldAliases();
  const ownerId = await resolveCurrentSharePointUserId(config);
  const baseFields = ["Id", "Created", "Modified"];
  const dynamicFields = Object.entries(aliases)
    .filter(([key, value]) => key !== "enquiryOwner" && Boolean(value))
    .map(([, value]) => value);
  const ownerIdField = aliases.enquiryOwner ? `${aliases.enquiryOwner}Id` : null;
  const selectFields = [...new Set([...baseFields, ...dynamicFields, ownerIdField].filter(Boolean))];
  const query = selectFields.join(",");

  const data = await api.request(
    `/_api/web/lists(guid'${listInfo.Id}')/items?$top=200&$orderby=Created desc&$select=${encodeURIComponent(query)}`,
  );

  currentSpUserId = ownerId;
  allEnquiries = data.d.results;
  renderDataQualityList();
  applyEnquiryFilters();
}

async function createEnquiry() {
  const config = getConfig();
  const api = getSpApi(config);

  if (!listInfo || !fieldMap) {
    await loadListInfo(config);
  }

  const aliases = resolveFieldAliases();
  const payload = {
    __metadata: { type: listInfo.ListItemEntityTypeFullName },
  };

  const clientName = $("qClientName").value.trim();
  if (!clientName) {
    setInputError($("qClientName"), true);
    $("qClientName").focus();
    throw new Error("Client full name is required.");
  }
  setInputError($("qClientName"), false);

  const sourceValue = $("qSource").value;
  if (!sourceValue) {
    setInputError($("qSource"), true);
    $("qSource").focus();
    throw new Error("Referral source is required.");
  }
  setInputError($("qSource"), false);

  if (aliases.fullName) payload[aliases.fullName] = clientName;
  if (aliases.callerName) payload[aliases.callerName] = $("qCallerName").value.trim();
  if (aliases.relationship) payload[aliases.relationship] = $("qRelationship").value;
  if (aliases.enquiryOwner) {
    const ownerId = await resolveCurrentSharePointUserId(config);
    payload[`${aliases.enquiryOwner}Id`] = ownerId;
  }
  if (aliases.phone) payload[aliases.phone] = $("qPhone").value.trim();
  if (aliases.email) payload[aliases.email] = $("qEmail").value.trim();
  if (aliases.source) payload[aliases.source] = sourceValue;
  if (aliases.location) payload[aliases.location] = getLocationValue("q");
  if (aliases.referralFrom) payload[aliases.referralFrom] = $("qReferralFrom").value.trim();
  if (aliases.busDevNotes) payload[aliases.busDevNotes] = $("qBusDevNotes").value.trim();

  const statusValue = $("qStatus").value;
  if (aliases.status) payload[aliases.status] = statusValue;
  if (aliases.currentStatus) payload[aliases.currentStatus] = statusValue;

  const callbackDateIso = toIsoDate($("qCallbackDate").value);
  if (aliases.callbackDate && callbackDateIso) payload[aliases.callbackDate] = callbackDateIso;

  const assessmentDateIso = toIsoDate($("qAssessmentDate").value);
  if (aliases.assessmentDate && assessmentDateIso) payload[aliases.assessmentDate] = assessmentDateIso;

  const notes = $("qNotes").value.trim();
  if (aliases.background && notes) payload[aliases.background] = notes;
  if (aliases.updates && notes) payload[aliases.updates] = notes;

  const digest = await getFormDigest(config);
  await api.request(`/_api/web/lists(guid'${listInfo.Id}')/items`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
    },
    body: JSON.stringify(payload),
  });

  $("qClientName").value = "";
  $("qCallerName").value = "";
  $("qRelationship").value = "";
  $("qPhone").value = "";
  $("qEmail").value = "";
  $("qSource").value = "";
  $("qLocationSelect").value = "";
  $("qLocationCustom").value = "";
  $("qReferralFrom").value = "";
  $("qBusDevNotes").value = "";
  $("qCallbackDate").value = "";
  $("qAssessmentDate").value = "";
  $("qNotes").value = "";
  qClientTbcBtn.textContent = "TBC";
  qLocationCustomWrap.hidden = true;
  tbcAutoMode = false;
  setInputError($("qClientName"), false);
  setInputError($("qSource"), false);
  syncReferralDetailsVisibility();

  await loadEnquiries();
}

async function saveUpdate() {
  if (!selectedItem) {
    throw new Error("Choose an enquiry first.");
  }

  const config = getConfig();
  const api = getSpApi(config);
  const aliases = resolveFieldAliases();
  const payload = {
    __metadata: { type: listInfo.ListItemEntityTypeFullName },
  };

  const statusValue = $("uStatus").value.trim();
  if (aliases.status) payload[aliases.status] = statusValue;
  if (aliases.currentStatus) payload[aliases.currentStatus] = $("uCurrentStatus").value.trim() || statusValue;
  if (aliases.followUp) payload[aliases.followUp] = $("uFollowUp").value.trim();
  if (aliases.likelihood) payload[aliases.likelihood] = parseOptionalNumber($("uLikelihood").value);
  if (aliases.rateHourly) payload[aliases.rateHourly] = parseOptionalNumber($("uRateHourly").value);
  if (aliases.source) payload[aliases.source] = $("uSource").value;
  if (aliases.location) payload[aliases.location] = getLocationValue("u");
  if (aliases.referralFrom) payload[aliases.referralFrom] = $("uReferralFrom").value.trim();
  if (aliases.updates) payload[aliases.updates] = $("uUpdates").value.trim();
  if (aliases.reasonForLoss) payload[aliases.reasonForLoss] = isLostStatus(statusValue) ? $("uReasonForLoss").value : "";
  if (aliases.busDevNotes) payload[aliases.busDevNotes] = $("uBusDevNotes").value.trim();

  const callbackDateIso = toIsoDate($("uCallbackDate").value);
  if (aliases.callbackDate) payload[aliases.callbackDate] = callbackDateIso;

  const assessmentDateIso = toIsoDate($("uAssessmentDate").value);
  if (aliases.assessmentDate) payload[aliases.assessmentDate] = assessmentDateIso;

  const digest = await getFormDigest(config);
  await api.request(`/_api/web/lists(guid'${listInfo.Id}')/items(${selectedItem.Id})`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE",
    },
    body: JSON.stringify(payload),
  });

  await loadEnquiries();
  const refreshed = enquiries.find((item) => item.Id === selectedItem.Id);
  if (refreshed) {
    selectEnquiry(refreshed);
  }
}

async function handleSignIn() {
  try {
    signInBtn.disabled = true;
    setStatus("Signing in...");
    account = await authController.signIn({
      scopes: ["openid", "profile"],
      prompt: "select_account",
    });

    if (!account) {
      return;
    }

    setSignedInUi();
    await loadEnquiries();
    setStatus("Enquiries loaded.");
  } catch (error) {
    console.error(error);
    setStatus("Sign-in failed.", true, formatAuthError(error));
  } finally {
    signInBtn.disabled = false;
  }
}

async function restoreSessionOnLoad() {
  try {
    account = await authController.restoreSession();

    if (!account) {
      setSignedOutUi();
      setStatus("Please sign in to load enquiries.");
      return;
    }

    setSignedInUi();
    await loadEnquiries();
    setStatus("Session restored.");
  } catch (error) {
    console.error(error);
    setSignedOutUi();
    setStatus("Could not restore sign-in session.", true, formatAuthError(error));
  } finally {
    document.body.classList.remove("auth-pending");
  }
}

signInBtn?.addEventListener("click", handleSignIn);

signOutBtn?.addEventListener("click", async () => {
  try {
    signOutBtn.disabled = true;
    await authController.signOut();
  } catch (error) {
    console.error(error);
    setStatus(error?.message || "Sign-out failed.", true);
    signOutBtn.disabled = false;
  }
});

createEnquiryBtn?.addEventListener("click", async () => {
  try {
    createEnquiryBtn.disabled = true;
    setStatus("Creating enquiry...");
    await createEnquiry();
    setStatus("Enquiry created.");
  } catch (error) {
    console.error(error);
    setStatus("Could not create enquiry.", true, error.message || String(error));
  } finally {
    createEnquiryBtn.disabled = false;
  }
});

saveUpdateBtn?.addEventListener("click", async () => {
  try {
    saveUpdateBtn.disabled = true;
    setStatus("Saving update...");
    await saveUpdate();
    setStatus("Enquiry updated.");
  } catch (error) {
    console.error(error);
    setStatus("Could not save update.", true, error.message || String(error));
  } finally {
    saveUpdateBtn.disabled = false;
  }
});

searchInput?.addEventListener("input", () => {
  filterEnquiries(searchInput.value);
});

ownerFilter?.addEventListener("change", applyEnquiryFilters);
scopeFilter?.addEventListener("change", applyEnquiryFilters);
sortFilter?.addEventListener("change", () => filterEnquiries(searchInput.value));
toggleFiltersBtn?.addEventListener("click", () => setFilterPanelVisible(filterPanel.hidden));
showBusinessInfoBtn?.addEventListener("click", () => {
  if (showBusinessInfoBtn.disabled) {
    return;
  }
  setBusinessInfoVisible(showBusinessInfoBtn.getAttribute("aria-pressed") !== "true");
});

$("uStatus")?.addEventListener("change", syncLostFieldsVisibility);
$("qSource")?.addEventListener("change", syncReferralDetailsVisibility);
$("qLocationSelect")?.addEventListener("change", () => syncLocationCustomVisibility("q"));
$("uLocationSelect")?.addEventListener("change", () => syncLocationCustomVisibility("u"));
quickEntryToggleBtn?.addEventListener("click", () => setQuickEntryVisible(quickEntryContent.hidden));
dataQualityToggleBtn?.addEventListener("click", () => setDataQualityVisible(dataQualityContent.hidden));

qClientTbcBtn?.addEventListener("click", () => {
  tbcAutoMode = !tbcAutoMode;
  $("qClientName").value = tbcAutoMode ? "TBC" : "";
  qClientTbcBtn.textContent = tbcAutoMode ? "Clear" : "TBC";
  setInputError($("qClientName"), false);
});

document.querySelectorAll(".chip-btn[data-status]").forEach((button) => {
  button.addEventListener("click", () => {
    $("uStatus").value = button.dataset.status || "";
    syncLostFieldsVisibility();
  });
});

dataQualityList?.addEventListener("click", async (event) => {
  const button = event.target.closest(".data-quality-save-btn");
  if (!button) {
    return;
  }

  const row = button.closest(".data-quality-row");
  if (!row) {
    return;
  }

  try {
    button.disabled = true;
    setStatus("Saving required info...");
    await saveDataQualityRow(row);
    await loadEnquiries();
    setStatus("Required info saved.");
  } catch (error) {
    console.error(error);
    setStatus("Could not save required info.", true, error.message || String(error));
  } finally {
    button.disabled = false;
  }
});

setQuickEntryVisible(true);
setDataQualityVisible(true);
syncReferralDetailsVisibility();
syncLocationCustomVisibility("q");
syncLocationCustomVisibility("u");
syncLostFieldsVisibility();
void restoreSessionOnLoad();
