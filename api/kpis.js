const { requireApiAuth } = require("./_lib/require-api-auth");
const { createGraphAppClient } = require("./_lib/graph-app-client");

const DEFAULT_KPI_SITE_URL = "https://planwithcare.sharepoint.com/sites/NewBusiness";
const DEFAULT_KPI_LIST_NAME = "KPIs Thrive";
const DEFAULT_KPI_LIST_WEB_URL =
  "https://planwithcare.sharepoint.com/sites/NewBusiness/Lists/KPIs%20Thrive/AllItems.aspx?viewid=091674c0%2D96cc%2D48ad%2Dbcac%2D787253e3cf31&env=WebViewList";
const DEFAULT_RECRUITMENT_SITE_URL = "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency";
const DEFAULT_RECRUITMENT_LIST_NAME = "Associate Recruitment";
const DEFAULT_ENQUIRIES_SITE_URL = "https://planwithcare.sharepoint.com/sites/ThriveCalls";
const DEFAULT_ENQUIRIES_LIST_NAME = "Enquiries Log";
const DEFAULT_ENQUIRIES_LIST_WEB_URL =
  "https://planwithcare.sharepoint.com/sites/ThriveCalls/Lists/Enquiries%20Log/AllItems.aspx";
const QUARTER_WEEK_COUNT = 13;
const ASSESSMENT_OUTCOME_MONTHS = 3;
const EDITABLE_KPI_FIELDS = new Set(["utilisationNotes", "trainingCompletion", "cqcReadiness"]);

const KPI_FIELD_DEFINITIONS = {
  weekCommencing: ["Week Commenceing", "Week Commencing", "WeekCommenceing"],
  activeContractedHours: ["Active Contracted Hours", "ContractedHours"],
  hoursDelivered: ["Hours Delivered", "Targetaveraging500hoursperweekby"],
  hoursDeliveredPercent: ["Hours Delivered %", "Hours Delivered Percentage", "HoursDeliveredPercent"],
  droppedHoursReasons: ["dropped hours reasons", "Dropped Hours Reasons", "droppedhoursreasons"],
  explanatoryNotes: ["Explanatory Notes", "ExplanatoryNotes"],
  subscriptionHours: ["Subscription 'Hours'", "Subscription Hours", "SubscriptionHours"],
  totalHours: ["Total Hours", "TotalHours"],
  utilisationPercent: ["Utilisation%", "Utilisation %", "Utilisation", "Utilization%", "Utilization %", "Utilization"],
  utilisationNotes: ["Utilisation Notes", "Utilisation Note", "Utilization Notes", "Utilization Note"],
  hoursWon: ["Hours won (mth)", "Hours Won", "Newweeklyhourswon"],
  hoursLost: ["Hours cancelled  (mth)", "Hours cancelled (mth)", "Hours lost", "Hourscancelled"],
  pendingHours: ["Pending Hours / wk", "Pending Hours", "PendingHours"],
  pendingHoursDetail: ["Pending Hours Detail", "PendingHoursDetail"],
  activeEnquiries: ["Active Enquiries", "ActiveEnquiries"],
  enquiriesTotal: ["Enquiries Total /wk", "Depreciated-Enquiries per week", "Enquiries per week", "Enquiriesperweek"],
  enquiriesSolicitor: ["Enquiries - Solicitor", "Enquiries_x002d_Solicitor"],
  enquiriesConsumer: ["Enquiries - Consumer", "Enquiries_x002d_Consumer"],
  enquiryConversion: ["Enquiry Conversion", "Enquiry Conversion %", "EnquiryConversion", "EnquiryConversionPercent"],
  instagramFollowers: ["Instagram Followers", "InstagramFollowers"],
  facebookFollowers: ["Facebook Followers", "FacebookFollowers"],
  newsletterSubscribers: ["Thrive Newletter Subs", "Thrive Newsletter Subs", "ThriveNewletterSubs"],
  webVisits: ["Web Visits", "WebVisits"],
  domainAuthorityThrive: ["Domain Authority - Thrive", "DomainAuthority_x002d_Thrive"],
  domainAuthorityPwc: ["Domain Authority - PWC", "DomainAuthority_x002d_PWC"],
  firstRoundInterviews: ["1st Round Interviews", "_x0031_stRoundInterviews"],
  trainingCompletion: ["Training Completion", "TrainingCompletion"],
  cqcReadiness: ["CQC Readiness", "CQC Readiness ", "CQCReadiness"],
};

const RECRUITMENT_FIELD_DEFINITIONS = {
  candidateName: ["Title", "Candidate Name"],
  status: ["Status"],
  active: ["Active"],
};

const ENQUIRY_FIELD_DEFINITIONS = {
  title: ["Title", "Name", "Client Name", "Enquirer Name"],
  status: ["Status"],
  currentStatus: ["Current Status"],
  modified: ["Modified"],
};

function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeToken(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/&amp;/g, "and")
    .replace(/%/g, "percent")
    .replace(/[^a-z0-9]/g, "");
}

function parseSharePointListConfig(siteUrlValue, listName, missingConfigMessage) {
  if (!siteUrlValue || !listName) {
    throw new Error(missingConfigMessage);
  }

  const siteUrl = new URL(siteUrlValue);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    throw new Error("Configured SharePoint site URL must include a site path.");
  }

  return {
    hostName: siteUrl.hostname,
    siteBaseUrl: `${siteUrl.protocol}//${siteUrl.host}${sitePath}`,
    sitePath,
    listName,
  };
}

function parseKpiConfig() {
  const listWebUrl = cleanText(process.env.SHAREPOINT_KPI_LIST_WEB_URL || DEFAULT_KPI_LIST_WEB_URL);
  return {
    ...parseSharePointListConfig(
      cleanText(process.env.SHAREPOINT_KPI_SITE_URL || DEFAULT_KPI_SITE_URL),
      cleanText(process.env.SHAREPOINT_KPI_LIST_NAME || DEFAULT_KPI_LIST_NAME),
      "Missing SHAREPOINT_KPI_SITE_URL or SHAREPOINT_KPI_LIST_NAME."
    ),
    listWebUrl,
    listPath: parseListPathFromWebUrl(listWebUrl),
  };
}

function parseRecruitmentConfig() {
  return parseSharePointListConfig(
    cleanText(process.env.SHAREPOINT_RECRUITMENT_SITE_URL || DEFAULT_RECRUITMENT_SITE_URL),
    cleanText(process.env.SHAREPOINT_RECRUITMENT_LIST_NAME || DEFAULT_RECRUITMENT_LIST_NAME),
    "Missing SHAREPOINT_RECRUITMENT_SITE_URL or SHAREPOINT_RECRUITMENT_LIST_NAME."
  );
}

function parseEnquiriesConfig() {
  const listWebUrl = cleanText(process.env.SHAREPOINT_ENQUIRIES_LIST_WEB_URL || DEFAULT_ENQUIRIES_LIST_WEB_URL);
  return {
    ...parseSharePointListConfig(
      cleanText(process.env.SHAREPOINT_ENQUIRIES_SITE_URL || DEFAULT_ENQUIRIES_SITE_URL),
      cleanText(process.env.SHAREPOINT_ENQUIRIES_LIST_NAME || DEFAULT_ENQUIRIES_LIST_NAME),
      "Missing SHAREPOINT_ENQUIRIES_SITE_URL or SHAREPOINT_ENQUIRIES_LIST_NAME."
    ),
    listWebUrl,
    listPath: parseListPathFromWebUrl(listWebUrl),
  };
}

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function normalizeSharePointPath(value) {
  const text = cleanText(value);
  if (!text) {
    return "";
  }
  try {
    return decodeURIComponent(text).replace(/\/+/g, "/").replace(/\/$/, "").toLowerCase();
  } catch {
    return text.replace(/\/+/g, "/").replace(/\/$/, "").toLowerCase();
  }
}

function parseListPathFromWebUrl(webUrl) {
  const text = cleanText(webUrl);
  if (!text) {
    return "";
  }
  try {
    const url = new URL(text);
    const path = normalizeSharePointPath(url.pathname);
    const allItemsIndex = path.toLowerCase().lastIndexOf("/allitems.aspx");
    return allItemsIndex >= 0 ? path.slice(0, allItemsIndex) : path;
  } catch {
    return "";
  }
}

async function resolveSiteId(graphClient, hostName, sitePath) {
  const url = `https://graph.microsoft.com/v1.0/sites/${hostName}:${sitePath}?$select=id`;
  const payload = await graphClient.fetchJson(url);
  if (!payload?.id) {
    throw new Error("Could not resolve SharePoint site id.");
  }
  return payload.id;
}

async function findListByWebUrlPath(graphClient, siteId, listPath) {
  const targetPath = normalizeSharePointPath(listPath);
  if (!targetPath) {
    return null;
  }

  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=id,displayName,webUrl&$top=200`;
  while (nextUrl) {
    const payload = await graphClient.fetchJson(nextUrl);
    const lists = Array.isArray(payload?.value) ? payload.value : [];
    const match = lists.find((list) => {
      const webUrl = cleanText(list?.webUrl);
      if (!webUrl) {
        return false;
      }
      try {
        const path = normalizeSharePointPath(new URL(webUrl).pathname);
        return path === targetPath || path.endsWith(targetPath);
      } catch {
        return false;
      }
    });
    if (match?.id) {
      return match;
    }
    nextUrl = cleanText(payload?.["@odata.nextLink"]);
  }

  return null;
}

async function resolveList(graphClient, siteId, listName, options = {}) {
  const params = new URLSearchParams({
    $select: "id,displayName,webUrl",
    $filter: `displayName eq ${quoteODataString(listName)}`,
    $top: "1",
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?${params.toString()}`;
  const payload = await graphClient.fetchJson(url);
  let list = Array.isArray(payload?.value) ? payload.value[0] : null;
  if (!list?.id) {
    list = await findListByWebUrlPath(graphClient, siteId, options.listPath);
  }
  if (!list?.id) {
    const listPath = cleanText(options.listPath);
    const suffix = listPath ? ` or path '${listPath}'` : "";
    throw new Error(`Could not find SharePoint list '${listName}'${suffix}.`);
  }
  return {
    id: String(list.id),
    webUrl: cleanText(list.webUrl),
  };
}

async function resolveListColumns(graphClient, siteId, listId) {
  const columns = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns?$select=name,displayName&$top=200`;
  while (nextUrl) {
    const payload = await graphClient.fetchJson(nextUrl);
    columns.push(...(Array.isArray(payload?.value) ? payload.value : []));
    nextUrl = cleanText(payload?.["@odata.nextLink"]);
  }
  return columns;
}

function createFieldMap(columns, definitions) {
  const byToken = new Map();
  for (const column of Array.isArray(columns) ? columns : []) {
    const name = cleanText(column?.name);
    const displayName = cleanText(column?.displayName);
    if (!name) {
      continue;
    }
    for (const value of [name, displayName]) {
      const token = normalizeToken(value);
      if (token && !byToken.has(token)) {
        byToken.set(token, name);
      }
    }
  }

  const fieldMap = {};
  for (const [key, candidates] of Object.entries(definitions)) {
    fieldMap[key] = "";
    for (const candidate of candidates) {
      const match = byToken.get(normalizeToken(candidate));
      if (match) {
        fieldMap[key] = match;
        break;
      }
    }
  }
  return fieldMap;
}

function uniqueFieldNames(fieldMap) {
  return Array.from(new Set(Object.values(fieldMap).map(cleanText).filter(Boolean)));
}

function getFieldValue(row, key) {
  const fieldName = row?.fieldMap?.[key];
  if (!fieldName) {
    return "";
  }
  return row?.fields?.[fieldName];
}

function hasValue(value) {
  if (value === null || value === undefined) {
    return false;
  }
  if (typeof value === "string") {
    return value.trim() !== "";
  }
  return true;
}

function parseNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  const text = cleanText(value).replace(/,/g, "");
  if (!text) {
    return null;
  }
  const match = text.match(/-?\d+(?:\.\d+)?/);
  if (!match) {
    return null;
  }
  const number = Number(match[0]);
  return Number.isFinite(number) ? number : null;
}

function parsePercent(value) {
  const number = parseNumber(value);
  if (number === null) {
    return null;
  }
  if (typeof value === "number" && number > 0 && number <= 1) {
    return number * 100;
  }
  return number;
}

function normalizeEditableKpiValue(key, value, latestRow, fieldName) {
  if (key === "trainingCompletion") {
    const percent = parsePercent(value);
    if (percent === null) {
      const error = new Error("Training Completion must be a percentage.");
      error.status = 400;
      error.code = "INVALID_KPI_VALUE";
      throw error;
    }
    const currentRaw = latestRow?.fields?.[fieldName];
    if (typeof currentRaw === "number" && currentRaw > 0 && currentRaw <= 1) {
      return percent / 100;
    }
    return percent;
  }
  return cleanText(value);
}

function toBoolean(value) {
  if (value === true || value === false) {
    return value;
  }
  const normalized = cleanText(value).toLowerCase();
  if (!normalized) {
    return false;
  }
  return ["true", "1", "yes", "y"].includes(normalized);
}

function formatWeek(value) {
  const date = new Date(value);
  if (!value || Number.isNaN(date.getTime())) {
    return "";
  }
  return new Intl.DateTimeFormat("en-GB", {
    day: "numeric",
    month: "short",
    year: "numeric",
    timeZone: "Europe/London",
  }).format(date);
}

function normalizeKpiRow(item, fieldMap) {
  const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
  const weekCommencing = getFieldValue({ fields, fieldMap }, "weekCommencing");
  return {
    id: cleanText(item?.id),
    fields,
    fieldMap,
    weekCommencing: cleanText(weekCommencing),
    weekLabel: formatWeek(weekCommencing),
    weekTime: new Date(weekCommencing).getTime() || 0,
  };
}

function deriveValue(rows, latestRow, key, derive) {
  const latestWeek = latestRow?.weekCommencing || "";
  for (const row of rows) {
    const raw = typeof derive === "function" ? derive(row) : getFieldValue(row, key);
    if (!hasValue(raw)) {
      continue;
    }
    const sourceWeek = row.weekCommencing || "";
    return {
      key,
      value: raw,
      sourceWeek,
      sourceWeekLabel: row.weekLabel,
      stale: Boolean(latestWeek && sourceWeek && sourceWeek !== latestWeek),
    };
  }
  return {
    key,
    value: "",
    sourceWeek: "",
    sourceWeekLabel: "",
    stale: false,
  };
}

function deriveHoursDeliveredPercent(row) {
  const stored = getFieldValue(row, "hoursDeliveredPercent");
  if (hasValue(stored)) {
    return parsePercent(stored);
  }
  const delivered = parseNumber(getFieldValue(row, "hoursDelivered"));
  const contracted = parseNumber(getFieldValue(row, "activeContractedHours"));
  if (delivered === null || contracted === null || contracted === 0) {
    return "";
  }
  return (delivered / contracted) * 100;
}

function deriveTotalHours(row) {
  const stored = getFieldValue(row, "totalHours");
  if (hasValue(stored)) {
    return parseNumber(stored);
  }
  const delivered = parseNumber(getFieldValue(row, "hoursDelivered"));
  const subscription = parseNumber(getFieldValue(row, "subscriptionHours"));
  if (delivered === null && subscription === null) {
    return "";
  }
  return (delivered || 0) + (subscription || 0);
}

function deriveEnquiriesTotal(row) {
  const stored = getFieldValue(row, "enquiriesTotal");
  if (hasValue(stored)) {
    return parseNumber(stored);
  }
  const solicitor = parseNumber(getFieldValue(row, "enquiriesSolicitor"));
  const consumer = parseNumber(getFieldValue(row, "enquiriesConsumer"));
  if (solicitor === null && consumer === null) {
    return "";
  }
  return (solicitor || 0) + (consumer || 0);
}

function deriveUtilisationPercent(row) {
  const stored = getFieldValue(row, "utilisationPercent");
  if (!hasValue(stored)) {
    return "";
  }
  return parsePercent(stored);
}

function deriveEnquiryConversion(row) {
  const stored = getFieldValue(row, "enquiryConversion");
  if (!hasValue(stored)) {
    return "";
  }
  return parsePercent(stored);
}

function resolveMetrics(rows) {
  const latestRow = rows[0] || null;
  const metric = (key, derive) => deriveValue(rows, latestRow, key, derive);
  return {
    latestWeek: latestRow?.weekCommencing || "",
    latestWeekLabel: latestRow?.weekLabel || "",
    values: {
      hoursDeliveredPercent: metric("hoursDeliveredPercent", deriveHoursDeliveredPercent),
      hoursDelivered: metric("hoursDelivered"),
      activeContractedHours: metric("activeContractedHours"),
      droppedHoursReasons: metric("droppedHoursReasons"),
      explanatoryNotes: metric("explanatoryNotes"),
      subscriptionHours: metric("subscriptionHours"),
      totalHours: metric("totalHours", deriveTotalHours),
      utilisationPercent: metric("utilisationPercent", deriveUtilisationPercent),
      utilisationNotes: metric("utilisationNotes"),
      hoursWon: metric("hoursWon"),
      hoursLost: metric("hoursLost"),
      pendingHours: metric("pendingHours"),
      pendingHoursDetail: metric("pendingHoursDetail"),
      activeEnquiries: metric("activeEnquiries"),
      enquiriesTotal: metric("enquiriesTotal", deriveEnquiriesTotal),
      enquiriesSolicitor: metric("enquiriesSolicitor"),
      enquiriesConsumer: metric("enquiriesConsumer"),
      enquiryConversion: metric("enquiryConversion", deriveEnquiryConversion),
      instagramFollowers: metric("instagramFollowers"),
      facebookFollowers: metric("facebookFollowers"),
      newsletterSubscribers: metric("newsletterSubscribers"),
      webVisits: metric("webVisits"),
      domainAuthorityThrive: metric("domainAuthorityThrive"),
      domainAuthorityPwc: metric("domainAuthorityPwc"),
      firstRoundInterviews: metric("firstRoundInterviews"),
      trainingCompletion: metric("trainingCompletion"),
      cqcReadiness: metric("cqcReadiness"),
    },
  };
}

function buildTrendSeries(rows) {
  return rows
    .slice(0, QUARTER_WEEK_COUNT)
    .map((row) => ({
      weekCommencing: row.weekCommencing,
      weekLabel: row.weekLabel,
      hoursDeliveredPercent: parseNumber(deriveHoursDeliveredPercent(row)),
      totalHours: parseNumber(deriveTotalHours(row)),
      utilisationPercent: parseNumber(deriveUtilisationPercent(row)),
      hoursWon: parseNumber(getFieldValue(row, "hoursWon")),
      hoursLost: parseNumber(getFieldValue(row, "hoursLost")),
      pendingHours: parseNumber(getFieldValue(row, "pendingHours")),
      activeEnquiries: parseNumber(getFieldValue(row, "activeEnquiries")),
      enquiriesTotal: parseNumber(deriveEnquiriesTotal(row)),
      enquiriesSolicitor: parseNumber(getFieldValue(row, "enquiriesSolicitor")),
      enquiriesConsumer: parseNumber(getFieldValue(row, "enquiriesConsumer")),
      enquiryConversion: parseNumber(deriveEnquiryConversion(row)),
      instagramFollowers: parseNumber(getFieldValue(row, "instagramFollowers")),
      facebookFollowers: parseNumber(getFieldValue(row, "facebookFollowers")),
      newsletterSubscribers: parseNumber(getFieldValue(row, "newsletterSubscribers")),
      webVisits: parseNumber(getFieldValue(row, "webVisits")),
      domainAuthorityThrive: parseNumber(getFieldValue(row, "domainAuthorityThrive")),
      domainAuthorityPwc: parseNumber(getFieldValue(row, "domainAuthorityPwc")),
    }))
    .reverse();
}

async function fetchKpiRows(graphClient, config) {
  const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
  const list = await resolveList(graphClient, siteId, config.listName, { listPath: config.listPath });
  const columns = await resolveListColumns(graphClient, siteId, list.id);
  const fieldMap = createFieldMap(columns, KPI_FIELD_DEFINITIONS);
  const selectFields = uniqueFieldNames(fieldMap);
  const expand = selectFields.length ? `fields($select=${selectFields.join(",")})` : "fields";
  const params = new URLSearchParams({
    $top: "40",
    $expand: expand,
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${list.id}/items?${params.toString()}`;
  const items = await graphClient.fetchAllPages(url);
  const rows = items
    .map((item) => normalizeKpiRow(item, fieldMap))
    .filter((row) => row.weekTime > 0)
    .sort((a, b) => b.weekTime - a.weekTime);
  return {
    rows,
    listUrl: list.webUrl || config.listWebUrl,
    fieldMap,
  };
}

async function updateLatestKpiFields(graphClient, config, updates) {
  if (!updates || typeof updates !== "object" || Array.isArray(updates)) {
    const error = new Error("KPI update payload must include fields.");
    error.status = 400;
    error.code = "INVALID_KPI_UPDATE";
    throw error;
  }

  const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
  const list = await resolveList(graphClient, siteId, config.listName, { listPath: config.listPath });
  const columns = await resolveListColumns(graphClient, siteId, list.id);
  const fieldMap = createFieldMap(columns, KPI_FIELD_DEFINITIONS);
  const selectFields = uniqueFieldNames(fieldMap);
  const expand = selectFields.length ? `fields($select=${selectFields.join(",")})` : "fields";
  const params = new URLSearchParams({
    $top: "40",
    $expand: expand,
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${list.id}/items?${params.toString()}`;
  const items = await graphClient.fetchAllPages(url);
  const rows = items
    .map((item) => normalizeKpiRow(item, fieldMap))
    .filter((row) => row.weekTime > 0)
    .sort((a, b) => b.weekTime - a.weekTime);
  const latestRow = rows[0] || null;
  if (!latestRow?.id) {
    const error = new Error("Could not find the latest KPI row to update.");
    error.status = 404;
    error.code = "KPI_LATEST_ROW_NOT_FOUND";
    throw error;
  }

  const patch = {};
  for (const [key, value] of Object.entries(updates)) {
    if (!EDITABLE_KPI_FIELDS.has(key)) {
      const error = new Error(`KPI field '${key}' cannot be edited here.`);
      error.status = 400;
      error.code = "KPI_FIELD_NOT_EDITABLE";
      throw error;
    }
    const fieldName = fieldMap[key];
    if (!fieldName) {
      const error = new Error(`Could not find SharePoint column for KPI field '${key}'.`);
      error.status = 400;
      error.code = "KPI_COLUMN_NOT_FOUND";
      throw error;
    }
    patch[fieldName] = normalizeEditableKpiValue(key, value, latestRow, fieldName);
  }

  if (!Object.keys(patch).length) {
    const error = new Error("No editable KPI fields were provided.");
    error.status = 400;
    error.code = "EMPTY_KPI_UPDATE";
    throw error;
  }

  await graphClient.fetchJson(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${list.id}/items/${encodeURIComponent(latestRow.id)}/fields`,
    {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(patch),
    }
  );

  return {
    updated: true,
    latestWeek: latestRow.weekCommencing,
    latestWeekLabel: latestRow.weekLabel,
    itemId: latestRow.id,
  };
}

function normalizeRecruitmentItem(item, fieldMap) {
  const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
  const row = { fields, fieldMap };
  return {
    candidateName: cleanText(getFieldValue(row, "candidateName")),
    status: cleanText(getFieldValue(row, "status")),
    active: toBoolean(getFieldValue(row, "active")),
  };
}

function firstName(fullName) {
  return cleanText(fullName).split(/\s+/).find(Boolean) || "";
}

function subtractMonths(date, months) {
  const next = new Date(date);
  next.setMonth(next.getMonth() - months);
  return next;
}

function getValidDate(value) {
  const date = new Date(value);
  return value && !Number.isNaN(date.getTime()) ? date : null;
}

function normalizeEnquiryStatus(value) {
  return cleanText(value).toLowerCase();
}

function isWonAfterAssessmentStatus(status) {
  return normalizeEnquiryStatus(status) === "won";
}

function isLostAfterAssessmentStatus(status) {
  const normalized = normalizeEnquiryStatus(status);
  return normalized.includes("lost") && normalized.includes("post assessment");
}

function isOnHoldEnquiryStatus(status) {
  return normalizeEnquiryStatus(status).includes("on hold");
}

function isActiveEnquiryStatus(status) {
  const normalized = normalizeEnquiryStatus(status);
  return normalized === "active" || normalized.includes("active enquiry") || normalized.includes("current active");
}

function normalizeEnquiryDetailItem(item, fieldMap) {
  const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
  const row = { fields, fieldMap };
  const modifiedDate = getValidDate(fields.Modified || item?.lastModifiedDateTime || getFieldValue(row, "modified"));
  const status = getFieldValue(row, "status") || getFieldValue(row, "currentStatus");
  return {
    title: cleanText(getFieldValue(row, "title") || fields.Title || `Item ${item?.id || ""}`),
    status: cleanText(status),
    modified: modifiedDate ? modifiedDate.toISOString() : "",
    modifiedLabel: modifiedDate ? formatWeek(modifiedDate.toISOString()) : "",
    webUrl: cleanText(item?.webUrl),
  };
}

async function fetchAcceptedOnboarding(graphClient) {
  const config = parseRecruitmentConfig();
  const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
  const list = await resolveList(graphClient, siteId, config.listName);
  const columns = await resolveListColumns(graphClient, siteId, list.id);
  const fieldMap = createFieldMap(columns, RECRUITMENT_FIELD_DEFINITIONS);
  const selectFields = uniqueFieldNames(fieldMap);
  const expand = selectFields.length ? `fields($select=${selectFields.join(",")})` : "fields";
  const params = new URLSearchParams({
    $top: "200",
    $expand: expand,
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${list.id}/items?${params.toString()}`;
  const items = await graphClient.fetchAllPages(url);
  const accepted = items
    .map((item) => normalizeRecruitmentItem(item, fieldMap))
    .filter((item) => item.status.toLowerCase() === "accepted" && item.active !== false)
    .map((item) => firstName(item.candidateName))
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));
  return {
    count: accepted.length,
    firstNames: accepted,
  };
}

async function fetchEnquiryLogItems(graphClient) {
  const config = parseEnquiriesConfig();
  const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
  const list = await resolveList(graphClient, siteId, config.listName, { listPath: config.listPath });
  const columns = await resolveListColumns(graphClient, siteId, list.id);
  const fieldMap = createFieldMap(columns, ENQUIRY_FIELD_DEFINITIONS);
  const selectFields = uniqueFieldNames(fieldMap);
  const expand = selectFields.length ? `fields($select=${selectFields.join(",")})` : "fields";
  const params = new URLSearchParams({
    $top: "200",
    $expand: expand,
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${list.id}/items?${params.toString()}`;
  return {
    items: await graphClient.fetchAllPages(url),
    fieldMap,
    listUrl: list.webUrl || config.listWebUrl,
  };
}

async function fetchActiveEnquiries(graphClient) {
  const { items, fieldMap, listUrl } = await fetchEnquiryLogItems(graphClient);
  const activeItems = items
    .map((item) => normalizeEnquiryDetailItem(item, fieldMap))
    .filter((item) => isActiveEnquiryStatus(item.status))
    .sort((a, b) => a.title.localeCompare(b.title));
  return {
    count: activeItems.length,
    items: activeItems,
    listUrl,
  };
}

async function fetchActiveEnquiriesSafe(graphClient) {
  try {
    return await fetchActiveEnquiries(graphClient);
  } catch (error) {
    console.warn("[kpis] Active enquiries unavailable", {
      message: error?.message || String(error),
      code: error?.code || "",
      status: error?.status || "",
    });
    return {
      count: 0,
      items: [],
      listUrl: "",
      unavailable: true,
      warning: error?.message || "Could not load active enquiries.",
    };
  }
}

async function fetchEnquiryAssessmentOutcome(graphClient) {
  const { items, fieldMap, listUrl } = await fetchEnquiryLogItems(graphClient);
  const startDate = subtractMonths(new Date(), ASSESSMENT_OUTCOME_MONTHS);
  startDate.setHours(0, 0, 0, 0);

  let won = 0;
  let lost = 0;
  let onHold = 0;
  const detail = {
    won: [],
    lost: [],
    onHold: [],
  };
  for (const item of items) {
    const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
    const row = { fields, fieldMap };
    const modifiedDate = getValidDate(fields.Modified || item?.lastModifiedDateTime || getFieldValue(row, "modified"));
    if (!modifiedDate || modifiedDate < startDate) {
      continue;
    }

    const status = getFieldValue(row, "status") || getFieldValue(row, "currentStatus");
    const detailItem = normalizeEnquiryDetailItem(item, fieldMap);
    if (isWonAfterAssessmentStatus(status)) {
      won += 1;
      detail.won.push(detailItem);
      continue;
    }
    if (isLostAfterAssessmentStatus(status)) {
      lost += 1;
      detail.lost.push(detailItem);
      continue;
    }
    if (isOnHoldEnquiryStatus(status)) {
      onHold += 1;
      detail.onHold.push(detailItem);
    }
  }

  const assessedOutcomes = won + lost;
  return {
    won,
    lost,
    onHold,
    assessedOutcomes,
    winPercent: assessedOutcomes ? (won / assessedOutcomes) * 100 : null,
    months: ASSESSMENT_OUTCOME_MONTHS,
    startDate: startDate.toISOString(),
    startDateLabel: formatWeek(startDate.toISOString()),
    listUrl,
    detail,
  };
}

async function fetchEnquiryAssessmentOutcomeSafe(graphClient) {
  try {
    return await fetchEnquiryAssessmentOutcome(graphClient);
  } catch (error) {
    console.warn("[kpis] Enquiry assessment outcome unavailable", {
      message: error?.message || String(error),
      code: error?.code || "",
      status: error?.status || "",
    });
    return {
      won: 0,
      lost: 0,
      onHold: 0,
      assessedOutcomes: 0,
      winPercent: null,
      months: ASSESSMENT_OUTCOME_MONTHS,
      startDate: "",
      startDateLabel: "",
      listUrl: "",
      detail: { won: [], lost: [], onHold: [] },
      unavailable: true,
      warning: error?.message || "Could not load Enquiries Log.",
    };
  }
}

function mapGraphError(error) {
  return {
    status: Number(error?.status) || 502,
    payload: {
      error: {
        code: cleanText(error?.code) || "KPI_REQUEST_FAILED",
        message: error?.message || "KPI request failed.",
      },
    },
  };
}

module.exports = async (req, res) => {
  if (req.method !== "GET" && req.method !== "PATCH") {
    res.status(405).json({
      error: { code: "METHOD_NOT_ALLOWED", message: "Method Not Allowed" },
    });
    return;
  }

  if (!(await requireApiAuth(req, res))) {
    return;
  }

  try {
    const graphClient = createGraphAppClient();
    const kpiConfig = parseKpiConfig();
    if (req.method === "PATCH") {
      const result = await updateLatestKpiFields(graphClient, kpiConfig, req.body?.fields || {});
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json(result);
      return;
    }

    const { rows, listUrl, fieldMap } = await fetchKpiRows(graphClient, kpiConfig);
    const metrics = resolveMetrics(rows);
    const onboarding = await fetchAcceptedOnboarding(graphClient);
    const enquiryAssessmentOutcome = await fetchEnquiryAssessmentOutcomeSafe(graphClient);
    const activeEnquiries = await fetchActiveEnquiriesSafe(graphClient);

    res.setHeader("Cache-Control", "private, max-age=60");
    res.status(200).json({
      listUrl,
      fieldMap,
      latestWeek: metrics.latestWeek,
      latestWeekLabel: metrics.latestWeekLabel,
      values: metrics.values,
      trendSeries: buildTrendSeries(rows),
      onboarding,
      enquiryAssessmentOutcome,
      activeEnquiries,
      rowCount: rows.length,
      refreshedAt: new Date().toISOString(),
    });
  } catch (error) {
    console.error("[kpis] KPI dashboard request failed", {
      message: error?.message || String(error),
      code: error?.code || "",
      status: error?.status || "",
    });
    const mapped = mapGraphError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
