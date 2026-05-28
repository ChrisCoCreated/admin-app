const { createGraphAppClient } = require("./graph-app-client");
const { parseEmailList, sendGraphMail } = require("./graph-mailer");

const DEFAULT_ENQUIRIES_SITE_URL = "https://planwithcare.sharepoint.com/sites/ThriveCalls";
const DEFAULT_ENQUIRIES_LIST_NAME = "Enquiries Log";
const DEFAULT_ENQUIRIES_LIST_WEB_URL =
  "https://planwithcare.sharepoint.com/sites/ThriveCalls/Lists/Enquiries%20Log/AllItems.aspx";
const OVERDUE_DAYS = 7;

const ENQUIRY_FIELD_DEFINITIONS = {
  title: ["Clients Full Name", "Title", "Name", "Client Name", "Enquirer Name"],
  callerName: ["Person Making Enquiry"],
  status: ["Status"],
  currentStatus: ["Current Status"],
  enquiryOwner: ["Enquiry owned by"],
  followUp: ["Follow up", "Followup"],
  updates: ["Updates"],
  background: ["Background", "How can we help you?"],
  modified: ["Modified"],
  created: ["Created"],
};

function cleanText(value) {
  return String(value ?? "").trim();
}

function normalizeToken(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/&amp;/g, "and")
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
    sitePath,
    listName,
  };
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

function parseEnquiriesConfig(env = process.env) {
  const listWebUrl = cleanText(env.SHAREPOINT_ENQUIRIES_LIST_WEB_URL || DEFAULT_ENQUIRIES_LIST_WEB_URL);
  return {
    ...parseSharePointListConfig(
      cleanText(env.SHAREPOINT_ENQUIRIES_SITE_URL || DEFAULT_ENQUIRIES_SITE_URL),
      cleanText(env.SHAREPOINT_ENQUIRIES_LIST_NAME || DEFAULT_ENQUIRIES_LIST_NAME),
      "Missing SHAREPOINT_ENQUIRIES_SITE_URL or SHAREPOINT_ENQUIRIES_LIST_NAME."
    ),
    listWebUrl,
    listPath: parseListPathFromWebUrl(listWebUrl),
  };
}

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

async function resolveSiteId(graphClient, hostName, sitePath) {
  const payload = await graphClient.fetchJson(`https://graph.microsoft.com/v1.0/sites/${hostName}:${sitePath}?$select=id`);
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
  const payload = await graphClient.fetchJson(`https://graph.microsoft.com/v1.0/sites/${siteId}/lists?${params.toString()}`);
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
    id: cleanText(list.id),
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

function createFieldMap(columns, definitions = ENQUIRY_FIELD_DEFINITIONS) {
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

function getFieldValue(fields, fieldMap, key) {
  const fieldName = fieldMap?.[key];
  if (!fieldName) {
    return "";
  }
  const value = fields?.[fieldName];
  if (value && typeof value === "object") {
    return cleanText(value.Title || value.title || value.Name || value.name || value.EMail || value.email || value.LookupValue);
  }
  return cleanText(value);
}

function normalizeStatus(value) {
  return cleanText(value).toLowerCase();
}

function isOnHoldStatus(status) {
  return normalizeStatus(status).includes("on hold");
}

function isCompletedStatus(status) {
  const normalized = normalizeStatus(status);
  return (
    normalized.includes("lost") ||
    normalized === "won" ||
    normalized === "didn't enquire" ||
    normalized === "didnt enquire" ||
    normalized === "not qualified"
  );
}

function isActiveReminderStatus(status) {
  return !isCompletedStatus(status) && !isOnHoldStatus(status);
}

function getValidDate(value) {
  const date = new Date(value);
  return value && !Number.isNaN(date.getTime()) ? date : null;
}

function isOverdue(modifiedValue, now = new Date(), overdueDays = OVERDUE_DAYS) {
  const modifiedDate = getValidDate(modifiedValue);
  if (!modifiedDate) {
    return false;
  }
  return modifiedDate.getTime() < now.getTime() - overdueDays * 24 * 60 * 60 * 1000;
}

function isFirstMonday(date = new Date()) {
  return date.getUTCDay() === 1 && date.getUTCDate() <= 7;
}

function createItemUrl(item, listWebUrl) {
  const directUrl = cleanText(item?.webUrl);
  if (directUrl) {
    return directUrl;
  }
  const id = cleanText(item?.id || item?.fields?.id || item?.fields?.Id);
  const baseUrl = cleanText(listWebUrl);
  if (!baseUrl || !id) {
    return "";
  }
  try {
    const url = new URL(baseUrl);
    const listPath = parseListPathFromWebUrl(baseUrl);
    url.pathname = `${listPath}/DispForm.aspx`;
    url.search = "";
    url.searchParams.set("ID", id);
    return url.toString();
  } catch {
    return "";
  }
}

function getDetailsExcerpt(fields, fieldMap) {
  const source = getFieldValue(fields, fieldMap, "updates") || getFieldValue(fields, fieldMap, "background");
  return cleanText(source).replace(/\s+/g, " ").slice(0, 240);
}

function normalizeEnquiryItem(item, fieldMap, listWebUrl) {
  const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
  const status = getFieldValue(fields, fieldMap, "status") || getFieldValue(fields, fieldMap, "currentStatus");
  const modified = cleanText(fields.Modified || item?.lastModifiedDateTime || getFieldValue(fields, fieldMap, "modified"));
  return {
    id: cleanText(item?.id || fields.id || fields.Id),
    title: getFieldValue(fields, fieldMap, "title") || cleanText(fields.Title) || `Enquiry ${cleanText(item?.id || fields.Id)}`,
    callerName: getFieldValue(fields, fieldMap, "callerName"),
    ownerName: getFieldValue(fields, fieldMap, "enquiryOwner") || "Unassigned",
    ownerEmail: "",
    status,
    followUp: getFieldValue(fields, fieldMap, "followUp"),
    modified,
    modifiedTime: getValidDate(modified)?.getTime() || 0,
    detailsExcerpt: getDetailsExcerpt(fields, fieldMap),
    webUrl: createItemUrl(item, listWebUrl),
  };
}

function classifyReminderItems(items, options = {}) {
  const now = options.now || new Date();
  const includeOnHold = options.includeOnHold !== false;
  const active = [];
  const onHold = [];

  for (const item of items || []) {
    if (!isOverdue(item.modified, now, options.overdueDays || OVERDUE_DAYS)) {
      continue;
    }
    if (isOnHoldStatus(item.status)) {
      if (includeOnHold) {
        onHold.push(item);
      }
      continue;
    }
    if (isActiveReminderStatus(item.status)) {
      active.push(item);
    }
  }

  const byOldestUpdate = (a, b) => (a.modifiedTime || 0) - (b.modifiedTime || 0) || a.title.localeCompare(b.title);
  return {
    active: active.sort(byOldestUpdate),
    onHold: onHold.sort(byOldestUpdate),
  };
}

async function fetchEnquiryItems(graphClient, config = parseEnquiriesConfig()) {
  const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
  const list = await resolveList(graphClient, siteId, config.listName, { listPath: config.listPath });
  const columns = await resolveListColumns(graphClient, siteId, list.id);
  const fieldMap = createFieldMap(columns);
  const selectFields = uniqueFieldNames(fieldMap);
  const expand = selectFields.length ? `fields($select=${selectFields.join(",")})` : "fields";
  const params = new URLSearchParams({
    $top: "200",
    $expand: expand,
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${list.id}/items?${params.toString()}`;
  const items = await graphClient.fetchAllPages(url);
  return {
    items: items.map((item) => normalizeEnquiryItem(item, fieldMap, list.webUrl || config.listWebUrl)),
    listUrl: list.webUrl || config.listWebUrl,
    fieldMap,
  };
}

function escapeHtml(value) {
  return cleanText(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function formatDateTime(value) {
  const date = getValidDate(value);
  if (!date) {
    return "-";
  }
  return new Intl.DateTimeFormat("en-GB", {
    timeZone: "Europe/London",
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function groupItemsByOwner(items) {
  const groups = new Map();
  for (const item of items || []) {
    const key = cleanText(item.ownerName) || "Unassigned";
    if (!groups.has(key)) {
      groups.set(key, []);
    }
    groups.get(key).push(item);
  }
  return Array.from(groups.entries()).sort(([a], [b]) => a.localeCompare(b));
}

function renderItemList(items) {
  if (!items.length) {
    return "<p>No matching enquiries.</p>";
  }

  return groupItemsByOwner(items)
    .map(([ownerName, ownerItems]) => {
      const rows = ownerItems
        .map((item) => {
          const title = escapeHtml(item.title);
          const link = item.webUrl ? `<a href="${escapeHtml(item.webUrl)}">${title}</a>` : title;
          const details = item.detailsExcerpt ? `<br><strong>Details:</strong> ${escapeHtml(item.detailsExcerpt)}` : "";
          return `<li>${link}<br><strong>Status:</strong> ${escapeHtml(item.status || "-")}<br><strong>Last updated:</strong> ${escapeHtml(
            formatDateTime(item.modified)
          )}<br><strong>Follow up:</strong> ${escapeHtml(item.followUp || "-")}${details}</li>`;
        })
        .join("");
      return `<h3>${escapeHtml(ownerName)}</h3><ul>${rows}</ul>`;
    })
    .join("");
}

function buildReminderEmail(classified, options = {}) {
  const now = options.now || new Date();
  const active = Array.isArray(classified?.active) ? classified.active : [];
  const onHold = Array.isArray(classified?.onHold) ? classified.onHold : [];
  const subject = `Enquiry check-in reminders: ${active.length} active, ${onHold.length} on hold`;
  const runLabel = formatDateTime(now.toISOString());
  const listLink = cleanText(options.listUrl) ? `<p><a href="${escapeHtml(options.listUrl)}">Open Enquiries Log</a></p>` : "";
  const onHoldNote = options.includeOnHold
    ? "On-hold reminders are included because this is the first Monday of the month."
    : "On-hold reminders are only included on the first Monday of each month.";
  const html = `<!doctype html>
<html>
  <body>
    <p>Hello,</p>
    <p>These enquiries may need a check-in. For v1, this reminder uses the SharePoint <strong>Modified</strong> date as the proxy for last spoken/contacted.</p>
    <p><strong>Run time:</strong> ${escapeHtml(runLabel)}<br><strong>Overdue threshold:</strong> ${OVERDUE_DAYS}+ days since last update<br>${escapeHtml(
      onHoldNote
    )}</p>
    ${listLink}
    <h2>Active enquiries overdue this week (${active.length})</h2>
    ${renderItemList(active)}
    <h2>On-hold enquiries overdue this month (${onHold.length})</h2>
    ${renderItemList(onHold)}
  </body>
</html>`;

  return {
    subject,
    html,
    text: html.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim(),
  };
}

function resolveRecipientOverride(env = process.env) {
  return parseEmailList(env.ENQUIRY_REMINDER_RECIPIENT_OVERRIDE);
}

async function runEnquiryReminderJob(options = {}) {
  const env = options.env || process.env;
  const now = options.now || new Date();
  const graphClient = options.graphClient || createGraphAppClient();
  const includeOnHold = isFirstMonday(now);
  const fromEmail = cleanText(env.ENQUIRY_REMINDER_FROM_EMAIL);
  const recipients = resolveRecipientOverride(env);

  const { items, listUrl, fieldMap } = options.items
    ? { items: options.items, listUrl: options.listUrl || "", fieldMap: options.fieldMap || {} }
    : await fetchEnquiryItems(graphClient, parseEnquiriesConfig(env));
  const classified = classifyReminderItems(items, { now, includeOnHold });
  const total = classified.active.length + classified.onHold.length;

  if (!total) {
    return {
      sent: false,
      reason: "NO_MATCHING_ENQUIRIES",
      activeCount: 0,
      onHoldCount: 0,
      includeOnHold,
      fieldMap,
    };
  }

  const email = buildReminderEmail(classified, { now, includeOnHold, listUrl });
  if (options.dryRun) {
    return {
      sent: false,
      dryRun: true,
      activeCount: classified.active.length,
      onHoldCount: classified.onHold.length,
      includeOnHold,
      recipients,
      email,
      fieldMap,
    };
  }

  if (!recipients.length) {
    const error = new Error("Missing ENQUIRY_REMINDER_RECIPIENT_OVERRIDE. Set it to chris@planwithcare.co.uk for trial mode.");
    error.code = "ENQUIRY_REMINDER_RECIPIENTS_MISSING";
    error.status = 500;
    throw error;
  }

  const delivery = await (options.sendMail || sendGraphMail)(graphClient, {
    fromEmail,
    to: recipients,
    subject: email.subject,
    html: email.html,
  });

  return {
    sent: true,
    activeCount: classified.active.length,
    onHoldCount: classified.onHold.length,
    includeOnHold,
    recipients: delivery.to,
    subject: delivery.subject,
    fieldMap,
  };
}

module.exports = {
  buildReminderEmail,
  classifyReminderItems,
  createFieldMap,
  fetchEnquiryItems,
  isActiveReminderStatus,
  isCompletedStatus,
  isFirstMonday,
  isOnHoldStatus,
  isOverdue,
  normalizeEnquiryItem,
  parseEmailList,
  parseEnquiriesConfig,
  resolveRecipientOverride,
  runEnquiryReminderJob,
};
