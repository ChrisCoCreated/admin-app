const { createGraphAppClient } = require("../graph-app-client");

const listCache = {
  key: "",
  siteId: "",
  listId: "",
  expiresAt: 0,
};

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeEmail(value) {
  return normalizeText(value).toLowerCase();
}

function quoteODataString(value) {
  return `'${String(value || "").replace(/'/g, "''")}'`;
}

function parseTaskSetConfig() {
  const siteUrlRaw = normalizeText(process.env.SHAREPOINT_TASK_SETS_SITE_URL || process.env.SHAREPOINT_SITE_URL);
  const listName = normalizeText(process.env.SHAREPOINT_TASK_SETS_LIST_NAME || "Actions for Task Sets");

  if (!siteUrlRaw || !listName) {
    throw new Error("Missing SHAREPOINT_TASK_SETS_SITE_URL/SHAREPOINT_SITE_URL or SHAREPOINT_TASK_SETS_LIST_NAME.");
  }

  const siteUrl = new URL(siteUrlRaw);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    throw new Error("Task set SharePoint site URL must include a site path.");
  }

  return {
    hostName: siteUrl.hostname,
    sitePath,
    listName,
  };
}

async function resolveSiteAndListIds(graphClient, config) {
  const cacheKey = `${config.hostName}|${config.sitePath}|${config.listName}`;
  if (
    listCache.key === cacheKey &&
    listCache.siteId &&
    listCache.listId &&
    listCache.expiresAt > Date.now()
  ) {
    return {
      siteId: listCache.siteId,
      listId: listCache.listId,
    };
  }

  const sitePayload = await graphClient.fetchJson(
    `https://graph.microsoft.com/v1.0/sites/${config.hostName}:${config.sitePath}?$select=id`
  );
  const siteId = normalizeText(sitePayload?.id);
  if (!siteId) {
    throw new Error("Could not resolve SharePoint site id for task sets.");
  }

  const params = new URLSearchParams({
    $select: "id,displayName",
    $filter: `displayName eq ${quoteODataString(config.listName)}`,
    $top: "1",
  });
  const listPayload = await graphClient.fetchJson(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?${params.toString()}`
  );
  const listId = normalizeText(listPayload?.value?.[0]?.id);
  if (!listId) {
    throw new Error(`Could not find task set list '${config.listName}'.`);
  }

  listCache.key = cacheKey;
  listCache.siteId = siteId;
  listCache.listId = listId;
  listCache.expiresAt = Date.now() + 5 * 60 * 1000;

  return { siteId, listId };
}

function extractEmailCandidate(value) {
  if (Array.isArray(value)) {
    for (const entry of value) {
      const match = extractEmailCandidate(entry);
      if (match) {
        return match;
      }
    }
    return "";
  }

  if (value && typeof value === "object") {
    return (
      extractEmailCandidate(value.email) ||
      extractEmailCandidate(value.Email) ||
      extractEmailCandidate(value.mail) ||
      extractEmailCandidate(value.Mail) ||
      extractEmailCandidate(value.userPrincipalName) ||
      extractEmailCandidate(value.UserPrincipalName) ||
      extractEmailCandidate(value.lookupValue) ||
      extractEmailCandidate(value.LookupValue) ||
      ""
    );
  }

  const raw = normalizeText(value);
  if (!raw) {
    return "";
  }

  const match = raw.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return match ? normalizeEmail(match[0]) : "";
}

function extractText(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => extractText(entry)).find(Boolean) || "";
  }
  if (value && typeof value === "object") {
    return normalizeText(
      value.lookupValue ||
        value.LookupValue ||
        value.displayName ||
        value.DisplayName ||
        value.title ||
        value.Title ||
        value.value ||
        value.Value ||
        value.text ||
        value.Text ||
        ""
    );
  }
  return normalizeText(value);
}

function toNumberOrDefault(value, fallbackValue) {
  if (value === "" || value === null || value === undefined) {
    return fallbackValue;
  }
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : fallbackValue;
}

function mapTaskSetRow(item) {
  const fields = item?.fields || {};
  const title = normalizeText(fields.Title || item?.name);
  const taskSet = normalizeText(fields.TaskSet);
  const area = normalizeText(fields.Area);
  const description = extractText(fields.Description);
  const responsiblePerson =
    extractEmailCandidate(fields.ResponsiblePerson) ||
    extractEmailCandidate(fields.ResponsiblePerson0) ||
    extractEmailCandidate(fields.Responsible_x0020_Person) ||
    normalizeEmail(extractText(fields.ResponsiblePerson));
  const dueDateDelay = toNumberOrDefault(fields.DueDateDelay, -1);

  return {
    itemId: normalizeText(item?.id),
    title,
    taskSet,
    area,
    description,
    responsiblePerson,
    dueDateDelay,
  };
}

async function listTaskSetTemplates(filters = {}) {
  const graphClient = createGraphAppClient();
  const config = parseTaskSetConfig();
  const { siteId, listId } = await resolveSiteAndListIds(graphClient, config);
  const params = new URLSearchParams({
    $expand: "fields($select=Title,TaskSet,Description,ResponsiblePerson,DueDateDelay,Area)",
    $top: "200",
  });

  const rows = await graphClient.fetchAllPages(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${params.toString()}`
  );

  const taskSetFilter = normalizeText(filters.taskSet).toLowerCase();
  const areaFilter = normalizeText(filters.area).toLowerCase();

  const templates = rows
    .map((item) => mapTaskSetRow(item))
    .filter((row) => row.title && row.taskSet)
    .filter((row) => (!taskSetFilter ? true : row.taskSet.toLowerCase() === taskSetFilter))
    .filter((row) => (!areaFilter ? true : row.area.toLowerCase() === areaFilter));

  return {
    templates,
    meta: {
      siteId,
      listId,
      listName: config.listName,
      count: templates.length,
    },
  };
}

module.exports = {
  listTaskSetTemplates,
};
