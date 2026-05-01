const DEFAULT_COLLEAGUES_SITE_URL = "https://planwithcare.sharepoint.com/sites/SupportTeam";
const DEFAULT_COLLEAGUES_LIST_NAME = "Colleagues";

function normalizeText(value) {
  return String(value || "").trim();
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  const data = text ? JSON.parse(text) : null;

  if (!response.ok) {
    const detail = data?.error?.message || text || `HTTP ${response.status}`;
    const error = new Error(detail);
    error.status = response.status;
    throw error;
  }

  return data;
}

async function getGraphAccessToken() {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_API_CLIENT_ID;
  const clientSecret = process.env.AZURE_API_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing AZURE_TENANT_ID, AZURE_API_CLIENT_ID, or AZURE_API_CLIENT_SECRET.");
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type: "client_credentials",
    client_id: clientId,
    client_secret: clientSecret,
    scope: "https://graph.microsoft.com/.default",
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const payload = await response.json();
  if (!response.ok || !payload?.access_token) {
    const errorText = payload?.error_description || payload?.error || "Could not get Graph token.";
    throw new Error(errorText);
  }

  return payload.access_token;
}

function requireColleaguesConfig() {
  const siteUrlValue = process.env.SHAREPOINT_COLLEAGUES_SITE_URL || DEFAULT_COLLEAGUES_SITE_URL;
  const listName = process.env.SHAREPOINT_COLLEAGUES_LIST_NAME || DEFAULT_COLLEAGUES_LIST_NAME;

  if (!siteUrlValue || !listName) {
    throw new Error("Missing SHAREPOINT_COLLEAGUES_SITE_URL or SHAREPOINT_COLLEAGUES_LIST_NAME.");
  }

  const siteUrl = new URL(siteUrlValue);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    throw new Error("SHAREPOINT_COLLEAGUES_SITE_URL must include a site path, e.g. /sites/SupportTeam.");
  }

  return {
    hostName: siteUrl.hostname,
    sitePath,
    listName: normalizeText(listName),
  };
}

function graphHeaders(token) {
  return {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
  };
}

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

async function resolveSiteId(token, hostName, sitePath) {
  const url = `https://graph.microsoft.com/v1.0/sites/${hostName}:${sitePath}?$select=id`;
  const data = await fetchJson(url, { headers: graphHeaders(token) });
  if (!data?.id) {
    throw new Error("Could not resolve SharePoint site id.");
  }
  return data.id;
}

async function resolveListId(token, siteId, listName) {
  const params = new URLSearchParams({
    $select: "id,displayName",
    $filter: `displayName eq ${quoteODataString(listName)}`,
    $top: "1",
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?${params.toString()}`;
  const data = await fetchJson(url, { headers: graphHeaders(token) });
  const list = Array.isArray(data?.value) ? data.value[0] : null;
  if (!list?.id) {
    throw new Error(`Could not find SharePoint list '${listName}'.`);
  }
  return list.id;
}

async function resolveColumns(token, siteId, listId) {
  const columns = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns?$select=name,displayName&$top=200`;

  while (nextUrl) {
    const payload = await fetchJson(nextUrl, { headers: graphHeaders(token) });
    columns.push(...(Array.isArray(payload?.value) ? payload.value : []));
    nextUrl = String(payload?.["@odata.nextLink"] || "");
  }

  return columns;
}

function normalizeToken(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function buildFieldMap(columns) {
  const byNameToken = new Map();
  const byDisplayToken = new Map();

  for (const column of columns) {
    const name = normalizeText(column?.name);
    const displayName = normalizeText(column?.displayName);
    if (!name) {
      continue;
    }
    const nameToken = normalizeToken(name);
    const displayToken = normalizeToken(displayName);
    if (nameToken && !byNameToken.has(nameToken)) {
      byNameToken.set(nameToken, name);
    }
    if (displayToken && !byDisplayToken.has(displayToken)) {
      byDisplayToken.set(displayToken, name);
    }
  }

  function findField(candidates, fallback = "") {
    for (const candidate of candidates) {
      const token = normalizeToken(candidate);
      const found = byDisplayToken.get(token) || byNameToken.get(token);
      if (found) {
        return found;
      }
    }
    return fallback;
  }

  return {
    title: findField(["Title", "Name"], "Title"),
    oneTouchId: findField(["OnetouchID", "OneTouchID", "OneTouch Id"], "OnetouchID"),
    archived: findField(["Archived"], "Archived"),
  };
}

const listConfigCache = {
  key: "",
  siteId: "",
  listId: "",
  fieldMap: null,
  expiresAt: 0,
};

async function resolveListConfig(token) {
  const { hostName, sitePath, listName } = requireColleaguesConfig();
  const key = `${hostName}|${sitePath}|${listName}`;

  if (
    listConfigCache.key === key &&
    listConfigCache.siteId &&
    listConfigCache.listId &&
    listConfigCache.fieldMap &&
    listConfigCache.expiresAt > Date.now()
  ) {
    return {
      siteId: listConfigCache.siteId,
      listId: listConfigCache.listId,
      fieldMap: listConfigCache.fieldMap,
    };
  }

  const siteId = await resolveSiteId(token, hostName, sitePath);
  const listId = await resolveListId(token, siteId, listName);
  const columns = await resolveColumns(token, siteId, listId);
  const fieldMap = buildFieldMap(columns);

  listConfigCache.key = key;
  listConfigCache.siteId = siteId;
  listConfigCache.listId = listId;
  listConfigCache.fieldMap = fieldMap;
  listConfigCache.expiresAt = Date.now() + 5 * 60 * 1000;

  return { siteId, listId, fieldMap };
}

function coerceBoolean(value) {
  if (value === true || value === false) {
    return value;
  }
  const normalized = normalizeText(value).toLowerCase();
  if (!normalized) {
    return false;
  }
  return normalized === "true" || normalized === "1" || normalized === "yes";
}

function mapGraphItemToColleague(item, fieldMap) {
  const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
  return {
    itemId: normalizeText(item?.id),
    name: normalizeText(fields[fieldMap.title] || fields.Title),
    oneTouchId: normalizeText(fields[fieldMap.oneTouchId]),
    archived: coerceBoolean(fields[fieldMap.archived]),
  };
}

async function listSharePointColleagues() {
  const token = await getGraphAccessToken();
  const { siteId, listId, fieldMap } = await resolveListConfig(token);
  const colleagues = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;

  while (nextUrl) {
    const data = await fetchJson(nextUrl, { headers: graphHeaders(token) });
    const items = Array.isArray(data?.value) ? data.value : [];
    for (const item of items) {
      const colleague = mapGraphItemToColleague(item, fieldMap);
      if (colleague.oneTouchId) {
        colleagues.push(colleague);
      }
    }
    nextUrl = String(data?.["@odata.nextLink"] || "");
  }

  return colleagues;
}

async function createSharePointColleague({ name = "", oneTouchId = "", archived = false } = {}) {
  const normalizedOneTouchId = normalizeText(oneTouchId);
  if (!normalizedOneTouchId) {
    throw new Error("A OneTouch ID is required to create a colleague.");
  }

  const token = await getGraphAccessToken();
  const { siteId, listId, fieldMap } = await resolveListConfig(token);
  if (!fieldMap.oneTouchId) {
    throw new Error("Could not find the Colleagues OnetouchID field.");
  }

  const fields = {
    [fieldMap.title]: normalizeText(name) || `OneTouch ${normalizedOneTouchId}`,
    [fieldMap.oneTouchId]: normalizedOneTouchId,
  };

  if (fieldMap.archived) {
    fields[fieldMap.archived] = Boolean(archived);
  }

  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
  const created = await fetchJson(url, {
    method: "POST",
    headers: {
      ...graphHeaders(token),
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields }),
  });

  return {
    itemId: normalizeText(created?.id),
  };
}

module.exports = {
  createSharePointColleague,
  listSharePointColleagues,
};
