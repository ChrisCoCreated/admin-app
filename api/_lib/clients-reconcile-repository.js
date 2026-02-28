function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function normalizeToken(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function parseDateOfBirth(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }

  const datePart = raw.includes("T") ? raw.split("T")[0] : raw;
  const isoMatch = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(datePart);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    if (
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
    ) {
      return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  const slashMatch = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/.exec(datePart);
  if (slashMatch) {
    const day = Number(slashMatch[1]);
    const month = Number(slashMatch[2]);
    const yearRaw = Number(slashMatch[3]);
    const year = yearRaw < 100 ? 2000 + yearRaw : yearRaw;
    const date = new Date(Date.UTC(year, month - 1, day));
    if (
      date.getUTCFullYear() === year &&
      date.getUTCMonth() === month - 1 &&
      date.getUTCDate() === day
    ) {
      return `${String(year).padStart(4, "0")}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return [
      String(parsed.getUTCFullYear()).padStart(4, "0"),
      String(parsed.getUTCMonth() + 1).padStart(2, "0"),
      String(parsed.getUTCDate()).padStart(2, "0"),
    ].join("-");
  }

  return "";
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  const payload = text ? JSON.parse(text) : null;
  if (!response.ok) {
    const detail = payload?.error?.message || text || `HTTP ${response.status}`;
    const error = new Error(detail);
    error.status = response.status;
    throw error;
  }
  return payload;
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

function graphHeaders(token) {
  return {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
  };
}

function requireSharePointConfig() {
  const siteUrlValue = process.env.SHAREPOINT_SITE_URL;
  const clientsListName = process.env.SHAREPOINT_CLIENTS_LIST_NAME;

  if (!siteUrlValue || !clientsListName) {
    throw new Error("Missing SHAREPOINT_SITE_URL or SHAREPOINT_CLIENTS_LIST_NAME.");
  }

  const siteUrl = new URL(siteUrlValue);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    throw new Error("SHAREPOINT_SITE_URL must include a site path.");
  }

  return {
    hostName: siteUrl.hostname,
    sitePath,
    clientsListName,
  };
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
  let nextUrl =
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}` +
    `/columns?$select=name,displayName&$top=200`;
  while (nextUrl) {
    const payload = await fetchJson(nextUrl, { headers: graphHeaders(token) });
    columns.push(...(Array.isArray(payload?.value) ? payload.value : []));
    nextUrl = String(payload?.["@odata.nextLink"] || "");
  }
  return columns;
}

function buildFieldMap(columns) {
  const byNameToken = new Map();
  const byDisplayToken = new Map();
  for (const column of columns) {
    const name = String(column?.name || "").trim();
    const displayName = String(column?.displayName || "").trim();
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
    name: findField(["Title", "Client", "ClientName"], "Title"),
    oneTouchId: findField(["OnetouchID", "OneTouchID", "OneTouchClientId", "OneTouch_Id"]),
    dateOfBirth: findField(["DateOfBirth", "DOB", "BirthDate"]),
    postcode: findField(["PostCode", "Postcode", "PostalCode", "ZipCode"]),
    email: findField(["Email", "EmailAddress", "ClientEmail"]),
    phone: findField(["Phone", "Mobile", "PhoneNumber"]),
    status: findField(["Status", "ClientStatus", "State", "CurrentStatus"]),
    address: findField(["Address", "AddressLine1", "Address1", "StreetAddress", "Line1"]),
    town: findField(["Town", "City", "Suburb"]),
    county: findField(["County", "Region", "State"]),
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
  const { hostName, sitePath, clientsListName } = requireSharePointConfig();
  const key = `${hostName}|${sitePath}|${clientsListName}`;
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
  const listId = await resolveListId(token, siteId, clientsListName);
  const columns = await resolveColumns(token, siteId, listId);
  const fieldMap = buildFieldMap(columns);

  listConfigCache.key = key;
  listConfigCache.siteId = siteId;
  listConfigCache.listId = listId;
  listConfigCache.fieldMap = fieldMap;
  listConfigCache.expiresAt = Date.now() + 5 * 60 * 1000;

  return { siteId, listId, fieldMap };
}

function asString(value) {
  return String(value || "").trim();
}

function normalizeSharePointClient(item, fieldMap) {
  const fields = item?.fields || {};
  const nameField = fieldMap.name;
  const oneTouchIdField = fieldMap.oneTouchId;
  const dobField = fieldMap.dateOfBirth;
  const row = {
    itemId: asString(item?.id),
    id: asString(fields.ID || item?.id),
    name: asString(fields[nameField] || fields.Title || fields.Client || fields.ClientName),
    oneTouchId: asString(oneTouchIdField ? fields[oneTouchIdField] : ""),
    dateOfBirth: parseDateOfBirth(dobField ? fields[dobField] : ""),
    postcode: asString(fieldMap.postcode ? fields[fieldMap.postcode] : ""),
    email: asString(fieldMap.email ? fields[fieldMap.email] : ""),
    phone: asString(fieldMap.phone ? fields[fieldMap.phone] : ""),
    status: asString(fieldMap.status ? fields[fieldMap.status] : "").toLowerCase(),
    address: asString(fieldMap.address ? fields[fieldMap.address] : ""),
    town: asString(fieldMap.town ? fields[fieldMap.town] : ""),
    county: asString(fieldMap.county ? fields[fieldMap.county] : ""),
  };
  return row;
}

async function listSharePointClientsWithItemIds() {
  const token = await getGraphAccessToken();
  const { siteId, listId, fieldMap } = await resolveListConfig(token);

  const items = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;
  while (nextUrl) {
    const payload = await fetchJson(nextUrl, { headers: graphHeaders(token) });
    items.push(...(Array.isArray(payload?.value) ? payload.value : []));
    nextUrl = String(payload?.["@odata.nextLink"] || "");
  }

  return {
    clients: items.map((item) => normalizeSharePointClient(item, fieldMap)),
    fieldMap,
    siteId,
    listId,
    token,
  };
}

function toSharePointFields(logicalFields, fieldMap) {
  const mapped = {};
  function setIfPresent(logicalKey, value) {
    const targetField = fieldMap[logicalKey];
    if (!targetField) {
      return;
    }
    if (value === undefined || value === null) {
      return;
    }
    mapped[targetField] = String(value).trim();
  }

  setIfPresent("name", logicalFields.name);
  setIfPresent("oneTouchId", logicalFields.oneTouchId);
  setIfPresent("dateOfBirth", logicalFields.dateOfBirth);
  setIfPresent("postcode", logicalFields.postcode);
  setIfPresent("email", logicalFields.email);
  setIfPresent("phone", logicalFields.phone);
  setIfPresent("status", logicalFields.status);
  setIfPresent("address", logicalFields.address);
  setIfPresent("town", logicalFields.town);
  setIfPresent("county", logicalFields.county);

  return mapped;
}

async function patchSharePointClient(itemId, logicalFields) {
  const token = await getGraphAccessToken();
  const { siteId, listId, fieldMap } = await resolveListConfig(token);
  const fieldsPayload = toSharePointFields(logicalFields, fieldMap);
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}/fields`;
  await fetchJson(url, {
    method: "PATCH",
    headers: {
      ...graphHeaders(token),
      "Content-Type": "application/json",
    },
    body: JSON.stringify(fieldsPayload),
  });
}

async function createSharePointClient(logicalFields) {
  const token = await getGraphAccessToken();
  const { siteId, listId, fieldMap } = await resolveListConfig(token);
  const fieldsPayload = toSharePointFields(logicalFields, fieldMap);
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
  const created = await fetchJson(url, {
    method: "POST",
    headers: {
      ...graphHeaders(token),
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields: fieldsPayload }),
  });

  return {
    itemId: asString(created?.id),
  };
}

module.exports = {
  createSharePointClient,
  listSharePointClientsWithItemIds,
  patchSharePointClient,
};
