const fs = require("fs/promises");
const path = require("path");

function normalizeClient(client) {
  return {
    id: String(client.id || "").trim(),
    name: String(client.name || "").trim(),
    location: String(client.location || "").trim(),
    address: String(client.address || "").trim(),
    town: String(client.town || "").trim(),
    county: String(client.county || "").trim(),
    postcode: String(client.postcode || "").trim(),
    email: String(client.email || "").trim(),
  };
}

async function readLocalClients() {
  const filePath = process.env.CLIENTS_DATA_FILE
    ? path.resolve(process.cwd(), process.env.CLIENTS_DATA_FILE)
    : path.join(process.cwd(), "data", "clients.json");

  const raw = await fs.readFile(filePath, "utf8");
  const parsed = JSON.parse(raw);

  if (!Array.isArray(parsed)) {
    throw new Error("Client data must be an array.");
  }

  return parsed.map(normalizeClient).filter((client) => client.id);
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  const text = await response.text();
  const data = text ? JSON.parse(text) : null;

  if (!response.ok) {
    const detail = data?.error?.message || text || `HTTP ${response.status}`;
    throw new Error(detail);
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

function requireSharePointConfig() {
  const siteUrlValue = process.env.SHAREPOINT_SITE_URL;
  const clientsListName = process.env.SHAREPOINT_CLIENTS_LIST_NAME;

  if (!siteUrlValue || !clientsListName) {
    throw new Error("Missing SHAREPOINT_SITE_URL or SHAREPOINT_CLIENTS_LIST_NAME.");
  }

  const siteUrl = new URL(siteUrlValue);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    throw new Error("SHAREPOINT_SITE_URL must include a site path, e.g. /sites/SupportTeam.");
  }

  return {
    hostName: siteUrl.hostname,
    sitePath,
    clientsListName,
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

async function readListColumns(token, siteId, listId) {
  const columns = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns?$select=name,displayName&$top=200`;

  while (nextUrl) {
    const data = await fetchJson(nextUrl, { headers: graphHeaders(token) });
    const page = Array.isArray(data?.value) ? data.value : [];
    for (const item of page) {
      columns.push(item);
    }
    nextUrl = data?.["@odata.nextLink"] || "";
  }

  return columns;
}

function normalizeToken(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function pickFieldName(columns, candidates) {
  const columnByToken = new Map();
  for (const column of columns) {
    const internalName = String(column?.name || "").trim();
    if (!internalName) {
      continue;
    }

    const nameToken = normalizeToken(internalName);
    if (nameToken && !columnByToken.has(nameToken)) {
      columnByToken.set(nameToken, internalName);
    }

    const displayToken = normalizeToken(column?.displayName);
    if (displayToken && !columnByToken.has(displayToken)) {
      columnByToken.set(displayToken, internalName);
    }
  }

  for (const candidate of candidates) {
    const token = normalizeToken(candidate);
    if (columnByToken.has(token)) {
      return columnByToken.get(token);
    }
  }

  return "";
}

function createError(message, statusCode) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function mapGraphItemToClient(item) {
  const fields = item?.fields || {};
  const graphId = fields.ID || item?.id;
  const name = fields.Title || fields.Client || fields.ClientName || "";
  const entries = Object.entries(fields)
    .map(([key, value]) => [String(key || ""), String(value || "").trim()])
    .filter(([, value]) => Boolean(value));

  const byToken = new Map();
  for (const [key, value] of entries) {
    const token = normalizeToken(key);
    if (token && !byToken.has(token)) {
      byToken.set(token, value);
    }
  }

  function pickFieldValue(candidates) {
    for (const candidate of candidates) {
      const match = byToken.get(normalizeToken(candidate));
      if (match) {
        return match;
      }
    }
    return "";
  }

  function inferAddressLikeValue() {
    const postcodePattern = /[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}$/i;
    for (const [key, value] of entries) {
      const keyToken = normalizeToken(key);
      if (!value || value.length < 10) {
        continue;
      }
      if (!postcodePattern.test(value)) {
        continue;
      }
      if (
        keyToken.includes("address") ||
        keyToken.includes("location") ||
        keyToken.includes("street")
      ) {
        return value;
      }
    }

    for (const [, value] of entries) {
      if (postcodePattern.test(value) && value.length >= 10) {
        return value;
      }
    }

    return "";
  }

  const address =
    pickFieldValue([
      "Address",
      "AddressLine1",
      "Address1",
      "StreetAddress",
      "Line1",
      "Address_x0020_Line_x0020_1",
    ]) || inferAddressLikeValue();
  const town = pickFieldValue(["Town", "City", "Suburb", "Town_x0020_City"]);
  const county = pickFieldValue(["County", "Region", "State"]);
  const postcode = pickFieldValue([
    "PostCode",
    "Postcode",
    "PostalCode",
    "ZipCode",
    "Post_x0020_Code",
  ]);
  const location =
    pickFieldValue(["Location", "Locality"]) || town || address || "";
  const email = fields.Email || fields.EmailAddress || fields.ClientEmail || "";

  return normalizeClient({
    id: graphId,
    name,
    location,
    address,
    town,
    county,
    postcode,
    email,
  });
}

async function loadClientsFromGraph() {
  const token = await getGraphAccessToken();
  const { hostName, sitePath, clientsListName } = requireSharePointConfig();
  const siteId = await resolveSiteId(token, hostName, sitePath);
  const listId = await resolveListId(token, siteId, clientsListName);

  const clients = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;

  while (nextUrl) {
    const data = await fetchJson(nextUrl, { headers: graphHeaders(token) });
    const items = Array.isArray(data?.value) ? data.value : [];

    for (const item of items) {
      const client = mapGraphItemToClient(item);
      if (client.id && client.name) {
        clients.push(client);
      }
    }

    nextUrl = data?.["@odata.nextLink"] || "";
  }

  clients.sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
  return clients;
}

async function readClients() {
  const useFallback = process.env.USE_LOCAL_CLIENTS_FALLBACK === "1";

  if (useFallback) {
    return { clients: await readLocalClients(), source: "local-fallback" };
  }

  try {
    return { clients: await loadClientsFromGraph(), source: "graph" };
  } catch (error) {
    if (process.env.ALLOW_LOCAL_CLIENTS_ON_GRAPH_ERROR === "1") {
      const clients = await readLocalClients();
      return {
        clients,
        source: "local-on-error",
        graphError: error && error.message ? error.message : String(error),
      };
    }

    throw error;
  }
}

async function persistClientLocationFields(clientId, locationFields) {
  if (process.env.USE_LOCAL_CLIENTS_FALLBACK === "1") {
    throw createError("Cannot persist changes while local fallback mode is enabled.", 409);
  }

  const token = await getGraphAccessToken();
  const { hostName, sitePath, clientsListName } = requireSharePointConfig();
  const siteId = await resolveSiteId(token, hostName, sitePath);
  const listId = await resolveListId(token, siteId, clientsListName);
  const columns = await readListColumns(token, siteId, listId);

  const fieldMap = {
    address: pickFieldName(columns, [
      "Address",
      "AddressLine1",
      "StreetAddress",
      "Address1",
      "Address Line 1",
    ]),
    town: pickFieldName(columns, ["Town", "City", "Suburb", "Town / City"]),
    county: pickFieldName(columns, ["County", "Region", "State"]),
    postcode: pickFieldName(columns, ["PostCode", "Postcode", "PostalCode", "ZipCode", "ZIP Code"]),
    location: pickFieldName(columns, ["Location"]),
  };

  const updates = {};
  const address = String(locationFields?.address || "").trim();
  const town = String(locationFields?.town || "").trim();
  const county = String(locationFields?.county || "").trim();
  const postcode = String(locationFields?.postcode || "").trim();

  if (fieldMap.address && address) {
    updates[fieldMap.address] = address;
  }
  if (fieldMap.town && town) {
    updates[fieldMap.town] = town;
  }
  if (fieldMap.county && county) {
    updates[fieldMap.county] = county;
  }
  if (fieldMap.postcode && postcode) {
    updates[fieldMap.postcode] = postcode;
  }
  if (fieldMap.location) {
    const combinedLocation = [town, county, postcode].filter(Boolean).join(", ");
    if (combinedLocation) {
      updates[fieldMap.location] = combinedLocation;
    }
  }

  if (!Object.keys(updates).length) {
    throw createError("No matching writable location columns were found in SharePoint.", 400);
  }

  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(
    String(clientId || "").trim()
  )}/fields`;

  await fetchJson(url, {
    method: "PATCH",
    headers: {
      ...graphHeaders(token),
      "Content-Type": "application/json",
    },
    body: JSON.stringify(updates),
  });

  return {
    updatedFields: Object.keys(updates),
  };
}

module.exports = {
  readClients,
  persistClientLocationFields,
};
