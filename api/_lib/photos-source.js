function normalizeToken(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
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
  if (!siteUrlValue) {
    throw new Error("Missing SHAREPOINT_SITE_URL.");
  }

  const siteUrl = new URL(siteUrlValue);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    throw new Error("SHAREPOINT_SITE_URL must include a site path, e.g. /sites/SupportTeam.");
  }

  return {
    hostName: siteUrl.hostname,
    sitePath,
    photosListName: process.env.SHAREPOINT_MARKETING_PHOTOS_LIST_NAME || "Photos and Lists",
  };
}

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function graphHeaders(token) {
  return {
    Authorization: `Bearer ${token}`,
    Accept: "application/json",
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

function toAbsoluteUrl(value, hostName) {
  const input = String(value || "").trim();
  if (!input) {
    return "";
  }
  if (/^https?:\/\//i.test(input)) {
    return input;
  }
  if (input.startsWith("/")) {
    return `https://${hostName}${input}`;
  }
  return "";
}

function parseMaybeJson(value) {
  const raw = String(value || "").trim();
  if (!raw || (!raw.startsWith("{") && !raw.startsWith("["))) {
    return null;
  }
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function asTruthy(value) {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return value > 0;
  }

  const normalized = String(value || "")
    .trim()
    .toLowerCase();
  if (!normalized) {
    return false;
  }

  return ["1", "true", "yes", "y", "agreed", "approved", "consented"].includes(normalized);
}

function pickClientConsent(fields) {
  const entries = Object.entries(fields || {});
  if (!entries.length) {
    return false;
  }

  const tokenMap = new Map();
  for (const [key, value] of entries) {
    const token = normalizeToken(key);
    if (token && !tokenMap.has(token)) {
      tokenMap.set(token, value);
    }
  }

  const directCandidates = [
    "clientconsent",
    "clientconsentgiven",
    "clientconsentobtained",
    "consentfromclient",
    "consentclient",
    "photoconsent",
    "consent",
  ];
  for (const token of directCandidates) {
    if (tokenMap.has(token)) {
      return asTruthy(tokenMap.get(token));
    }
  }

  for (const [key, value] of entries) {
    const token = normalizeToken(key);
    if (token.includes("consent") && token.includes("client")) {
      return asTruthy(value);
    }
  }

  return false;
}

function pickImageUrl(fields, hostName) {
  const values = Object.values(fields || {});
  for (const value of values) {
    if (!value) {
      continue;
    }

    if (typeof value === "object" && !Array.isArray(value)) {
      const objectUrl =
        toAbsoluteUrl(value.url, hostName) ||
        toAbsoluteUrl(value.nativeUrl, hostName) ||
        toAbsoluteUrl(value.serverRelativeUrl, hostName) ||
        toAbsoluteUrl(value.serverUrl, hostName);
      if (objectUrl) {
        return objectUrl;
      }
    }

    if (typeof value === "string") {
      const parsed = parseMaybeJson(value);
      if (parsed && typeof parsed === "object") {
        const parsedUrl =
          toAbsoluteUrl(parsed.url, hostName) ||
          toAbsoluteUrl(parsed.nativeUrl, hostName) ||
          toAbsoluteUrl(parsed.serverRelativeUrl, hostName) ||
          toAbsoluteUrl(parsed.serverUrl, hostName);
        if (parsedUrl) {
          return parsedUrl;
        }
      }

      const asUrl = toAbsoluteUrl(value, hostName);
      if (asUrl && /(\.png|\.jpe?g|\.gif|\.webp|\.bmp|\/_layouts\/15\/getpreview)/i.test(asUrl)) {
        return asUrl;
      }
    }
  }

  return "";
}

function mapGraphItemToPhoto(item, hostName) {
  const fields = item?.fields || {};
  const consent = pickClientConsent(fields);
  if (!consent) {
    return null;
  }

  const imageUrl = pickImageUrl(fields, hostName);
  if (!imageUrl) {
    return null;
  }

  const id = String(fields.ID || item?.id || "").trim();
  if (!id) {
    return null;
  }

  const title = String(fields.Title || fields.Client || fields.ClientName || `Photo ${id}`).trim();
  const client = String(fields.Client || fields.ClientName || title).trim();

  return {
    id,
    title,
    client,
    imageUrl,
  };
}

async function readMarketingPhotos() {
  const token = await getGraphAccessToken();
  const { hostName, sitePath, photosListName } = requireSharePointConfig();
  const siteId = await resolveSiteId(token, hostName, sitePath);
  const listId = await resolveListId(token, siteId, photosListName);

  const photos = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;

  while (nextUrl) {
    const data = await fetchJson(nextUrl, { headers: graphHeaders(token) });
    const items = Array.isArray(data?.value) ? data.value : [];

    for (const item of items) {
      const photo = mapGraphItemToPhoto(item, hostName);
      if (photo) {
        photos.push(photo);
      }
    }

    nextUrl = data?.["@odata.nextLink"] || "";
  }

  photos.sort((a, b) => a.client.localeCompare(b.client, undefined, { sensitivity: "base" }));
  return photos;
}

module.exports = {
  readMarketingPhotos,
};
