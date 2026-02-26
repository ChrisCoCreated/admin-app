function normalizeToken(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

const MARKETING_PHOTOS_DEBUG = process.env.MARKETING_PHOTOS_DEBUG === "1";

function logMarketingDebug(message, details) {
  if (!MARKETING_PHOTOS_DEBUG) {
    return;
  }
  if (details !== undefined) {
    console.log(`[marketing-photos] ${message}`, details);
    return;
  }
  console.log(`[marketing-photos] ${message}`);
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

async function resolveList(token, siteId, listName) {
  const params = new URLSearchParams({
    $select: "id,displayName,webUrl",
    $filter: `displayName eq ${quoteODataString(listName)}`,
    $top: "1",
  });

  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?${params.toString()}`;
  const data = await fetchJson(url, { headers: graphHeaders(token) });
  const list = Array.isArray(data?.value) ? data.value[0] : null;

  if (!list?.id) {
    throw new Error(`Could not find SharePoint list '${listName}'.`);
  }

  return {
    id: list.id,
    webUrl: String(list.webUrl || "").trim(),
  };
}

async function resolveListColumns(token, siteId, listId) {
  const columns = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns?$select=name,displayName&$top=200`;

  while (nextUrl) {
    const data = await fetchJson(nextUrl, { headers: graphHeaders(token) });
    const values = Array.isArray(data?.value) ? data.value : [];
    columns.push(...values);
    nextUrl = data?.["@odata.nextLink"] || "";
  }

  return columns;
}

function resolveConsentSelectFields(columns) {
  const result = new Set();
  for (const column of columns) {
    const name = String(column?.name || "").trim();
    const displayName = String(column?.displayName || "").trim();
    if (!name) {
      continue;
    }

    const nameToken = normalizeToken(name);
    const displayToken = normalizeToken(displayName);
    const isConsentMarketing =
      (nameToken.includes("consent") && nameToken.includes("marketing")) ||
      (displayToken.includes("consent") && displayToken.includes("marketing"));
    if (isConsentMarketing) {
      result.add(name);
    }
  }
  return Array.from(result);
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

function extractFileName(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }
  const cleaned = text.replace(/^['"]|['"]$/g, "");
  const withoutQuery = cleaned.split("?")[0];
  const match = /([^/\\]+)$/.exec(withoutQuery);
  return match ? match[1] : "";
}

function extractFileNameWithExtension(value) {
  const name = extractFileName(value);
  if (!name) {
    return "";
  }
  return /\.(?:png|jpe?g|gif|webp|bmp|mp4|mov|webm)$/i.test(name) ? name : "";
}

function extractKnownFileName(value) {
  return extractFileNameWithExtension(value) || extractFileName(value);
}

function hasKnownMediaExtension(name) {
  return /\.(?:png|jpe?g|gif|webp|bmp|mp4|mov|webm)$/i.test(String(name || "").trim());
}

function inferMediaType(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text) {
    return "image";
  }
  if (text.includes(".mp4") || text.includes(".mov") || text.includes(".webm")) {
    return "video";
  }
  return "image";
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

function extractLookupText(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (Array.isArray(value)) {
    const firstText = value.map((entry) => extractLookupText(entry)).find((entry) => Boolean(entry));
    return firstText || "";
  }
  if (typeof value === "string") {
    const parsed = parseMaybeJson(value);
    if (parsed && typeof parsed === "object") {
      return String(
        parsed.lookupValue ||
          parsed.LookupValue ||
          parsed.value ||
          parsed.Value ||
          parsed.label ||
          parsed.Label ||
          parsed.text ||
          parsed.Text ||
          ""
      ).trim();
    }
    return value.trim();
  }
  if (typeof value === "object") {
    if (Array.isArray(value.results)) {
      const firstResult = value.results
        .map((entry) => extractLookupText(entry))
        .find((entry) => Boolean(entry));
      if (firstResult) {
        return firstResult;
      }
    }
    return String(
      value.lookupValue ||
        value.LookupValue ||
        value.value ||
        value.Value ||
        value.label ||
        value.Label ||
        value.text ||
        value.Text ||
        ""
    ).trim();
  }
  return String(value).trim();
}

function lookupValueHasGiven(value) {
  const text = extractLookupText(value);
  if (!text) {
    return false;
  }
  const normalized = normalizeToken(text);
  if (normalized === "given") {
    return true;
  }
  if (normalized.includes("given")) {
    return true;
  }
  // Handles legacy SharePoint lookup serialization such as "1;#Given".
  return /(?:^|[;#,\s])given(?:$|[;#,\s])/i.test(text);
}

function pickClientConsent(fields) {
  const entries = Object.entries(fields || {});
  if (!entries.length) {
    return {
      consentValue: "",
      consented: false,
    };
  }

  const tokenPairs = [];
  for (const [key, value] of entries) {
    const token = normalizeToken(key);
    if (token) {
      tokenPairs.push([token, value]);
    }
  }

  const preferredTokens = new Set([
    "clientconsentformarketing",
    "clientx003aconsentx0020forx0020marketing",
  ]);
  const candidateValues = [];
  const fallbackConsentValues = [];
  for (const [token, value] of tokenPairs) {
    const isPreferred = preferredTokens.has(token);
    const isClientMarketingConsentField =
      token.includes("client") &&
      token.includes("consent") &&
      token.includes("marketing") &&
      !token.endsWith("lookupid");
    const isMarketingConsentField =
      token.includes("consent") && token.includes("marketing") && !token.endsWith("lookupid");
    if (isPreferred || isClientMarketingConsentField || isMarketingConsentField) {
      candidateValues.push(value);
    }
    if (token.includes("consent") && !token.endsWith("lookupid")) {
      fallbackConsentValues.push(value);
    }
  }

  for (const value of candidateValues) {
    if (lookupValueHasGiven(value)) {
      return {
        consentValue: extractLookupText(value),
        consented: true,
      };
    }
  }

  if (candidateValues.length > 0) {
    return {
      consentValue: extractLookupText(candidateValues[0]),
      consented: false,
    };
  }

  for (const value of fallbackConsentValues) {
    if (lookupValueHasGiven(value)) {
      return {
        consentValue: extractLookupText(value),
        consented: true,
      };
    }
  }

  if (fallbackConsentValues.length > 0) {
    return {
      consentValue: extractLookupText(fallbackConsentValues[0]),
      consented: false,
    };
  }

  return {
    consentValue: "",
    consented: false,
  };
}

function isVideoUrl(url) {
  return /\.(mp4|mov|webm)(?:$|[?#])/i.test(String(url || ""));
}

function isImageUrl(url) {
  return /\.(png|jpe?g|gif|webp|bmp)(?:$|[?#])/i.test(String(url || ""));
}

function isPreviewUrl(url) {
  return /\/_layouts\/15\/getpreview/i.test(String(url || ""));
}

function collectUrlsFromValue(value, hostName) {
  const urls = [];
  const pushUrl = (candidate) => {
    const absolute = toAbsoluteUrl(candidate, hostName);
    if (absolute && !urls.includes(absolute)) {
      urls.push(absolute);
    }
  };

  if (!value) {
    return urls;
  }

  if (Array.isArray(value)) {
    for (const entry of value) {
      for (const nested of collectUrlsFromValue(entry, hostName)) {
        pushUrl(nested);
      }
    }
    return urls;
  }

  if (typeof value === "object") {
    pushUrl(value.url);
    pushUrl(value.nativeUrl);
    pushUrl(value.serverRelativeUrl);
    pushUrl(value.serverUrl);
    pushUrl(value.fileUrl);
    pushUrl(value.src);
    return urls;
  }

  if (typeof value === "string") {
    const parsed = parseMaybeJson(value);
    if (parsed && typeof parsed === "object") {
      for (const nested of collectUrlsFromValue(parsed, hostName)) {
        pushUrl(nested);
      }
    }
    pushUrl(value);
    return urls;
  }

  return urls;
}

function pickAssetUrls(fields, hostName, mediaType) {
  const candidates = [];
  const addCandidates = (value) => {
    for (const url of collectUrlsFromValue(value, hostName)) {
      if (!candidates.includes(url)) {
        candidates.push(url);
      }
    }
  };

  addCandidates(fields?.Photo);
  for (const value of Object.values(fields || {})) {
    addCandidates(value);
  }

  const directVideoUrl = candidates.find((url) => isVideoUrl(url)) || "";
  const imagePreviewUrl =
    candidates.find((url) => isImageUrl(url) || isPreviewUrl(url)) || candidates[0] || "";
  const nonPreviewUrl = candidates.find((url) => !isPreviewUrl(url)) || "";

  if (mediaType === "video") {
    return {
      previewUrl: imagePreviewUrl,
      mediaUrl: directVideoUrl || nonPreviewUrl || imagePreviewUrl,
      directVideoUrl,
    };
  }

  return {
    previewUrl: imagePreviewUrl,
    mediaUrl: nonPreviewUrl || imagePreviewUrl,
    directVideoUrl: "",
  };
}

function mapGraphItemToPhoto(item, hostName, listPathname) {
  const fields = item?.fields || {};
  const consent = pickClientConsent(fields);

  const id = String(fields.ID || item?.id || "").trim();
  if (!id) {
    return null;
  }

  const title = String(fields.Title || fields.Client || fields.ClientName || `Photo ${id}`).trim();
  const client = String(fields.Client || fields.ClientName || title).trim();
  const fileName = String(fields.FileLeafRef || fields.FileName || "").trim();
  const fallbackNameWithExtension =
    extractFileNameWithExtension(fileName) ||
    extractFileNameWithExtension(fields.Photo?.fileName || fields.Photo?.FileName || "") ||
    extractFileNameWithExtension(fields.Photo);
  const fallbackName =
    fallbackNameWithExtension ||
    extractKnownFileName(fileName) ||
    extractKnownFileName(fields.Photo?.fileName || fields.Photo?.FileName || "") ||
    extractKnownFileName(fields.Photo);

  const mediaType = inferMediaType(fallbackName || fileName || title);
  const assetUrls = pickAssetUrls(fields, hostName, mediaType);
  const previewUrl = assetUrls.previewUrl;
  const directMediaUrl = assetUrls.mediaUrl;
  let attachmentUrl = "";
  if (fallbackName) {
    const safeListPath = String(listPathname || "")
      .split("/")
      .map((part) => encodeURIComponent(decodeURIComponent(part)))
      .join("/");
    const encodedFile = encodeURIComponent(fallbackName);
    attachmentUrl = `https://${hostName}${safeListPath}/Attachments/${id}/${encodedFile}`;
  }

  const mediaUrl =
    mediaType === "video"
      ? attachmentUrl || directMediaUrl || previewUrl
      : previewUrl || attachmentUrl || directMediaUrl;
  const imageUrl = previewUrl || attachmentUrl;

  if (!imageUrl) {
    logMarketingDebug("skipped-photo-no-image-url", {
      id,
      title,
      client,
      fileName,
      fallbackName,
      previewUrl,
      directMediaUrl,
      attachmentUrl,
      fieldKeys: Object.keys(fields),
    });
    return null;
  }

  const consentDebugFields = Object.entries(fields)
    .filter(([key]) => {
      const token = normalizeToken(key);
      return token.includes("consent") && !token.endsWith("lookupid");
    })
    .map(([key, value]) => ({
      key,
      value: extractLookupText(value),
      rawType: Array.isArray(value) ? "array" : typeof value,
    }));

  logMarketingDebug("parsed-photo-consent", {
    id,
    title,
    client,
    fileName,
    mediaType,
    imageUrl,
    mediaUrl,
    directMediaUrl: assetUrls.directVideoUrl,
    fieldKeys: Object.keys(fields),
    consentValue: consent.consentValue,
    consented: consent.consented,
    consentFields: consentDebugFields,
  });

  return {
    id,
    title,
    client,
    imageUrl,
    mediaUrl,
    attachmentUrl,
    mediaType,
    fileName: fallbackName || fileName || title,
    needsAttachmentLookup: !hasKnownMediaExtension(fallbackName),
    consentValue: consent.consentValue,
    consented: consent.consented,
  };
}

async function resolveAttachmentUrlFromGraph(token, siteId, listId, itemId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${itemId}/driveItem/children?$select=name,webUrl`;
  const data = await fetchJson(url, { headers: graphHeaders(token) });
  const children = Array.isArray(data?.value) ? data.value : [];
  for (const child of children) {
    const name = String(child?.name || "").trim();
    const webUrl = String(child?.webUrl || "").trim();
    if (!webUrl) {
      continue;
    }
    if (hasKnownMediaExtension(name) || hasKnownMediaExtension(webUrl)) {
      return webUrl;
    }
  }
  return "";
}

async function readMarketingPhotos() {
  const token = await getGraphAccessToken();
  const { hostName, sitePath, photosListName } = requireSharePointConfig();
  const siteId = await resolveSiteId(token, hostName, sitePath);
  const list = await resolveList(token, siteId, photosListName);
  const listId = list.id;
  const listPathname = (() => {
    try {
      return new URL(list.webUrl).pathname || `${sitePath}/Lists/${photosListName}`;
    } catch {
      return `${sitePath}/Lists/${photosListName}`;
    }
  })();
  const columns = await resolveListColumns(token, siteId, listId);
  const consentSelectFields = resolveConsentSelectFields(columns);
  const fieldSelect = [
    "ID",
    "Title",
    "Photo",
    "Client",
    "FileLeafRef",
    "FileName",
    ...consentSelectFields,
  ];
  const query = new URLSearchParams({
    $expand: `fields($select=${fieldSelect.join(",")})`,
    $top: "200",
  });

  logMarketingDebug("list-field-select", {
    listId,
    listPathname,
    selectedFields: fieldSelect,
    consentSelectFields,
  });

  const photos = [];
  let totalItemsRead = 0;
  let skippedNoImage = 0;
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${query.toString()}`;

  while (nextUrl) {
    const data = await fetchJson(nextUrl, { headers: graphHeaders(token) });
    const items = Array.isArray(data?.value) ? data.value : [];
    totalItemsRead += items.length;

    for (const item of items) {
      const photo = mapGraphItemToPhoto(item, hostName, listPathname);
      if (photo) {
        if (photo.needsAttachmentLookup) {
          const resolvedUrl = await resolveAttachmentUrlFromGraph(token, siteId, listId, photo.id);
          if (resolvedUrl) {
            photo.attachmentUrl = resolvedUrl;
            if (!photo.imageUrl || !isImageUrl(photo.imageUrl)) {
              photo.imageUrl = resolvedUrl;
            }
            photo.mediaUrl = resolvedUrl;
            photo.fileName = extractFileName(resolvedUrl) || photo.fileName;
          }
        }
        delete photo.needsAttachmentLookup;
        photos.push(photo);
      } else {
        skippedNoImage += 1;
      }
    }

    nextUrl = data?.["@odata.nextLink"] || "";
  }

  photos.sort((a, b) => a.client.localeCompare(b.client, undefined, { sensitivity: "base" }));
  logMarketingDebug("photos-fetch-summary", {
    totalItemsRead,
    returnedPhotos: photos.length,
    skippedNoImage,
  });
  return photos;
}

module.exports = {
  readMarketingPhotos,
};
