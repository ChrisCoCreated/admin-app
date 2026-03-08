const { requireApiAuth } = require("../_lib/require-api-auth");

const MARKETING_MEDIA_DEBUG = process.env.MARKETING_MEDIA_DEBUG === "1";

let siteAndListCache = {
  key: "",
  value: null,
  expiresAt: 0,
};

function logMediaDebug(message, details) {
  if (!MARKETING_MEDIA_DEBUG) {
    return;
  }
  if (details !== undefined) {
    console.log(`[marketing-media] ${message}`, details);
    return;
  }
  console.log(`[marketing-media] ${message}`);
}

function decodeLoose(value) {
  let output = String(value || "").replace(/\+/g, " ");
  for (let index = 0; index < 3; index += 1) {
    try {
      const next = decodeURIComponent(output);
      if (next === output) {
        break;
      }
      output = next;
    } catch {
      break;
    }
  }
  return output;
}

function quoteODataString(value) {
  return `'${String(value || "").replace(/'/g, "''")}'`;
}

function getSharePointConfig() {
  const siteUrlRaw = String(process.env.SHAREPOINT_SITE_URL || "").trim();
  if (!siteUrlRaw) {
    throw new Error("Missing SHAREPOINT_SITE_URL.");
  }

  const siteUrl = new URL(siteUrlRaw);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    throw new Error("SHAREPOINT_SITE_URL must include a site path.");
  }

  return {
    hostName: siteUrl.hostname,
    sitePath,
    sitePrefix: `https://${siteUrl.hostname}${sitePath}/`,
    listName: String(process.env.SHAREPOINT_MARKETING_PHOTOS_LIST_NAME || "Photos and Lists").trim(),
  };
}

function parseAttachmentFromUrl(rawUrl, sitePrefix, sitePath) {
  const targetUrl = new URL(rawUrl);
  if (!targetUrl.href.startsWith(sitePrefix)) {
    throw new Error("URL not allowed.");
  }

  const decodedPath = decodeLoose(targetUrl.pathname || "");
  if (!decodedPath.startsWith(sitePath)) {
    throw new Error("URL path not in allowed site.");
  }

  const relative = decodedPath.slice(sitePath.length).replace(/^\/+/, "");
  const match = /^Lists\/(.+)\/Attachments\/(\d+)\/([^/]+)$/i.exec(relative);
  if (!match) {
    throw new Error("Expected a SharePoint attachment URL.");
  }

  return {
    targetUrl,
    listNameFromUrl: decodeLoose(match[1]).trim(),
    itemId: Number(match[2]),
    fileName: decodeLoose(match[3]).trim(),
    normalizedPath: relative,
  };
}

async function fetchJson(url, headers = {}) {
  const response = await fetch(url, { headers: { Accept: "application/json", ...headers } });
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }
  if (!response.ok) {
    throw new Error(payload?.error?.message || text || `Request failed (${response.status}).`);
  }
  return payload;
}

async function getAppToken(scope) {
  logMediaDebug("token-request-start", { scope });
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
    scope,
  });

  const response = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }

  if (!response.ok || !payload?.access_token) {
    throw new Error(payload?.error_description || payload?.error || `Token request failed (${response.status}).`);
  }

  logMediaDebug("token-request-success", { scope });
  return payload.access_token;
}

async function resolveSiteAndListIds(config) {
  const cacheKey = `${config.hostName}|${config.sitePath}|${config.listName}`;
  const now = Date.now();
  if (siteAndListCache.key === cacheKey && siteAndListCache.value && siteAndListCache.expiresAt > now) {
    logMediaDebug("site-list-cache-hit", { cacheKey });
    return siteAndListCache.value;
  }

  logMediaDebug("site-list-resolve-start", {
    hostName: config.hostName,
    sitePath: config.sitePath,
    listName: config.listName,
  });
  const graphToken = await getAppToken("https://graph.microsoft.com/.default");
  const graphHeaders = {
    Authorization: `Bearer ${graphToken}`,
  };

  const site = await fetchJson(
    `https://graph.microsoft.com/v1.0/sites/${config.hostName}:${config.sitePath}?$select=id`,
    graphHeaders
  );

  const listsPayload = await fetchJson(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/lists?$select=id,displayName&$top=200`,
    graphHeaders
  );
  const rows = Array.isArray(listsPayload?.value) ? listsPayload.value : [];
  const wanted = config.listName.toLowerCase();
  const match = rows.find((row) => String(row?.displayName || "").trim().toLowerCase() === wanted);
  if (!match?.id) {
    throw new Error(`Could not find list '${config.listName}'.`);
  }

  const resolved = {
    siteId: site.id,
    listId: match.id,
  };

  siteAndListCache = {
    key: cacheKey,
    value: resolved,
    expiresAt: now + 10 * 60 * 1000,
  };

  logMediaDebug("site-list-resolve-success", {
    siteId: resolved.siteId,
    listId: resolved.listId,
    listName: config.listName,
  });
  return resolved;
}

async function fetchSharePointAttachmentRows(config, listId, itemId, sharePointToken) {
  logMediaDebug("attachment-metadata-start", { listId, itemId });
  const url = `https://${config.hostName}${config.sitePath}/_api/web/lists(guid'${String(listId).toUpperCase()}')/items(${itemId})/AttachmentFiles?$select=FileName,ServerRelativeUrl`;
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${sharePointToken}`,
      Accept: "application/json;odata=nometadata",
    },
  });

  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }

  if (!response.ok) {
    throw new Error(`Attachment metadata request failed (${response.status}).`);
  }

  const rows = Array.isArray(payload?.value)
    ? payload.value
    : Array.isArray(payload?.d?.results)
      ? payload.d.results
      : [];

  logMediaDebug("attachment-metadata-success", { listId, itemId, count: rows.length });
  return rows;
}

async function fetchSharePointFileBinary(config, serverRelativeUrl, sharePointToken) {
  const rel = String(serverRelativeUrl || "").trim();
  if (!rel) {
    throw new Error("Missing ServerRelativeUrl for attachment.");
  }

  const decodedRel = decodeLoose(rel);
  const fileApi = `https://${config.hostName}${config.sitePath}/_api/web/GetFileByServerRelativePath(decodedurl=${quoteODataString(
    decodedRel
  )})/$value`;
  logMediaDebug("attachment-content-start", { fileApi });

  let response = await fetch(fileApi, {
    headers: {
      Authorization: `Bearer ${sharePointToken}`,
      Accept: "*/*",
    },
  });

  if (!response.ok) {
    const absolute = `https://${config.hostName}${decodedRel.startsWith("/") ? decodedRel : `/${decodedRel}`}`;
    response = await fetch(absolute, {
      headers: {
        Authorization: `Bearer ${sharePointToken}`,
        Accept: "*/*",
      },
    });
  }

  if (!response.ok) {
    const detail = await response.text();
    throw new Error(`Attachment content request failed (${response.status}): ${detail.slice(0, 300)}`);
  }

  logMediaDebug("attachment-content-success", { status: response.status });
  return response;
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "marketing"] }))) {
    return;
  }

  const rawUrl = String(req.query.url || "").trim();
  if (!rawUrl) {
    res.status(400).json({ error: "Missing url query parameter." });
    return;
  }

  try {
    const config = getSharePointConfig();
    const attachment = parseAttachmentFromUrl(rawUrl, config.sitePrefix, config.sitePath);

    logMediaDebug("normalized-attachment-tuple", {
      listName: attachment.listNameFromUrl,
      itemId: attachment.itemId,
      fileName: attachment.fileName,
      normalizedPath: attachment.normalizedPath,
    });

    const { listId } = await resolveSiteAndListIds(config);
    const sharePointToken = await getAppToken(`https://${config.hostName}/.default`);

    const rows = await fetchSharePointAttachmentRows(config, listId, attachment.itemId, sharePointToken);
    if (!rows.length) {
      res.status(404).json({ error: "Attachment not found for item.", detail: `itemId=${attachment.itemId}` });
      return;
    }

    const normalizedRequested = attachment.fileName.toLowerCase();
    const exact = rows.find(
      (row) => String(row?.FileName || "").trim().toLowerCase() === normalizedRequested
    );
    const picked = exact || rows[0];

    logMediaDebug("attachment-selected", {
      requested: attachment.fileName,
      selected: String(picked?.FileName || ""),
      itemId: attachment.itemId,
    });

    const upstream = await fetchSharePointFileBinary(
      config,
      String(picked?.ServerRelativeUrl || ""),
      sharePointToken
    );

    const mimeType = upstream.headers.get("content-type") || "application/octet-stream";
    const arrayBuffer = await upstream.arrayBuffer();
    const dataBase64 = Buffer.from(arrayBuffer).toString("base64");

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      mimeType,
      dataBase64,
    });
  } catch (error) {
    logMediaDebug("media-endpoint-error", {
      detail: error?.message || String(error),
      stack: String(error?.stack || "").split("\n").slice(0, 4).join(" | "),
    });
    res.status(500).json({
      error: "Server error",
      detail: error?.message || String(error),
    });
  }
};
