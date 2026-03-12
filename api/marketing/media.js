const { requireApiAuth } = require("../_lib/require-api-auth");

const MARKETING_MEDIA_DEBUG = process.env.MARKETING_MEDIA_DEBUG === "1";

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
    listPathFromUrl: decodeLoose(match[1]).trim(),
    itemId: Number(match[2]),
    fileName: decodeLoose(match[3]).trim(),
    normalizedPath: relative,
  };
}

function getIncomingBearerToken(req) {
  const authHeader = String(req.headers.authorization || "");
  const match = /^Bearer\s+(.+)$/i.exec(authHeader);
  return match ? String(match[1] || "").trim() : "";
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

async function getOboToken(assertion, scope) {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_API_CLIENT_ID;
  const clientSecret = process.env.AZURE_API_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing AZURE_TENANT_ID, AZURE_API_CLIENT_ID, or AZURE_API_CLIENT_SECRET.");
  }

  logMediaDebug("obo-token-request-start", { scope });
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    requested_token_use: "on_behalf_of",
    assertion: String(assertion || ""),
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
    const error = new Error(payload?.error_description || payload?.error || `OBO token failed (${response.status}).`);
    error.status = response.status;
    throw error;
  }
  logMediaDebug("obo-token-request-success", { scope });
  return payload.access_token;
}

async function fetchSharePointAttachmentRows(config, listName, listPathFromUrl, itemId, sharePointToken) {
  logMediaDebug("attachment-metadata-start", { listName, listPathFromUrl, itemId });
  const listServerRelative = `${config.sitePath}/Lists/${String(listPathFromUrl || listName || "").trim()}`;
  const url = `https://${config.hostName}${config.sitePath}/_api/web/GetList(@listUrl)/items(${itemId})/AttachmentFiles?$select=FileName,ServerRelativeUrl&@listUrl=${encodeURIComponent(
    quoteODataString(listServerRelative)
  )}`;
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
    const error = new Error(`Attachment metadata request failed (${response.status}).`);
    error.status = response.status;
    throw error;
  }

  const rows = Array.isArray(payload?.value)
    ? payload.value
    : Array.isArray(payload?.d?.results)
      ? payload.d.results
      : [];

  logMediaDebug("attachment-metadata-success", { listName, listPathFromUrl, itemId, count: rows.length });
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
    const error = new Error(`Attachment content request failed (${response.status}): ${detail.slice(0, 300)}`);
    error.status = response.status;
    throw error;
  }

  logMediaDebug("attachment-content-success", { status: response.status });
  return response;
}

async function fetchByKnownAttachmentPath(config, attachment, sharePointToken) {
  const serverRelative = `${config.sitePath}/${attachment.normalizedPath}`.replace(/\/{2,}/g, "/");
  return fetchSharePointFileBinary(config, serverRelative, sharePointToken);
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireApiAuth(req, res, { allowedRoles: ["admin", "marketing", "photo_layout"] }))) {
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

    const spScope = `https://${config.hostName}/.default`;
    const sharePointToken = await getAppToken(spScope);
    const incomingToken = getIncomingBearerToken(req);
    let oboSharePointToken = "";

    let rows = [];
    let upstream = null;

    // Prefer direct attachment fetch from the parsed URL path to avoid list-title mismatches.
    try {
      upstream = await fetchByKnownAttachmentPath(config, attachment, sharePointToken);
    } catch (error) {
      const status = Number(error?.status || 0);
      if ((status === 401 || status === 403) && incomingToken) {
        logMediaDebug("attachment-direct-app-token-denied-retrying-obo", { status });
        oboSharePointToken = await getOboToken(incomingToken, spScope);
        upstream = await fetchByKnownAttachmentPath(config, attachment, oboSharePointToken);
      } else {
        logMediaDebug("attachment-direct-fetch-error", { status, detail: error?.message || String(error) });
      }
    }

    if (!upstream) {
      try {
        rows = await fetchSharePointAttachmentRows(
          config,
          attachment.listNameFromUrl || config.listName,
          attachment.listPathFromUrl || attachment.listNameFromUrl || config.listName,
          attachment.itemId,
          oboSharePointToken || sharePointToken
        );
      } catch (error) {
        const status = Number(error?.status || 0);
        if ((status === 401 || status === 403) && incomingToken && !oboSharePointToken) {
          logMediaDebug("attachment-metadata-app-token-denied-retrying-obo", { status });
          oboSharePointToken = await getOboToken(incomingToken, spScope);
          rows = await fetchSharePointAttachmentRows(
            config,
            attachment.listNameFromUrl || config.listName,
            attachment.listPathFromUrl || attachment.listNameFromUrl || config.listName,
            attachment.itemId,
            oboSharePointToken
          );
        } else {
          throw error;
        }
      }

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

      try {
        upstream = await fetchSharePointFileBinary(
          config,
          String(picked?.ServerRelativeUrl || ""),
          oboSharePointToken || sharePointToken
        );
      } catch (error) {
        const status = Number(error?.status || 0);
        if ((status === 401 || status === 403) && incomingToken && !oboSharePointToken) {
          logMediaDebug("attachment-content-app-token-denied-retrying-obo", { status });
          oboSharePointToken = await getOboToken(incomingToken, spScope);
          upstream = await fetchSharePointFileBinary(
            config,
            String(picked?.ServerRelativeUrl || ""),
            oboSharePointToken
          );
        } else {
          throw error;
        }
      }
    }

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
    const status = Number(error?.status || 0);
    const safeStatus = status >= 400 && status < 600 ? status : 500;
    res.status(safeStatus).json({
      error: "Server error",
      detail: error?.message || String(error),
    });
  }
};
