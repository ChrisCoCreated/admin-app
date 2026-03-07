const { requireGraphAuth } = require("../_lib/require-graph-auth");

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

function getSharePointSiteParts() {
  const raw = String(process.env.SHAREPOINT_SITE_URL || "").trim();
  if (!raw) {
    throw new Error("Missing SHAREPOINT_SITE_URL.");
  }
  const siteUrl = new URL(raw);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  return {
    hostName: siteUrl.hostname,
    sitePath,
  };
}

function getAllowedSharePointPrefix() {
  const { hostName, sitePath } = getSharePointSiteParts();
  return `https://${hostName}${sitePath}/`;
}

function quoteODataString(value) {
  return `'${String(value || "").replace(/'/g, "''")}'`;
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

function parseAttachmentPath(targetUrl, sitePath) {
  const decodedPath = decodeLoose(String(targetUrl.pathname || ""));
  if (!decodedPath.startsWith(sitePath)) {
    return null;
  }
  const relative = decodedPath.slice(sitePath.length).replace(/^\/+/, "");
  const match = /^Lists\/(.+)\/Attachments\/(\d+)\/([^/]+)$/i.exec(relative);
  if (!match) {
    return null;
  }
  return {
    listName: decodeLoose(match[1]).trim(),
    itemId: Number(match[2]),
    fileName: decodeLoose(match[3]).trim(),
    normalizedPath: relative,
  };
}

async function fetchGraphJson(url, graphAccessToken) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${graphAccessToken}`,
      Accept: "application/json",
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
    const error = new Error(payload?.error?.message || text || `Graph API failed (${response.status}).`);
    error.status = response.status;
    throw error;
  }
  return payload;
}

async function fetchGraphBinary(url, graphAccessToken, label) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${graphAccessToken}`,
    },
  });
  if (!response.ok) {
    const detail = await response.text();
    const error = new Error(`${label} failed (${response.status}): ${detail.slice(0, 300)}`);
    error.status = response.status;
    throw error;
  }
  return response;
}

async function exchangeForDelegatedSharePointToken(userAccessToken, hostName) {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_API_CLIENT_ID;
  const clientSecret = process.env.AZURE_API_CLIENT_SECRET;
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Missing AZURE_TENANT_ID, AZURE_API_CLIENT_ID, or AZURE_API_CLIENT_SECRET.");
  }
  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
    requested_token_use: "on_behalf_of",
    assertion: String(userAccessToken || ""),
    scope: `https://${hostName}/.default`,
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
    throw new Error(payload?.error_description || payload?.error || `OBO token failed (${response.status}).`);
  }
  return payload.access_token;
}

async function fetchGraphSiteId(graphAccessToken, hostName, sitePath) {
  const payload = await fetchGraphJson(
    `https://graph.microsoft.com/v1.0/sites/${hostName}:${sitePath}?$select=id`,
    graphAccessToken
  );
  if (!payload?.id) {
    throw new Error("Could not resolve site id.");
  }
  return payload.id;
}

async function fetchGraphListIdByName(graphAccessToken, siteId, listName) {
  function normalizeListKey(value) {
    return decodeLoose(String(value || ""))
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  }

  function getListPathTail(webUrl) {
    try {
      const pathname = decodeLoose(new URL(String(webUrl || "")).pathname || "");
      const parts = pathname.split("/").filter(Boolean);
      return parts.length ? parts[parts.length - 1] : "";
    } catch {
      return "";
    }
  }

  const wanted = normalizeListKey(listName);
  const rows = [];
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?$select=id,displayName,webUrl&$top=200`;
  while (nextUrl) {
    const payload = await fetchGraphJson(nextUrl, graphAccessToken);
    const pageRows = Array.isArray(payload?.value) ? payload.value : [];
    rows.push(...pageRows);
    nextUrl = String(payload?.["@odata.nextLink"] || "").trim();
  }

  const match = rows.find((row) => {
    const display = normalizeListKey(row?.displayName);
    const tail = normalizeListKey(getListPathTail(row?.webUrl));
    return display === wanted || tail === wanted;
  });

  if (!match?.id) {
    throw new Error(`Could not find list '${listName}'.`);
  }
  return match.id;
}

async function fetchGraphListDriveId(graphAccessToken, siteId, listId) {
  const payload = await fetchGraphJson(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/drive?$select=id,webUrl`,
    graphAccessToken
  );
  if (!payload?.id) {
    throw new Error("Could not resolve list drive id.");
  }
  return {
    id: payload.id,
    webUrl: String(payload.webUrl || "").trim(),
  };
}

async function fetchViaDelegatedListItemDriveItem(targetUrl, graphAccessToken) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const parsed = parseAttachmentPath(targetUrl, sitePath);
  if (!parsed || !parsed.listName || !Number.isFinite(parsed.itemId)) {
    return null;
  }

  logMediaDebug("delegated-list-item-driveitem-start", {
    listName: parsed.listName,
    itemId: parsed.itemId,
    fileName: parsed.fileName,
  });

  const siteId = await fetchGraphSiteId(graphAccessToken, hostName, sitePath);
  const listId = await fetchGraphListIdByName(graphAccessToken, siteId, parsed.listName);
  const contentUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${parsed.itemId}/driveItem/content`;
  logMediaDebug("delegated-list-item-driveitem-content-url", { contentUrl });

  const response = await fetchGraphBinary(contentUrl, graphAccessToken, "Delegated list item driveItem fetch");
  logMediaDebug("delegated-list-item-driveitem-success", { status: response.status });
  return response;
}

function toGraphDrivePathForAttachments(targetUrl, sitePath) {
  const decodedPath = decodeLoose(String(targetUrl.pathname || ""));
  if (!decodedPath.startsWith(sitePath)) {
    return "";
  }
  const relative = decodedPath.slice(sitePath.length).replace(/^\/+/, "");
  if (!/^Lists\/.+\/Attachments\/\d+\/.+/i.test(relative)) {
    return "";
  }
  return relative;
}

async function fetchViaDelegatedSharePointAttachment(targetUrl, graphAccessToken) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const parsed = parseAttachmentPath(targetUrl, sitePath);
  if (!parsed || !parsed.listName || !Number.isFinite(parsed.itemId)) {
    return null;
  }

  logMediaDebug("delegated-sharepoint-obo-start", {
    listName: parsed.listName,
    itemId: parsed.itemId,
    fileName: parsed.fileName,
  });

  const spToken = await exchangeForDelegatedSharePointToken(graphAccessToken, hostName);
  const attachmentsUrl = `https://${hostName}${sitePath}/_api/web/lists/GetByTitle(${quoteODataString(
    parsed.listName
  )})/items(${parsed.itemId})/AttachmentFiles?$select=FileName,ServerRelativeUrl`;
  const metaResponse = await fetch(attachmentsUrl, {
    headers: {
      Authorization: `Bearer ${spToken}`,
      Accept: "application/json;odata=nometadata",
    },
  });
  const metaText = await metaResponse.text();
  let metaPayload = null;
  try {
    metaPayload = metaText ? JSON.parse(metaText) : null;
  } catch {
    metaPayload = null;
  }
  if (!metaResponse.ok) {
    throw new Error(`SharePoint attachment metadata failed (${metaResponse.status}).`);
  }

  const rows = Array.isArray(metaPayload?.value)
    ? metaPayload.value
    : Array.isArray(metaPayload?.d?.results)
      ? metaPayload.d.results
      : [];
  if (!rows.length) {
    throw new Error("No attachment rows returned from SharePoint.");
  }

  const normalizeName = (value) => decodeLoose(String(value || "")).trim().toLowerCase();
  const requested = normalizeName(parsed.fileName);
  const exact = rows.find((row) => normalizeName(row?.FileName) === requested);
  const picked = exact || rows[0];
  const serverRelative = String(picked?.ServerRelativeUrl || "").trim();
  if (!serverRelative) {
    throw new Error("Attachment row missing ServerRelativeUrl.");
  }

  const absolute = `https://${hostName}${serverRelative.startsWith("/") ? serverRelative : `/${serverRelative}`}`;
  logMediaDebug("delegated-sharepoint-obo-content-url", { absolute });
  const fileResponse = await fetch(absolute, {
    headers: {
      Authorization: `Bearer ${spToken}`,
    },
  });
  if (!fileResponse.ok) {
    const detail = await fileResponse.text();
    throw new Error(`SharePoint attachment content failed (${fileResponse.status}): ${detail.slice(0, 200)}`);
  }
  logMediaDebug("delegated-sharepoint-obo-success", { status: fileResponse.status });
  return fileResponse;
}

async function fetchViaDelegatedListDrive(targetUrl, graphAccessToken) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const parsed = parseAttachmentPath(targetUrl, sitePath);
  if (!parsed || !parsed.listName || !parsed.fileName || !Number.isFinite(parsed.itemId)) {
    return null;
  }

  logMediaDebug("delegated-list-drive-start", {
    listName: parsed.listName,
    itemId: parsed.itemId,
    fileName: parsed.fileName,
  });

  const siteId = await fetchGraphSiteId(graphAccessToken, hostName, sitePath);
  const listId = await fetchGraphListIdByName(graphAccessToken, siteId, parsed.listName);
  const drive = await fetchGraphListDriveId(graphAccessToken, siteId, listId);
  const encodedSegments = ["Attachments", String(parsed.itemId), parsed.fileName]
    .map((part) => encodeURIComponent(part))
    .join("/");
  const contentUrl = `https://graph.microsoft.com/v1.0/drives/${drive.id}/root:/${encodedSegments}:/content`;
  logMediaDebug("delegated-list-drive-content-url", { contentUrl, driveWebUrl: drive.webUrl });

  const response = await fetchGraphBinary(contentUrl, graphAccessToken, "Delegated list drive fetch");
  logMediaDebug("delegated-list-drive-success", { status: response.status });
  return response;
}

async function fetchViaDelegatedSiteDrive(targetUrl, graphAccessToken) {
  const { hostName, sitePath } = getSharePointSiteParts();
  const graphDrivePath = toGraphDrivePathForAttachments(targetUrl, sitePath);
  if (!graphDrivePath) {
    return null;
  }

  logMediaDebug("delegated-site-drive-start", { graphDrivePath });
  const siteId = await fetchGraphSiteId(graphAccessToken, hostName, sitePath);
  const encodedPath = graphDrivePath
    .split("/")
    .map((part) => encodeURIComponent(part))
    .join("/");
  const contentUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/content`;
  logMediaDebug("delegated-site-drive-content-url", { contentUrl });

  const response = await fetchGraphBinary(contentUrl, graphAccessToken, "Delegated site drive fetch");
  logMediaDebug("delegated-site-drive-success", { status: response.status });
  return response;
}

function toGraphShareTokenFromUrl(absoluteUrl) {
  const base64 = Buffer.from(String(absoluteUrl || ""), "utf8").toString("base64");
  const base64Url = base64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
  return `u!${base64Url}`;
}

async function fetchViaDelegatedShares(targetUrl, graphAccessToken) {
  const shareToken = toGraphShareTokenFromUrl(targetUrl.href);
  const contentUrl = `https://graph.microsoft.com/v1.0/shares/${shareToken}/driveItem/content`;
  logMediaDebug("delegated-shares-start", { contentUrl });
  const response = await fetchGraphBinary(contentUrl, graphAccessToken, "Delegated shares fetch");
  logMediaDebug("delegated-shares-success", { status: response.status });
  return response;
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  if (!(await requireGraphAuth(req, res, { allowedRoles: ["admin", "marketing"] }))) {
    return;
  }

  const rawUrl = String(req.query.url || "").trim();
  if (!rawUrl) {
    res.status(400).json({ error: "Missing url query parameter." });
    return;
  }

  try {
    const targetUrl = new URL(rawUrl);
    const allowedPrefix = getAllowedSharePointPrefix();
    if (!targetUrl.href.startsWith(allowedPrefix)) {
      res.status(403).json({ error: "URL not allowed." });
      return;
    }

    const { sitePath } = getSharePointSiteParts();
    const parsedAttachment = parseAttachmentPath(targetUrl, sitePath);
    if (parsedAttachment) {
      logMediaDebug("normalized-attachment-tuple", {
        listName: parsedAttachment.listName,
        itemId: parsedAttachment.itemId,
        fileName: parsedAttachment.fileName,
      });
    }

    const graphAccessToken = req.authUser?.graphAccessToken;
    if (!graphAccessToken) {
      res.status(401).json({ error: "Unauthorized." });
      return;
    }

    const resolverErrors = [];
    let upstream = null;

    try {
      upstream = await fetchViaDelegatedListDrive(targetUrl, graphAccessToken);
    } catch (error) {
      resolverErrors.push(`list-drive:${error?.message || String(error)}`);
      logMediaDebug("delegated-list-drive-error", { detail: error?.message || String(error) });
    }

    if (!upstream) {
      try {
        upstream = await fetchViaDelegatedListItemDriveItem(targetUrl, graphAccessToken);
      } catch (error) {
        resolverErrors.push(`list-item-driveitem:${error?.message || String(error)}`);
        logMediaDebug("delegated-list-item-driveitem-error", { detail: error?.message || String(error) });
      }
    }

    if (!upstream) {
      try {
        upstream = await fetchViaDelegatedSharePointAttachment(targetUrl, graphAccessToken);
      } catch (error) {
        resolverErrors.push(`sharepoint-obo:${error?.message || String(error)}`);
        logMediaDebug("delegated-sharepoint-obo-error", { detail: error?.message || String(error) });
      }
    }

    if (!upstream) {
      try {
        upstream = await fetchViaDelegatedSiteDrive(targetUrl, graphAccessToken);
      } catch (error) {
        resolverErrors.push(`site-drive:${error?.message || String(error)}`);
        logMediaDebug("delegated-site-drive-error", { detail: error?.message || String(error) });
      }
    }

    if (!upstream) {
      try {
        upstream = await fetchViaDelegatedShares(targetUrl, graphAccessToken);
      } catch (error) {
        resolverErrors.push(`shares:${error?.message || String(error)}`);
        logMediaDebug("delegated-shares-error", { detail: error?.message || String(error) });
      }
    }

    if (!upstream) {
      const normalizedPath = parsedAttachment?.normalizedPath || decodeLoose(String(targetUrl.pathname || ""));
      res.status(401).json({
        error: "Upstream request failed.",
        detail: `All delegated resolvers failed for '${normalizedPath}'. ${resolverErrors.join(" | ")}`.slice(
          0,
          500
        ),
      });
      return;
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
    res.status(500).json({
      error: "Server error",
      detail: error && error.message ? error.message : String(error),
    });
  }
};
