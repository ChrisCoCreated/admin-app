const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { RECRUITMENT_ALLOWED_ROLES } = require("../_lib/recruitment-access");

const DEFAULT_SITE_URL = "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency";
const DEFAULT_LIST_NAME = "Associate Recruitment";
const CURRENT_OWNER_OPTIONS = new Set([
  "chris@planwithcare.co.uk",
  "rebecca@planwithcare.co.uk",
  "michalina@thrivehomecare.co.uk",
  "peter@planwithcare.co.uk",
]);

const ALLOWED_ROLES = RECRUITMENT_ALLOWED_ROLES;

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeEmail(value) {
  return normalizeText(value).toLowerCase();
}

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function parseSiteConfig() {
  const siteUrlValue = normalizeText(process.env.SHAREPOINT_RECRUITMENT_SITE_URL || DEFAULT_SITE_URL);
  const listName = normalizeText(process.env.SHAREPOINT_RECRUITMENT_LIST_NAME || DEFAULT_LIST_NAME);
  if (!siteUrlValue || !listName) {
    throw new Error("Missing SHAREPOINT_RECRUITMENT_SITE_URL or SHAREPOINT_RECRUITMENT_LIST_NAME.");
  }
  const siteUrl = new URL(siteUrlValue);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    throw new Error("SHAREPOINT_RECRUITMENT_SITE_URL must include a site path.");
  }
  return {
    hostName: siteUrl.hostname,
    sitePath,
    siteBaseUrl: `${siteUrl.origin}${sitePath}`,
    listName,
  };
}

async function getAppToken(scope) {
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
    const error = new Error(payload?.error_description || payload?.error || `Token request failed (${response.status}).`);
    error.status = response.status;
    throw error;
  }

  return payload.access_token;
}

async function fetchSharePointJson(url, token, options = {}) {
  const response = await fetch(url, {
    method: options.method || "GET",
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json;odata=verbose",
      ...(options.headers || {}),
    },
    body: options.body,
  });
  const text = await response.text();
  let payload = null;
  try {
    payload = text ? JSON.parse(text) : null;
  } catch {
    payload = null;
  }

  if (!response.ok) {
    const message =
      payload?.error?.message?.value ||
      payload?.odata?.error?.message?.value ||
      payload?.error_description ||
      text ||
      `SharePoint request failed (${response.status}).`;
    const error = new Error(message);
    error.status = response.status;
    throw error;
  }

  return payload;
}

async function getFormDigest(siteBaseUrl, token) {
  const payload = await fetchSharePointJson(`${siteBaseUrl}/_api/contextinfo`, token, {
    method: "POST",
    headers: {
      "Content-Type": "application/json;odata=verbose",
    },
  });
  return normalizeText(payload?.d?.GetContextWebInformation?.FormDigestValue);
}

async function resolveListInfo(siteBaseUrl, listName, token) {
  const url = `${siteBaseUrl}/_api/web/lists?$select=Id,Title,ListItemEntityTypeFullName&$filter=Title eq ${encodeURIComponent(
    quoteODataString(listName)
  )}`;
  const payload = await fetchSharePointJson(url, token);
  const list = Array.isArray(payload?.d?.results) ? payload.d.results[0] : null;
  if (!list?.Id || !list?.ListItemEntityTypeFullName) {
    throw new Error(`Could not resolve SharePoint list '${listName}' over REST.`);
  }
  return {
    id: normalizeText(list.Id),
    entityTypeName: normalizeText(list.ListItemEntityTypeFullName),
  };
}

async function ensureUser(siteBaseUrl, token, digest, email) {
  const payload = await fetchSharePointJson(`${siteBaseUrl}/_api/web/ensureuser`, token, {
    method: "POST",
    headers: {
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
    },
    body: JSON.stringify({
      logonName: `i:0#.f|membership|${normalizeEmail(email)}`,
    }),
  });

  const userId = Number(payload?.d?.Id);
  if (!Number.isFinite(userId) || userId <= 0) {
    throw new Error(`Could not resolve SharePoint user lookup id for ${email}.`);
  }
  return userId;
}

async function patchCurrentOwner({ siteBaseUrl, listName, itemId, ownerEmail, token }) {
  const digest = await getFormDigest(siteBaseUrl, token);
  const listInfo = await resolveListInfo(siteBaseUrl, listName, token);
  const ownerLookupId = ownerEmail ? await ensureUser(siteBaseUrl, token, digest, ownerEmail) : null;
  const payload = {
    __metadata: { type: listInfo.entityTypeName },
    CurrentOwnerId: ownerLookupId,
  };

  await fetchSharePointJson(`${siteBaseUrl}/_api/web/lists(guid'${listInfo.id}')/items(${encodeURIComponent(itemId)})`, token, {
    method: "POST",
    headers: {
      "Content-Type": "application/json;odata=verbose",
      "X-RequestDigest": digest,
      "IF-MATCH": "*",
      "X-HTTP-Method": "MERGE",
    },
    body: JSON.stringify(payload),
  });

  return ownerLookupId;
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({
      error: {
        code: "METHOD_NOT_ALLOWED",
        message: "Method Not Allowed",
      },
    });
    return;
  }

  if (!(await requireGraphAuth(req, res, { allowedRoles: ALLOWED_ROLES }))) {
    return;
  }

  const itemId = normalizeText(req.body?.itemId);
  const ownerEmail = normalizeEmail(req.body?.currentOwnerEmail);
  if (!itemId) {
    res.status(400).json({
      error: {
        code: "BAD_REQUEST",
        message: "Missing itemId.",
      },
    });
    return;
  }
  if (ownerEmail && !CURRENT_OWNER_OPTIONS.has(ownerEmail)) {
    res.status(400).json({
      error: {
        code: "BAD_REQUEST",
        message: "Invalid current owner.",
      },
    });
    return;
  }

  try {
    const config = parseSiteConfig();
    const scope = `https://${config.hostName}/.default`;
    const appToken = await getAppToken(scope);
    const ownerLookupId = await patchCurrentOwner({
      siteBaseUrl: config.siteBaseUrl,
      listName: config.listName,
      itemId,
      ownerEmail,
      token: appToken,
    });

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      success: true,
      itemId,
      currentOwnerEmail: ownerEmail,
      ownerLookupId,
    });
  } catch (error) {
    console.error("[recruitment-owner] CurrentOwner save failed", {
      itemId,
      ownerEmail,
      message: error?.message || String(error),
      status: Number(error?.status || 0) || undefined,
    });
    res.status(Number(error?.status) || 502).json({
      error: {
        code: String(error?.code || "OWNER_PATCH_FAILED"),
        message: error?.message || "Could not update recruitment owner.",
      },
    });
  }
};
