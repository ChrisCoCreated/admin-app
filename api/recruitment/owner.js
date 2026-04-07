const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { createGraphDelegatedClient } = require("../_lib/tasks/graph-delegated-client");
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

function extractEmail(value) {
  const matched = String(value || "").match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  return matched ? normalizeEmail(matched[0]) : "";
}

function inferOwnerEmail(value) {
  const normalized = normalizeEmail(value);
  if (!normalized) {
    return "";
  }
  if (CURRENT_OWNER_OPTIONS.has(normalized)) {
    return normalized;
  }
  const extracted = extractEmail(value);
  if (CURRENT_OWNER_OPTIONS.has(extracted)) {
    return extracted;
  }
  return "";
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
    listName,
  };
}

async function resolveSiteId(graphClient, hostName, sitePath) {
  const url = `https://graph.microsoft.com/v1.0/sites/${hostName}:${sitePath}?$select=id`;
  const payload = await graphClient.fetchJson(url);
  if (!payload?.id) {
    throw new Error("Could not resolve SharePoint site id.");
  }
  return payload.id;
}

async function resolveListId(graphClient, siteId, listName) {
  const params = new URLSearchParams({
    $select: "id,displayName",
    $filter: `displayName eq ${quoteODataString(listName)}`,
    $top: "1",
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?${params.toString()}`;
  const payload = await graphClient.fetchJson(url);
  const list = Array.isArray(payload?.value) ? payload.value[0] : null;
  if (!list?.id) {
    throw new Error(`Could not find SharePoint list '${listName}'.`);
  }
  return String(list.id);
}

async function inferCurrentOwnerLookupId(graphClient, siteId, listId, ownerEmail) {
  const normalizedOwnerEmail = normalizeEmail(ownerEmail);
  if (!normalizedOwnerEmail) {
    return null;
  }

  const params = new URLSearchParams({
    $top: "200",
    $expand: "fields($select=CurrentOwner,CurrentOwnerLookupId)",
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${params.toString()}`;
  const payload = await graphClient.fetchJson(url);
  const items = Array.isArray(payload?.value) ? payload.value : [];

  for (const item of items) {
    const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
    const inferredEmail = inferOwnerEmail(fields.CurrentOwner);
    const lookupId = Number(fields.CurrentOwnerLookupId);
    if (inferredEmail === normalizedOwnerEmail && Number.isFinite(lookupId) && lookupId > 0) {
      return lookupId;
    }
  }

  return null;
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
    const graphClient = createGraphDelegatedClient(req.authUser?.graphAccessToken);
    const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
    const listId = await resolveListId(graphClient, siteId, config.listName);
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}/fields`;
    const lookupId = ownerEmail ? await inferCurrentOwnerLookupId(graphClient, siteId, listId, ownerEmail) : null;
    const patchBody = lookupId
      ? { CurrentOwnerLookupId: lookupId }
      : { CurrentOwner: ownerEmail || "" };

    await graphClient.fetchJson(url, {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(patchBody),
    });

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      success: true,
      itemId,
      currentOwnerEmail: ownerEmail,
      lookupId: lookupId || null,
    });
  } catch (error) {
    res.status(Number(error?.status) || 502).json({
      error: {
        code: String(error?.code || "OWNER_PATCH_FAILED"),
        message: error?.message || "Could not update recruitment owner.",
      },
    });
  }
};
