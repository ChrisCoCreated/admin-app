const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { createGraphDelegatedClient } = require("../_lib/tasks/graph-delegated-client");
const { RECRUITMENT_ALLOWED_ROLES } = require("../_lib/recruitment-access");

const DEFAULT_SITE_URL = "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency";
const DEFAULT_LIST_NAME = "Associate Recruitment";

const ALLOWED_ROLES = RECRUITMENT_ALLOWED_ROLES;

function normalizeText(value) {
  return String(value || "").trim();
}

function toBoolean(value) {
  if (value === true || value === false) {
    return value;
  }
  const normalized = normalizeText(value).toLowerCase();
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "y";
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

function parsePersonField(value, lookupIdValue) {
  const text = normalizeText(value);
  if (text) {
    return text;
  }
  const lookupId = normalizeText(lookupIdValue);
  return lookupId ? `User #${lookupId}` : "";
}

function parseHyperlink(value) {
  if (value && typeof value === "object") {
    const fromObject = normalizeText(value.Url || value.url || value.Description || value.description);
    if (fromObject) {
      return fromObject;
    }
  }
  const text = normalizeText(value);
  if (!text) {
    return "";
  }
  const commaIndex = text.indexOf(",");
  if (commaIndex <= 0) {
    return text;
  }
  return text.slice(0, commaIndex).trim();
}

function extractIndeedProfileUrl(notes) {
  const text = normalizeText(notes);
  if (!text) {
    return "";
  }
  const matched = text.match(/Indeed\s+Profile:\s*(https?:\/\/\S+)/i);
  return matched ? normalizeText(matched[1]) : "";
}

function normalizePhoneKey(value) {
  return normalizeText(value).replace(/\D/g, "");
}

function normalizeRecruitmentItem(item) {
  const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
  const active = toBoolean(fields.Active);
  const interviewWith = parsePersonField(fields.InterviewWith, fields.InterviewWithLookupId);

  return {
    id: normalizeText(item?.id),
    candidateName: normalizeText(fields.Title),
    location: normalizeText(fields.Location),
    email: normalizeText(fields.Email),
    phoneNumber: normalizeText(fields.PhoneNumber),
    interviewBooked: toBoolean(fields.InterviewBooked),
    interviewWith,
    status: normalizeText(fields.Status),
    active,
    keepInMind: toBoolean(fields.KeepinMind),
    livesIn: normalizeText(fields.LivesIn),
    firstInterviewDate: normalizeText(fields._x0031_stInterviewDate),
    notes: normalizeText(fields.Notes),
    indeedProfileUrl: normalizeText(fields.IndeedURL) || extractIndeedProfileUrl(fields.Notes),
    source: normalizeText(fields.Source),
    earmarkedFor: normalizeText(fields.EarmarkedFor),
    oneTouchLink: parseHyperlink(fields.OnetouchLink),
    created: normalizeText(fields.Created || item?.createdDateTime),
  };
}

async function fetchRecruitmentItemsByPhone(graphClient, siteId, listId) {
  const params = new URLSearchParams({
    $top: "200",
    $expand:
      "fields($select=Title,Email,PhoneNumber,LivesIn,Location,Source,Status,Notes,IndeedURL,Active,InterviewBooked,InterviewWith,InterviewWithLookupId,KeepinMind,_x0031_stInterviewDate,EarmarkedFor,OnetouchLink,Created)",
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${params.toString()}`;
  const items = await graphClient.fetchAllPages(url);
  return items.map(normalizeRecruitmentItem);
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

  const candidateName = normalizeText(req.body?.candidateName);
  const indeedUrl = normalizeText(req.body?.indeedUrl);
  if (!candidateName) {
    res.status(400).json({
      error: {
        code: "BAD_REQUEST",
        message: "Candidate name is required.",
      },
    });
    return;
  }

  try {
    const config = parseSiteConfig();
    const graphClient = createGraphDelegatedClient(req.authUser?.graphAccessToken);
    const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
    const listId = await resolveListId(graphClient, siteId, config.listName);
    const phoneKey = normalizePhoneKey(req.body?.phoneNumber);
    const shouldUpdateExisting = req.body?.updateExistingByPhone === true && Boolean(phoneKey);

    if (shouldUpdateExisting) {
      const existingItems = await fetchRecruitmentItemsByPhone(graphClient, siteId, listId);
      const matched = existingItems.find((item) => normalizePhoneKey(item.phoneNumber) === phoneKey);
      if (matched?.id) {
        const patchFields = {
          Title: candidateName || matched.candidateName,
          Email: normalizeText(req.body?.email) || matched.email,
          PhoneNumber: normalizeText(req.body?.phoneNumber) || matched.phoneNumber,
          LivesIn: normalizeText(req.body?.livesIn) || matched.livesIn,
          Location: normalizeText(req.body?.location) || matched.location,
          Source: normalizeText(req.body?.source) || matched.source,
          Status: normalizeText(req.body?.status) || matched.status || "Initial Call",
          Notes: normalizeText(req.body?.notes) || matched.notes,
          IndeedURL: indeedUrl || matched.indeedProfileUrl,
          Active: req.body?.active === undefined ? matched.active : toBoolean(req.body?.active),
        };
        const patchUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(matched.id)}/fields`;
        await graphClient.fetchJson(patchUrl, {
          method: "PATCH",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify(patchFields),
        });
        const itemUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(
          matched.id
        )}?$expand=fields`;
        const fullItem = await graphClient.fetchJson(itemUrl);
        res.setHeader("Cache-Control", "no-store");
        res.status(200).json({
          success: true,
          updatedExisting: true,
          item: normalizeRecruitmentItem(fullItem),
        });
        return;
      }
    }

    const fields = {
      Title: candidateName,
      Email: normalizeText(req.body?.email),
      PhoneNumber: normalizeText(req.body?.phoneNumber),
      LivesIn: normalizeText(req.body?.livesIn),
      Location: normalizeText(req.body?.location),
      Source: normalizeText(req.body?.source),
      Status: normalizeText(req.body?.status) || "Initial Call",
      Notes: normalizeText(req.body?.notes),
      IndeedURL: indeedUrl,
      Active: req.body?.active === undefined ? true : toBoolean(req.body?.active),
    };

    const createUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
    const created = await graphClient.fetchJson(createUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ fields }),
    });

    const itemUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(
      created?.id
    )}?$expand=fields`;
    const fullItem = await graphClient.fetchJson(itemUrl);

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      success: true,
      updatedExisting: false,
      item: normalizeRecruitmentItem(fullItem),
    });
  } catch (error) {
    res.status(Number(error?.status) || 502).json({
      error: {
        code: String(error?.code || "RECRUITMENT_CREATE_FAILED"),
        message: error?.message || "Could not create recruitment candidate.",
      },
    });
  }
};
