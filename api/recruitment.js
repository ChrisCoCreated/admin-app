const { requireGraphAuth } = require("./_lib/require-graph-auth");
const { createGraphDelegatedClient } = require("./_lib/tasks/graph-delegated-client");

const DEFAULT_SITE_URL = "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency";
const DEFAULT_LIST_NAME = "Associate Recruitment";
const DEFAULT_LIST_WEB_URL =
  "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency/Lists/Associate%20Recruitment/Active.aspx?env=WebViewList";

const ALLOWED_ROLES = [
  "admin",
  "care_manager",
  "operations",
  "hr_only",
  "hr_clients",
  "time_hr",
  "time_hr_clients",
];

function normalizeText(value) {
  return String(value || "").trim();
}

function toBoolean(value) {
  if (value === true || value === false) {
    return value;
  }

  const normalized = String(value || "")
    .trim()
    .toLowerCase();
  if (!normalized) {
    return false;
  }
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "y";
}

function parseSiteConfig() {
  const siteUrlValue = normalizeText(process.env.SHAREPOINT_RECRUITMENT_SITE_URL || DEFAULT_SITE_URL);
  const listName = normalizeText(process.env.SHAREPOINT_RECRUITMENT_LIST_NAME || DEFAULT_LIST_NAME);
  const listWebUrl = normalizeText(process.env.SHAREPOINT_RECRUITMENT_LIST_WEB_URL || DEFAULT_LIST_WEB_URL);

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
    listWebUrl,
  };
}

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

async function resolveSiteId(graphClient, hostName, sitePath) {
  const url = `https://graph.microsoft.com/v1.0/sites/${hostName}:${sitePath}?$select=id`;
  const payload = await graphClient.fetchJson(url);
  if (!payload?.id) {
    throw new Error("Could not resolve SharePoint site id.");
  }
  return payload.id;
}

async function resolveList(graphClient, siteId, listName) {
  const params = new URLSearchParams({
    $select: "id,displayName,webUrl",
    $filter: `displayName eq ${quoteODataString(listName)}`,
    $top: "1",
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists?${params.toString()}`;
  const payload = await graphClient.fetchJson(url);
  const list = Array.isArray(payload?.value) ? payload.value[0] : null;
  if (!list?.id) {
    throw new Error(`Could not find SharePoint list '${listName}'.`);
  }
  return {
    id: String(list.id),
    webUrl: normalizeText(list.webUrl),
  };
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

function normalizeRecruitmentItem(item) {
  const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
  const active = toBoolean(fields.Active);
  const interviewWith = parsePersonField(fields.InterviewWith, fields.InterviewWithLookupId);

  return {
    id: normalizeText(item?.id),
    candidateName: normalizeText(fields.Title),
    location: normalizeText(fields.Location),
    phoneNumber: normalizeText(fields.PhoneNumber),
    interviewBooked: toBoolean(fields.InterviewBooked),
    interviewWith,
    status: normalizeText(fields.Status),
    active,
    keepInMind: toBoolean(fields.KeepinMind),
    livesIn: normalizeText(fields.LivesIn),
    firstInterviewDate: normalizeText(fields._x0031_stInterviewDate),
    notes: normalizeText(fields.Notes),
    source: normalizeText(fields.Source),
    earmarkedFor: normalizeText(fields.EarmarkedFor),
    oneTouchLink: parseHyperlink(fields.OnetouchLink),
    created: normalizeText(fields.Created || item?.createdDateTime),
  };
}

async function fetchRecruitmentItems(graphClient, siteId, listId) {
  const selectFields = [
    "Title",
    "Location",
    "PhoneNumber",
    "InterviewBooked",
    "InterviewWith",
    "InterviewWithLookupId",
    "Status",
    "Active",
    "KeepinMind",
    "LivesIn",
    "_x0031_stInterviewDate",
    "Notes",
    "Source",
    "EarmarkedFor",
    "OnetouchLink",
    "Created",
  ];

  const params = new URLSearchParams({
    $top: "200",
    $expand: `fields($select=${selectFields.join(",")})`,
  });

  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${params.toString()}`;
  const items = await graphClient.fetchAllPages(url);
  return items.map(normalizeRecruitmentItem).filter((item) => item.active);
}

function mapGraphError(error) {
  const status = Number(error?.status) || 502;
  const code = String(error?.code || "GRAPH_REQUEST_FAILED");
  const message = error?.message || "Recruitment request failed.";
  const retryable = Boolean(error?.retryable);
  return {
    status,
    payload: {
      error: {
        code,
        message,
        retryable,
      },
    },
  };
}

module.exports = async (req, res) => {
  if (req.method !== "GET") {
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

  try {
    const config = parseSiteConfig();
    const graphClient = createGraphDelegatedClient(req.authUser?.graphAccessToken);
    const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
    const list = await resolveList(graphClient, siteId, config.listName);
    const items = await fetchRecruitmentItems(graphClient, siteId, list.id);

    res.setHeader("Cache-Control", "private, max-age=30");
    res.status(200).json({
      listUrl: list.webUrl || config.listWebUrl,
      count: items.length,
      items,
    });
  } catch (error) {
    const mapped = mapGraphError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
