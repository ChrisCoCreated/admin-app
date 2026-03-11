const { requireGraphAuth } = require("./_lib/require-graph-auth");
const { createGraphDelegatedClient } = require("./_lib/tasks/graph-delegated-client");
const { createCarer } = require("./_lib/onetouch-client");

const DEFAULT_SITE_URL = "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency";
const DEFAULT_LIST_NAME = "Associate Recruitment";
const DEFAULT_LIST_WEB_URL =
  "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency/Lists/Associate%20Recruitment/Active.aspx?env=WebViewList";
const ONETOUCH_CARER_PROFILE_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carer/carerSummaryProfile.php";

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

function buildOneTouchProfileUrl(oneTouchId) {
  const url = new URL(ONETOUCH_CARER_PROFILE_BASE_URL);
  url.searchParams.set("p", normalizeText(oneTouchId));
  return url.toString();
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
    "Email",
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

async function fetchRecruitmentItem(graphClient, siteId, listId, itemId) {
  const selectFields = [
    "Title",
    "Location",
    "Email",
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
    $expand: `fields($select=${selectFields.join(",")})`,
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}?${params.toString()}`;
  const payload = await graphClient.fetchJson(url);
  return normalizeRecruitmentItem(payload);
}

async function patchRecruitmentOneTouchLink(graphClient, siteId, listId, itemId, oneTouchProfileUrl) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}/fields`;
  await graphClient.fetchJson(url, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      OnetouchLink: normalizeText(oneTouchProfileUrl),
    }),
  });
}

function buildOneTouchCreatePayload(candidate, overrides = {}) {
  return {
    external_id: normalizeText(candidate.id),
    full_name: normalizeText(candidate.candidateName),
    phone: normalizeText(candidate.phoneNumber),
    livesIn: normalizeText(candidate.livesIn),
    location: normalizeText(candidate.location),
    area: normalizeText(overrides.area || candidate.earmarkedFor),
    recruitment_source: normalizeText(overrides.recruitmentSource || candidate.source),
    source: normalizeText(overrides.recruitmentSource || candidate.source),
    position: normalizeText(overrides.position || "Carer"),
    status: normalizeText(overrides.status),
    notes: normalizeText(candidate.notes),
  };
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
  if (req.method !== "GET" && req.method !== "POST") {
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
    if (req.method === "POST") {
      const itemId = normalizeText(req.body?.itemId);
      if (!itemId) {
        res.status(400).json({
          error: {
            code: "BAD_REQUEST",
            message: "Missing itemId.",
          },
        });
        return;
      }

      const candidate = await fetchRecruitmentItem(graphClient, siteId, list.id, itemId);
      if (!candidate.id) {
        res.status(404).json({
          error: {
            code: "NOT_FOUND",
            message: "Recruitment candidate not found.",
          },
        });
        return;
      }
      if (!candidate.active) {
        res.status(409).json({
          error: {
            code: "INACTIVE_CANDIDATE",
            message: "Only active candidates can be added to OneTouch.",
          },
        });
        return;
      }
      if (normalizeText(candidate.oneTouchLink)) {
        res.status(409).json({
          error: {
            code: "ALREADY_LINKED",
            message: "Candidate already has a OneTouch link.",
          },
        });
        return;
      }

      const createResult = await createCarer(
        buildOneTouchCreatePayload(candidate, {
          area: req.body?.area,
          recruitmentSource: req.body?.recruitmentSource,
          position: req.body?.position,
          status: req.body?.status,
        })
      );
      const oneTouchProfileUrl = buildOneTouchProfileUrl(createResult.id);
      await patchRecruitmentOneTouchLink(graphClient, siteId, list.id, itemId, oneTouchProfileUrl);
      const updatedCandidate = await fetchRecruitmentItem(graphClient, siteId, list.id, itemId);

      res.setHeader("Cache-Control", "no-store");
      res.status(200).json({
        success: true,
        itemId,
        oneTouchId: createResult.id,
        oneTouchLink: oneTouchProfileUrl,
        item: updatedCandidate,
      });
      return;
    }

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
