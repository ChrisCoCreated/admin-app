const { requireGraphAuth } = require("./_lib/require-graph-auth");
const { createGraphDelegatedClient } = require("./_lib/tasks/graph-delegated-client");
const { createCarer } = require("./_lib/onetouch-client");
const { RECRUITMENT_ALLOWED_ROLES } = require("./_lib/recruitment-access");
const { patchSharePointUrlField } = require("./_lib/sharepoint-url-field");
const { createSharePointColleague } = require("./_lib/colleagues-source");

const DEFAULT_SITE_URL = "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency";
const DEFAULT_LIST_NAME = "Associate Recruitment";
const DEFAULT_LIST_WEB_URL =
  "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency/Lists/Associate%20Recruitment/Active.aspx?env=WebViewList";
const ONETOUCH_CARER_PROFILE_BASE_URL = "https://care2.onetouchhealth.net/cm/in/carer/carerSummaryProfile.php";
const DEFAULT_EXTERNAL_ID_PREFIX = "thrive-recruitment";

const ALLOWED_ROLES = RECRUITMENT_ALLOWED_ROLES;

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

function parseSharePointListConfig(siteUrlValue, listName, missingConfigMessage) {
  if (!siteUrlValue || !listName) {
    throw new Error(missingConfigMessage);
  }

  const siteUrl = new URL(siteUrlValue);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath) {
    throw new Error("Configured SharePoint site URL must include a site path.");
  }

  return {
    hostName: siteUrl.hostname,
    siteBaseUrl: `${siteUrl.protocol}//${siteUrl.host}${sitePath}`,
    sitePath,
    listName,
  };
}

function parseSiteConfig() {
  const siteUrlValue = normalizeText(process.env.SHAREPOINT_RECRUITMENT_SITE_URL || DEFAULT_SITE_URL);
  const listName = normalizeText(process.env.SHAREPOINT_RECRUITMENT_LIST_NAME || DEFAULT_LIST_NAME);
  const listWebUrl = normalizeText(process.env.SHAREPOINT_RECRUITMENT_LIST_WEB_URL || DEFAULT_LIST_WEB_URL);
  return {
    ...parseSharePointListConfig(
      siteUrlValue,
      listName,
      "Missing SHAREPOINT_RECRUITMENT_SITE_URL or SHAREPOINT_RECRUITMENT_LIST_NAME."
    ),
    listWebUrl,
  };
}

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function normalizeToken(value) {
  return normalizeText(value).replace(/[^a-z0-9]/g, "");
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

async function resolveListColumns(graphClient, siteId, listId, select = "name,displayName") {
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns?$select=${select}&$top=200`;
  const columns = [];
  while (nextUrl) {
    const payload = await graphClient.fetchJson(nextUrl);
    const values = Array.isArray(payload?.value) ? payload.value : [];
    columns.push(...values);
    nextUrl = String(payload?.["@odata.nextLink"] || "");
  }
  return columns;
}

function findColumnInternalName(columns, candidates, fallback = "") {
  const byNameToken = new Map();
  const byDisplayToken = new Map();

  for (const column of Array.isArray(columns) ? columns : []) {
    const name = normalizeText(column?.name);
    const displayName = normalizeText(column?.displayName);
    if (!name) {
      continue;
    }
    const nameToken = normalizeToken(name);
    const displayToken = normalizeToken(displayName);
    if (nameToken && !byNameToken.has(nameToken)) {
      byNameToken.set(nameToken, name);
    }
    if (displayToken && !byDisplayToken.has(displayToken)) {
      byDisplayToken.set(displayToken, name);
    }
  }

  for (const candidate of candidates) {
    const token = normalizeToken(candidate);
    const match = byDisplayToken.get(token) || byNameToken.get(token);
    if (match) {
      return match;
    }
  }

  return fallback;
}

async function resolveOneTouchLinkFieldName(graphClient, siteId, listId) {
  const columns = await resolveListColumns(graphClient, siteId, listId);

  const candidates = [];
  for (const column of columns) {
    const name = normalizeText(column?.name);
    const displayName = normalizeText(column?.displayName);
    const nameToken = normalizeToken(name);
    const displayToken = normalizeToken(displayName);
    const isOneTouchLink =
      nameToken.includes("onetouchlink") ||
      displayToken.includes("onetouchlink") ||
      (displayToken.includes("onetouch") && displayToken.includes("link"));
    if (isOneTouchLink && name) {
      candidates.push(name);
    }
  }

  return candidates[0] || "OnetouchLink";
}

async function resolveChoiceColumnOptions(graphClient, siteId, listId) {
  let nextUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/columns?$select=name,displayName,choice&$top=200`;
  const options = {
    status: [],
    currentOwner: [],
  };

  while (nextUrl) {
    const payload = await graphClient.fetchJson(nextUrl);
    const values = Array.isArray(payload?.value) ? payload.value : [];
    for (const column of values) {
      const internalName = normalizeText(column?.name);
      const displayName = normalizeText(column?.displayName).toLowerCase();
      const choices = Array.isArray(column?.choice?.choices) ? column.choice.choices.map(normalizeText).filter(Boolean) : [];
      if (!choices.length) {
        continue;
      }
      if (internalName === "Status" || displayName === "status") {
        options.status = choices;
      }
      if (internalName === "Current_x0020_Owner" || displayName === "current owner") {
        options.currentOwner = choices;
      }
    }
    nextUrl = String(payload?.["@odata.nextLink"] || "");
  }

  return options;
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

function buildOneTouchProfileUrl(oneTouchId) {
  const url = new URL(ONETOUCH_CARER_PROFILE_BASE_URL);
  url.searchParams.set("p", normalizeText(oneTouchId));
  return url.toString();
}

function buildRecruitmentExternalId(candidateId) {
  const id = normalizeText(candidateId);
  if (!id) {
    return "";
  }
  const configuredPrefix = normalizeText(
    process.env.ONETOUCH_RECRUITMENT_EXTERNAL_ID_PREFIX || DEFAULT_EXTERNAL_ID_PREFIX
  );
  if (!configuredPrefix) {
    return id;
  }
  const prefixWithDash = `${configuredPrefix}-`;
  if (id.toLowerCase().startsWith(prefixWithDash.toLowerCase())) {
    return id;
  }
  return `${configuredPrefix}-${id}`;
}

function normalizeRecruitmentItem(item) {
  const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
  const active = toBoolean(fields.Active);
  const interviewWith = parsePersonField(fields.InterviewWith, fields.InterviewWithLookupId);
  const currentOwnerRaw = normalizeText(fields.Current_x0020_Owner);

  return {
    id: normalizeText(item?.id),
    candidateName: normalizeText(fields.Title),
    location: normalizeText(fields.Location),
    email: normalizeText(fields.Email),
    phoneNumber: normalizeText(fields.PhoneNumber),
    interviewBooked: toBoolean(fields.InterviewBooked),
    interviewWith,
    status: normalizeText(fields.Status),
    currentOwner: currentOwnerRaw,
    active,
    keepInMind: toBoolean(fields.KeepinMind),
    liveInMailingList: toBoolean(fields.Live_x002d_inmailinglist),
    livesIn: normalizeText(fields.LivesIn),
    firstInterviewDate: normalizeText(fields._x0031_stInterviewDate),
    notes: normalizeText(fields.Notes),
    indeedProfileUrl: normalizeText(fields.IndeedURL) || extractIndeedProfileUrl(fields.Notes),
    screenOutcome: normalizeText(fields.ScreenOutcome),
    screenNextSteps: normalizeText(fields.ScreenNextSteps),
    firstInterviewOutcome: normalizeText(fields.FirstInterviewOutcome),
    firstInterviewNextSteps: normalizeText(fields.FirstInterviewNextSteps),
    secondInterviewOutcome: normalizeText(fields.SecondInterviewOutcome),
    secondInterviewNextSteps: normalizeText(fields.SecondInterviewNextSteps),
    tags: normalizeText(fields.Tags),
    source: normalizeText(fields.Source),
    earmarkedFor: normalizeText(fields.EarmarkedFor),
    oneTouchLink: parseHyperlink(fields.OnetouchLink),
    created: normalizeText(fields.Created || item?.createdDateTime),
    updated: normalizeText(fields.Modified || item?.lastModifiedDateTime || fields.Created || item?.createdDateTime),
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
    "Current_x0020_Owner",
    "Active",
    "KeepinMind",
    "Live_x002d_inmailinglist",
    "LivesIn",
    "_x0031_stInterviewDate",
    "Notes",
    "IndeedURL",
    "ScreenOutcome",
    "ScreenNextSteps",
    "FirstInterviewOutcome",
    "FirstInterviewNextSteps",
    "SecondInterviewOutcome",
    "SecondInterviewNextSteps",
    "Tags",
    "Source",
    "EarmarkedFor",
    "OnetouchLink",
    "Created",
    "Modified",
  ];

  const params = new URLSearchParams({
    $top: "200",
    $expand: `fields($select=${selectFields.join(",")})`,
  });

  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${params.toString()}`;
  const items = await graphClient.fetchAllPages(url);
  return items.map(normalizeRecruitmentItem);
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
    "Current_x0020_Owner",
    "Active",
    "KeepinMind",
    "Live_x002d_inmailinglist",
    "LivesIn",
    "_x0031_stInterviewDate",
    "Notes",
    "IndeedURL",
    "ScreenOutcome",
    "ScreenNextSteps",
    "FirstInterviewOutcome",
    "FirstInterviewNextSteps",
    "SecondInterviewOutcome",
    "SecondInterviewNextSteps",
    "Tags",
    "Source",
    "EarmarkedFor",
    "OnetouchLink",
    "Created",
    "Modified",
  ];
  const params = new URLSearchParams({
    $expand: `fields($select=${selectFields.join(",")})`,
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}?${params.toString()}`;
  const payload = await graphClient.fetchJson(url);
  return normalizeRecruitmentItem(payload);
}

async function patchRecruitmentOneTouchLink(graphClient, config, siteId, listId, itemId, oneTouchProfileUrl, incomingToken) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}/fields`;
  const fieldName = await resolveOneTouchLinkFieldName(graphClient, siteId, listId);
  const link = normalizeText(oneTouchProfileUrl);
  const attempts = [
    { label: "plain_url", body: { [fieldName]: link } },
    { label: "url_plus_description", body: { [fieldName]: `${link}, OneTouch` } },
    { label: "url_object", body: { [fieldName]: { Url: link, Description: "OneTouch" } } },
  ];

  let lastError = null;
  for (const attempt of attempts) {
    try {
      await graphClient.fetchJson(url, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(attempt.body),
      });
      return;
    } catch (error) {
      lastError = error;
      console.warn("[recruitment] SharePoint OnetouchLink patch attempt failed", {
        itemId,
        fieldName,
        attempt: attempt.label,
        message: error?.message || String(error),
      });
    }
  }

  await patchSharePointUrlField({
    incomingToken,
    siteBaseUrl: config.siteBaseUrl,
    hostName: config.hostName,
    listName: config.listName,
    itemId,
    fieldInternalName: fieldName,
    urlValue: link,
    description: "OneTouch",
  });
}

function appendWarning(existingWarning, nextWarning) {
  const current = normalizeText(existingWarning);
  const next = normalizeText(nextWarning);
  if (!next) {
    return current;
  }
  if (!current) {
    return next;
  }
  return `${current} ${next}`;
}

function buildOneTouchCreatePayload(candidate, overrides = {}) {
  return {
    external_id: buildRecruitmentExternalId(candidate.id),
    full_name: normalizeText(candidate.candidateName),
    primary_email: normalizeText(candidate.email),
    phone: normalizeText(candidate.phoneNumber),
    livesIn: normalizeText(candidate.livesIn),
    location: normalizeText(candidate.location),
    area: normalizeText(overrides.area || candidate.earmarkedFor),
    recruitment_source: normalizeText(overrides.recruitmentSource || candidate.source),
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
      let sharePointLinkPatched = false;
      let colleagueCreated = false;
      let warning = "";
      try {
        await patchRecruitmentOneTouchLink(
          graphClient,
          config,
          siteId,
          list.id,
          itemId,
          oneTouchProfileUrl,
          req.authUser?.graphAccessToken
        );
        sharePointLinkPatched = true;
      } catch (error) {
        warning = error?.message || "Created in OneTouch, but could not update the recruitment list link.";
        console.warn("[recruitment] OneTouch create succeeded but SharePoint link patch failed", {
          itemId,
          oneTouchId: createResult.id,
          message: warning,
        });
      }
      try {
        await createSharePointColleague({
          name: candidate.candidateName,
          oneTouchId: createResult.id,
          archived: false,
        });
        colleagueCreated = true;
      } catch (error) {
        const colleaguesWarning =
          error?.message || "Created in OneTouch, but could not add the colleague to the Colleagues list.";
        warning = appendWarning(warning, colleaguesWarning);
        console.warn("[recruitment] OneTouch create succeeded but Colleagues list create failed", {
          itemId,
          oneTouchId: createResult.id,
          message: colleaguesWarning,
        });
      }

      res.setHeader("Cache-Control", "no-store");
      res.status(200).json({
        success: true,
        itemId,
        oneTouchId: createResult.id,
        oneTouchLink: oneTouchProfileUrl,
        sharePointLinkPatched,
        colleagueCreated,
        warning,
      });
      return;
    }

    const items = await fetchRecruitmentItems(graphClient, siteId, list.id);
    const choiceOptions = await resolveChoiceColumnOptions(graphClient, siteId, list.id);

    res.setHeader("Cache-Control", "private, max-age=30");
    res.status(200).json({
      listUrl: list.webUrl || config.listWebUrl,
      count: items.length,
      choiceOptions,
      items,
    });
  } catch (error) {
    console.error("[recruitment] Add to OneTouch failed", {
      itemId: normalizeText(req?.body?.itemId),
      area: normalizeText(req?.body?.area),
      recruitmentSource: normalizeText(req?.body?.recruitmentSource),
      position: normalizeText(req?.body?.position),
      statusInput: normalizeText(req?.body?.status),
      message: error?.message || String(error),
    });
    const mapped = mapGraphError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
