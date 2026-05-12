const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { createGraphDelegatedClient } = require("../_lib/tasks/graph-delegated-client");
const { RECRUITMENT_ALLOWED_ROLES } = require("../_lib/recruitment-access");

const DEFAULT_SITE_URL = "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency";
const DEFAULT_LIST_NAME = "Associate Recruitment";

const ALLOWED_ROLES = RECRUITMENT_ALLOWED_ROLES;

const STAGE_FIELD_MAP = {
  initial_screen: {
    outcome: "ScreenOutcome",
    nextSteps: "ScreenNextSteps",
    firstInterviewDate: "_x0031_stInterviewDate",
  },
  first_interview: {
    outcome: "FirstInterviewOutcome",
    nextSteps: "FirstInterviewNextSteps",
  },
  second_interview: {
    outcome: "SecondInterviewOutcome",
    nextSteps: "SecondInterviewNextSteps",
  },
};

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeDateOnly(value) {
  const normalized = normalizeText(value);
  return /^\d{4}-\d{2}-\d{2}$/.test(normalized) ? normalized : "";
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
  const stageKey = normalizeText(req.body?.stageKey).toLowerCase();
  const fieldMap = STAGE_FIELD_MAP[stageKey];
  if (!itemId || !fieldMap) {
    res.status(400).json({
      error: {
        code: "BAD_REQUEST",
        message: "Missing itemId or invalid stageKey.",
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

    await graphClient.fetchJson(url, {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        [fieldMap.outcome]: normalizeText(req.body?.outcome),
        [fieldMap.nextSteps]: normalizeText(req.body?.nextSteps),
        ...(fieldMap.firstInterviewDate
          ? {
              [fieldMap.firstInterviewDate]: normalizeDateOnly(req.body?.firstInterviewDate),
            }
          : {}),
      }),
    });

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      success: true,
      itemId,
      stageKey,
      outcome: normalizeText(req.body?.outcome),
      nextSteps: normalizeText(req.body?.nextSteps),
      firstInterviewDate: normalizeDateOnly(req.body?.firstInterviewDate),
    });
  } catch (error) {
    res.status(Number(error?.status) || 502).json({
      error: {
        code: String(error?.code || "STAGE_PATCH_FAILED"),
        message: error?.message || "Could not update recruitment stage.",
      },
    });
  }
};
