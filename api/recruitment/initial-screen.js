const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { createGraphDelegatedClient } = require("../_lib/tasks/graph-delegated-client");

const DEFAULT_SITE_URL = "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency";
const DEFAULT_LIST_NAME = "Associate Recruitment";

const ALLOWED_ROLES = [
  "admin",
  "care_manager",
  "operations",
  "hr_only",
  "hr_clients",
  "time_hr",
  "time_hr_clients",
];

const SCREEN_FIELDS = [
  "Title",
  "Status",
  "Location",
  "PhoneNumber",
  "Active",
  "Q1_Notes_Availability",
  "Q1_Score",
  "Q2_Notes_ShortNotice",
  "Q2_Score",
  "Q3_Notes_Travel",
  "Q3_Score",
  "Q4_Notes_ValuesFit",
  "Q4_Score",
  "Q5_Notes_GoodCare",
  "Q5_Score",
  "Q6_Notes_Flexibility",
  "Q6_Score",
  "Q7_Notes_Wellbeing",
  "Q7_Score",
  "InitialCallSummary",
];

function normalizeText(value) {
  return String(value || "").trim();
}

function quoteODataString(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

function toBoolean(value) {
  if (value === true || value === false) {
    return value;
  }
  const normalized = String(value || "").trim().toLowerCase();
  if (!normalized) {
    return false;
  }
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "y";
}

function normalizeScore(value) {
  const normalized = normalizeText(value).toLowerCase();
  if (normalized === "green") {
    return "Green";
  }
  if (normalized === "amber") {
    return "Amber";
  }
  if (normalized === "red") {
    return "Red";
  }
  return "";
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

function mapInitialScreenItem(item) {
  const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
  return {
    itemId: normalizeText(item?.id),
    candidateName: normalizeText(fields.Title),
    status: normalizeText(fields.Status),
    location: normalizeText(fields.Location),
    phoneNumber: normalizeText(fields.PhoneNumber),
    active: toBoolean(fields.Active),
    responses: {
      q1NotesAvailability: normalizeText(fields.Q1_Notes_Availability),
      q1Score: normalizeScore(fields.Q1_Score),
      q2NotesShortNotice: normalizeText(fields.Q2_Notes_ShortNotice),
      q2Score: normalizeScore(fields.Q2_Score),
      q3NotesTravel: normalizeText(fields.Q3_Notes_Travel),
      q3Score: normalizeScore(fields.Q3_Score),
      q4NotesValuesFit: normalizeText(fields.Q4_Notes_ValuesFit),
      q4Score: normalizeScore(fields.Q4_Score),
      q5NotesGoodCare: normalizeText(fields.Q5_Notes_GoodCare),
      q5Score: normalizeScore(fields.Q5_Score),
      q6NotesFlexibility: normalizeText(fields.Q6_Notes_Flexibility),
      q6Score: normalizeScore(fields.Q6_Score),
      q7NotesWellbeing: normalizeText(fields.Q7_Notes_Wellbeing),
      q7Score: normalizeScore(fields.Q7_Score),
      initialCallSummary: normalizeText(fields.InitialCallSummary),
    },
  };
}

function buildPatchBody(input = {}) {
  return {
    Q1_Notes_Availability: normalizeText(input.q1NotesAvailability),
    Q1_Score: normalizeScore(input.q1Score),
    Q2_Notes_ShortNotice: normalizeText(input.q2NotesShortNotice),
    Q2_Score: normalizeScore(input.q2Score),
    Q3_Notes_Travel: normalizeText(input.q3NotesTravel),
    Q3_Score: normalizeScore(input.q3Score),
    Q4_Notes_ValuesFit: normalizeText(input.q4NotesValuesFit),
    Q4_Score: normalizeScore(input.q4Score),
    Q5_Notes_GoodCare: normalizeText(input.q5NotesGoodCare),
    Q5_Score: normalizeScore(input.q5Score),
    Q6_Notes_Flexibility: normalizeText(input.q6NotesFlexibility),
    Q6_Score: normalizeScore(input.q6Score),
    Q7_Notes_Wellbeing: normalizeText(input.q7NotesWellbeing),
    Q7_Score: normalizeScore(input.q7Score),
    InitialCallSummary: normalizeText(input.initialCallSummary),
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

  const itemId = normalizeText(req.method === "GET" ? req.query?.itemId : req.body?.itemId);
  if (!itemId) {
    res.status(400).json({
      error: {
        code: "BAD_REQUEST",
        message: "Missing itemId.",
      },
    });
    return;
  }

  try {
    const config = parseSiteConfig();
    const graphClient = createGraphDelegatedClient(req.authUser?.graphAccessToken);
    const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
    const listId = await resolveListId(graphClient, siteId, config.listName);
    const itemUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}?$expand=fields($select=${SCREEN_FIELDS.join(",")})`;

    if (req.method === "POST") {
      const currentItem = mapInitialScreenItem(await graphClient.fetchJson(itemUrl));
      if (!currentItem.active) {
        res.status(409).json({
          error: {
            code: "INACTIVE_CANDIDATE",
            message: "Only active candidates can be screened.",
          },
        });
        return;
      }

      const patchUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(itemId)}/fields`;
      await graphClient.fetchJson(patchUrl, {
        method: "PATCH",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(buildPatchBody(req.body?.responses)),
      });

      const updatedItem = mapInitialScreenItem(await graphClient.fetchJson(itemUrl));
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json({
        success: true,
        item: updatedItem,
      });
      return;
    }

    const item = mapInitialScreenItem(await graphClient.fetchJson(itemUrl));
    if (!item.active) {
      res.status(404).json({
        error: {
          code: "NOT_FOUND",
          message: "Candidate is not active.",
        },
      });
      return;
    }

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      item,
    });
  } catch (error) {
    res.status(Number(error?.status) || 502).json({
      error: {
        code: String(error?.code || "INITIAL_SCREEN_FAILED"),
        message: error?.message || "Could not load or save initial screen data.",
      },
    });
  }
};
