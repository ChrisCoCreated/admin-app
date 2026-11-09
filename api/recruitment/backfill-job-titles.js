const { requireGraphAuth } = require("../_lib/require-graph-auth");
const { createGraphDelegatedClient } = require("../_lib/tasks/graph-delegated-client");
const { RECRUITMENT_ALLOWED_ROLES } = require("../_lib/recruitment-access");

const DEFAULT_SITE_URL = "https://planwithcare.sharepoint.com/sites/OperationsSupportTeam_TE1079-RecruitmentandAgency";
const DEFAULT_LIST_NAME = "Associate Recruitment";
const BACKFILL_ADMIN_EMAIL = "chris@planwithcare.co.uk";

function normalizeText(value) {
  return String(value || "").trim();
}

function parseSiteConfig() {
  const siteUrl = new URL(normalizeText(process.env.SHAREPOINT_RECRUITMENT_SITE_URL || DEFAULT_SITE_URL));
  const listName = normalizeText(process.env.SHAREPOINT_RECRUITMENT_LIST_NAME || DEFAULT_LIST_NAME);
  const sitePath = siteUrl.pathname.replace(/\/$/, "");
  if (!sitePath || !listName) {
    throw new Error("Missing Recruitment SharePoint site or list configuration.");
  }
  return { hostName: siteUrl.hostname, sitePath, listName };
}

async function resolveListId(graphClient, config) {
  const site = await graphClient.fetchJson(
    `https://graph.microsoft.com/v1.0/sites/${config.hostName}:${config.sitePath}?$select=id`
  );
  if (!site?.id) {
    throw new Error("Could not resolve the Recruitment SharePoint site.");
  }

  const params = new URLSearchParams({
    $select: "id,displayName",
    $filter: `displayName eq '${config.listName.replace(/'/g, "''")}'`,
    $top: "1",
  });
  const payload = await graphClient.fetchJson(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/lists?${params.toString()}`
  );
  const list = Array.isArray(payload?.value) ? payload.value[0] : null;
  if (!list?.id) {
    throw new Error(`Could not find the Recruitment list '${config.listName}'.`);
  }
  return { siteId: site.id, listId: String(list.id) };
}

function extractAppliedForTitle(notes) {
  const match = String(notes || "").match(/(?:^|[\r\n]|[|;][ \t]*)Applied[ \t]+for(?:[ \t]*:[ \t]*|[ \t]+)([^\r\n|;]+)/i);
  return normalizeText(match?.[1]);
}

async function fetchAllRecruitmentItems(graphClient, siteId, listId) {
  const initialUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?$top=200&$expand=fields($select=JobTitle,Notes,Title)`;
  return graphClient.fetchAllPages(initialUrl);
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({ error: { code: "METHOD_NOT_ALLOWED", message: "Method Not Allowed" } });
    return;
  }

  if (!(await requireGraphAuth(req, res, { allowedRoles: RECRUITMENT_ALLOWED_ROLES }))) {
    return;
  }
  if (normalizeText(req.authUser?.email).toLowerCase() !== BACKFILL_ADMIN_EMAIL) {
    res.status(403).json({ error: { code: "FORBIDDEN", message: "Only the configured app owner can run this one-time backfill." } });
    return;
  }

  try {
    const config = parseSiteConfig();
    const graphClient = createGraphDelegatedClient(req.authUser?.graphAccessToken);
    const { siteId, listId } = await resolveListId(graphClient, config);
    const items = await fetchAllRecruitmentItems(graphClient, siteId, listId);
    const result = { scanned: items.length, updated: 0, skippedWithTitle: 0, skippedWithoutAppliedFor: 0, failed: 0 };

    for (const item of items) {
      const fields = item?.fields && typeof item.fields === "object" ? item.fields : {};
      if (normalizeText(fields.JobTitle)) {
        result.skippedWithTitle += 1;
        continue;
      }

      const jobTitle = extractAppliedForTitle(fields.Notes);
      if (!jobTitle) {
        result.skippedWithoutAppliedFor += 1;
        continue;
      }

      try {
        await graphClient.fetchJson(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items/${encodeURIComponent(item.id)}/fields`,
          {
            method: "PATCH",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ JobTitle: jobTitle }),
          }
        );
        result.updated += 1;
      } catch (error) {
        result.failed += 1;
        console.warn("[recruitment-job-title-backfill] Could not update a candidate", {
          itemId: normalizeText(item?.id),
          message: error?.message || String(error),
        });
      }
    }

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({ success: true, ...result });
  } catch (error) {
    res.status(Number(error?.status) || 502).json({
      error: {
        code: String(error?.code || "JOB_TITLE_BACKFILL_FAILED"),
        message: error?.message || "Could not backfill Recruitment job titles.",
      },
    });
  }
};
