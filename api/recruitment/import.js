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

const STATUS_PIPELINE_MAP = [
  { match: /\b(contacting|applied|application|new)\b/i, status: "Initial Call" },
  { match: /\b(interview|screening|screen)\b/i, status: "1st Interview" },
  { match: /\boffer\b/i, status: "Offered" },
  { match: /\b(hired|accepted)\b/i, status: "Accepted" },
  { match: /\brejected\b/i, status: "Rejected" },
  { match: /\blost\b/i, status: "Lost" },
];

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeKey(value) {
  return String(value || "")
    .trim()
    .toLowerCase();
}

function normalizeEmail(value) {
  return normalizeText(value).toLowerCase();
}

function normalizePhone(value) {
  return normalizeText(value).replace(/[^\d+]/g, "");
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
  return { id: String(list.id) };
}

function toBoolean(value) {
  if (value === true || value === false) {
    return value;
  }
  const normalized = normalizeText(value).toLowerCase();
  if (!normalized) {
    return false;
  }
  return normalized === "true" || normalized === "1" || normalized === "yes" || normalized === "y";
}

function getRowValue(row, key) {
  const needle = normalizeKey(key);
  for (const [rawKey, value] of Object.entries(row || {})) {
    if (normalizeKey(rawKey) === needle) {
      return normalizeText(value);
    }
  }
  return "";
}

function stripTrailingUkPostcode(input) {
  const value = normalizeText(input);
  if (!value) {
    return "";
  }

  const withSpaces = value.replace(/([A-Za-z])(\d)/g, "$1 $2").replace(/(\d)([A-Za-z])/g, "$1 $2");
  const normalized = withSpaces.replace(/\s+/g, " ").trim();
  const fullPostcode = /^(.*?)(?:[\s,;-]+)?([A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})$/i.exec(normalized);
  if (fullPostcode) {
    const base = normalizeText(fullPostcode[1]);
    return base || normalized;
  }

  const outwardOnly = /^(.*?)(?:[\s,;-]+)?([A-Z]{1,2}\d[A-Z\d]?)$/i.exec(normalized);
  if (outwardOnly) {
    const base = normalizeText(outwardOnly[1]);
    return base || normalized;
  }

  return value;
}

function normalizePipelineStatus(statusValue, interestLevelValue) {
  const statusRaw = normalizeText(statusValue);
  const interestRaw = normalizeText(interestLevelValue);
  const candidate = `${statusRaw} ${interestRaw}`.trim();
  if (!candidate) {
    return "Initial Call";
  }

  for (const rule of STATUS_PIPELINE_MAP) {
    if (rule.match.test(candidate)) {
      return rule.status;
    }
  }

  return statusRaw || "Initial Call";
}

function buildNotes(row) {
  const sections = [
    ["Relevant Experience", getRowValue(row, "relevant experience")],
    ["Education", getRowValue(row, "education")],
    ["Job Title", getRowValue(row, "job title")],
    ["Date", getRowValue(row, "date")],
    ["Interest Level", getRowValue(row, "interest level")],
  ].filter((entry) => entry[1]);

  return sections.map((entry) => `${entry[0]}: ${entry[1]}`).join("\n");
}

function mapRowToSharePointFields(row) {
  const candidateName = getRowValue(row, "name");
  if (!candidateName) {
    return { error: "Missing candidate name." };
  }

  const phoneNumber = getRowValue(row, "phone");
  const email = getRowValue(row, "email");
  const source = getRowValue(row, "source") || "Indeed";
  const livesIn = getRowValue(row, "candidate location");
  const locationRaw = getRowValue(row, "job location");
  const location = stripTrailingUkPostcode(locationRaw);
  const status = normalizePipelineStatus(getRowValue(row, "status"), getRowValue(row, "interest level"));
  const notes = buildNotes(row);

  return {
    fields: {
      Title: candidateName,
      PhoneNumber: phoneNumber,
      Email: email,
      LivesIn: livesIn,
      Location: location,
      Source: source,
      Status: status,
      Notes: notes,
      Active: true,
    },
    candidateName,
    phoneNumber,
    email,
  };
}

async function fetchActiveRecruitmentItems(graphClient, siteId, listId) {
  const selectFields = ["Title", "PhoneNumber", "Email", "Active"];
  const params = new URLSearchParams({
    $top: "200",
    $expand: `fields($select=${selectFields.join(",")})`,
  });
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items?${params.toString()}`;
  const items = await graphClient.fetchAllPages(url);
  return items.filter((item) => toBoolean(item?.fields?.Active));
}

async function createRecruitmentItem(graphClient, siteId, listId, fields) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listId}/items`;
  return graphClient.fetchJson(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ fields }),
  });
}

function mapGraphError(error) {
  const status = Number(error?.status) || 502;
  const code = String(error?.code || "GRAPH_REQUEST_FAILED");
  const message = error?.message || "Recruitment import request failed.";
  return {
    status,
    payload: {
      error: {
        code,
        message,
        retryable: Boolean(error?.retryable),
      },
    },
  };
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

  try {
    const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];
    const dryRun = req.body?.dryRun === true;
    if (!rows.length) {
      res.status(400).json({
        error: {
          code: "BAD_REQUEST",
          message: "No CSV rows provided.",
        },
      });
      return;
    }

    const hasNameHeader = Object.keys(rows[0] || {}).some((key) => normalizeKey(key) === "name");
    if (!hasNameHeader) {
      res.status(400).json({
        error: {
          code: "INVALID_HEADERS",
          message: "Missing required CSV header: name.",
        },
      });
      return;
    }

    const config = parseSiteConfig();
    const graphClient = createGraphDelegatedClient(req.authUser?.graphAccessToken);
    const siteId = await resolveSiteId(graphClient, config.hostName, config.sitePath);
    const list = await resolveList(graphClient, siteId, config.listName);
    const existingItems = await fetchActiveRecruitmentItems(graphClient, siteId, list.id);

    const knownEmails = new Set();
    const knownPhones = new Set();
    for (const item of existingItems) {
      const fields = item?.fields || {};
      const email = normalizeEmail(fields.Email);
      const phone = normalizePhone(fields.PhoneNumber);
      if (email) {
        knownEmails.add(email);
      }
      if (phone) {
        knownPhones.add(phone);
      }
    }

    let rejected = 0;
    let skippedDuplicates = 0;
    let created = 0;
    const errors = [];

    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index];
      const mapped = mapRowToSharePointFields(row);
      if (mapped.error) {
        rejected += 1;
        errors.push({ row: index + 2, message: mapped.error });
        continue;
      }

      const emailKey = normalizeEmail(mapped.email);
      const phoneKey = normalizePhone(mapped.phoneNumber);
      const isDuplicate =
        (emailKey && knownEmails.has(emailKey)) ||
        (phoneKey && knownPhones.has(phoneKey));
      if (isDuplicate) {
        skippedDuplicates += 1;
        continue;
      }

      if (!dryRun) {
        await createRecruitmentItem(graphClient, siteId, list.id, mapped.fields);
        created += 1;
      }

      if (emailKey) {
        knownEmails.add(emailKey);
      }
      if (phoneKey) {
        knownPhones.add(phoneKey);
      }
    }

    const wouldInsert = rows.length - rejected - skippedDuplicates;

    res.setHeader("Cache-Control", "no-store");
    res.status(200).json({
      dryRun,
      totalRows: rows.length,
      wouldInsert,
      inserted: dryRun ? 0 : created,
      skippedDuplicates,
      rejected,
      errors,
    });
  } catch (error) {
    const mapped = mapGraphError(error);
    res.status(mapped.status).json(mapped.payload);
  }
};
