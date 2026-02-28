const crypto = require("crypto");
const { listClients: listOneTouchClients } = require("./onetouch-client");
const {
  createSharePointClient,
  listSharePointClientsWithItemIds,
  patchSharePointClient,
} = require("./clients-reconcile-repository");

const CORE_FIELDS = [
  "name",
  "dateOfBirth",
  "oneTouchId",
  "postcode",
  "email",
  "phone",
  "status",
  "address",
  "town",
  "county",
];

function normalizeText(value) {
  return String(value || "").trim();
}

function collectNonEmptyValues(values) {
  return values.map((value) => normalizeText(value)).filter(Boolean);
}

function dedupeEmails(values) {
  const seen = new Set();
  const output = [];
  for (const value of values) {
    const key = normalizeText(value).toLowerCase();
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    output.push(normalizeText(value));
  }
  return output;
}

function normalizePhoneForCompare(value) {
  return normalizeText(value).replace(/[^\d+]/g, "").toLowerCase();
}

function dedupePhones(values) {
  const seen = new Set();
  const output = [];
  for (const value of values) {
    const raw = normalizeText(value);
    const key = normalizePhoneForCompare(raw);
    if (!raw || !key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    output.push(raw);
  }
  return output;
}

function buildCombinedEmail(oneTouchClient) {
  const values = collectNonEmptyValues([
    oneTouchClient?.emailPrimary,
    oneTouchClient?.emailSecondary,
    oneTouchClient?.email,
    oneTouchClient?.raw?.primary_email,
    oneTouchClient?.raw?.secondary_email,
    oneTouchClient?.raw?.email,
    oneTouchClient?.raw?.email_address,
  ]);
  return dedupeEmails(values).join("; ");
}

function buildCombinedPhone(oneTouchClient) {
  const values = collectNonEmptyValues([
    oneTouchClient?.phoneMobile,
    oneTouchClient?.phoneHome,
    oneTouchClient?.phone,
    oneTouchClient?.raw?.phone_mobile,
    oneTouchClient?.raw?.phone_home,
    oneTouchClient?.raw?.phone,
    oneTouchClient?.raw?.mobile,
  ]);
  return dedupePhones(values).join(" / ");
}

function normalizeName(value) {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[^\w\s]/g, " ")
    .replace(/\s+/g, " ");
}

function nameTokens(value) {
  return normalizeName(value)
    .split(" ")
    .map((item) => item.trim())
    .filter(Boolean);
}

function fuzzyNameMatch(left, right) {
  const a = normalizeName(left);
  const b = normalizeName(right);
  if (!a || !b) {
    return false;
  }
  if (a === b) {
    return true;
  }
  if (a.includes(b) || b.includes(a)) {
    return true;
  }

  const aTokens = new Set(nameTokens(a));
  const bTokens = new Set(nameTokens(b));
  if (!aTokens.size || !bTokens.size) {
    return false;
  }

  let shared = 0;
  for (const token of aTokens) {
    if (bTokens.has(token)) {
      shared += 1;
    }
  }

  const ratio = shared / Math.max(aTokens.size, bTokens.size);
  return ratio >= 0.67;
}

function normalizeId(value) {
  return normalizeText(value).toLowerCase();
}

function parseDateParts(value) {
  const raw = normalizeText(value);
  if (!raw) {
    return null;
  }

  const datePart = raw.includes("T") ? raw.split("T")[0] : raw;
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(datePart);
  if (match) {
    return {
      year: Number(match[1]),
      month: Number(match[2]),
      day: Number(match[3]),
    };
  }

  match = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/.exec(datePart);
  if (match) {
    const yearRaw = Number(match[3]);
    return {
      year: yearRaw < 100 ? 2000 + yearRaw : yearRaw,
      month: Number(match[2]),
      day: Number(match[1]),
    };
  }

  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) {
    return {
      year: parsed.getUTCFullYear(),
      month: parsed.getUTCMonth() + 1,
      day: parsed.getUTCDate(),
    };
  }

  return null;
}

function datesMatch(left, right) {
  const a = parseDateParts(left);
  const b = parseDateParts(right);
  if (!a || !b) {
    return false;
  }
  return a.year === b.year && a.month === b.month && a.day === b.day;
}

function findNameCandidates(sharePointClient, oneTouchClients) {
  const sharePointName = normalizeName(sharePointClient?.name);
  if (!sharePointName) {
    return { candidates: [], matchType: "none" };
  }

  const validClients = oneTouchClients.filter((client) => {
    if (!normalizeId(client?.id)) {
      return false;
    }
    return Boolean(normalizeName(client?.name));
  });

  const exact = validClients.filter((client) => normalizeName(client?.name) === sharePointName);
  if (exact.length) {
    return { candidates: exact, matchType: "exact" };
  }

  const fuzzy = validClients.filter((client) => fuzzyNameMatch(client?.name, sharePointClient?.name));
  if (fuzzy.length) {
    return { candidates: fuzzy, matchType: "fuzzy" };
  }

  return { candidates: [], matchType: "none" };
}

function summarizeClient(client) {
  return {
    id: normalizeText(client?.id),
    itemId: normalizeText(client?.itemId),
    oneTouchId: normalizeText(client?.oneTouchId),
    name: normalizeText(client?.name),
    dateOfBirth: normalizeText(client?.dateOfBirth),
    postcode: normalizeText(client?.postcode),
    email: normalizeText(client?.email),
    phone: normalizeText(client?.phone),
    status: normalizeText(client?.status),
    address: normalizeText(client?.address),
    town: normalizeText(client?.town),
    county: normalizeText(client?.county),
  };
}

function oneTouchToSharePointFields(client) {
  const email = buildCombinedEmail(client);
  const phone = buildCombinedPhone(client);
  return {
    name: normalizeText(client?.name),
    dateOfBirth: normalizeText(client?.dateOfBirth),
    oneTouchId: normalizeText(client?.id),
    postcode: normalizeText(client?.postcode),
    email,
    phone,
    status: normalizeText(client?.status).toLowerCase(),
    address: normalizeText(client?.address),
    town: normalizeText(client?.town),
    county: normalizeText(client?.county),
  };
}

function resolveSyncableFields(fieldMap) {
  return CORE_FIELDS.filter((field) => {
    if (field === "name") {
      return true;
    }
    return Boolean(fieldMap?.[field]);
  });
}

function diffFields(sharePointClient, oneTouchClient, syncableFields = CORE_FIELDS) {
  const target = summarizeClient(sharePointClient);
  const source = oneTouchToSharePointFields(oneTouchClient);
  const diffs = [];
  for (const field of syncableFields) {
    const left = normalizeText(target[field]);
    const right = normalizeText(source[field]);
    if (!right) {
      continue;
    }
    if (left !== right) {
      diffs.push({
        field,
        sharePoint: left,
        oneTouch: right,
      });
    }
  }
  return diffs;
}

function computePatchPlan(sharePointClient, oneTouchClient, syncableFields = CORE_FIELDS) {
  const target = summarizeClient(sharePointClient);
  const source = oneTouchToSharePointFields(oneTouchClient);
  const differences = [];
  const patchFields = {};
  const skippedFields = [];
  for (const field of syncableFields) {
    const left = normalizeText(target[field]);
    const right = normalizeText(source[field]);
    if (!right) {
      skippedFields.push({ field, reason: "source_blank" });
      continue;
    }
    if (left === right) {
      continue;
    }
    patchFields[field] = right;
    differences.push({
        field,
        sharePoint: left,
        oneTouch: right,
    });
  }
  return {
    differences,
    patchFields,
    skippedFields,
    source,
  };
}

function buildFingerprint(payload) {
  const stable = JSON.stringify(payload);
  return crypto.createHash("sha256").update(stable).digest("hex").slice(0, 16);
}

async function loadSnapshot() {
  const [oneTouchClients, sharePointBundle] = await Promise.all([
    listOneTouchClients(),
    listSharePointClientsWithItemIds(),
  ]);
  const sharePointClients = Array.isArray(sharePointBundle?.clients) ? sharePointBundle.clients : [];
  const sharePointFieldMap = sharePointBundle?.fieldMap || {};
  return {
    oneTouchClients,
    sharePointClients,
    sharePointFieldMap,
  };
}

async function buildReconciliationPreview() {
  const { oneTouchClients, sharePointClients, sharePointFieldMap } = await loadSnapshot();
  const syncableFields = resolveSyncableFields(sharePointFieldMap);

  const oneTouchById = new Map();
  for (const client of oneTouchClients) {
    const key = normalizeId(client?.id);
    if (!key || oneTouchById.has(key)) {
      continue;
    }
    oneTouchById.set(key, client);
  }

  const matchedOneTouchIds = new Set();
  const consumedForCopy = new Set();
  const copyOneTouchIdCandidates = [];
  const ambiguousMatches = [];
  const updateCandidates = [];
  const errors = [];

  for (const sharePointClient of sharePointClients) {
    const oneTouchIdKey = normalizeId(sharePointClient?.oneTouchId);
    if (oneTouchIdKey) {
      const source = oneTouchById.get(oneTouchIdKey);
      if (!source) {
        errors.push({
          type: "missing_onetouch_source",
          sharePoint: summarizeClient(sharePointClient),
        });
        continue;
      }

      matchedOneTouchIds.add(oneTouchIdKey);
      const patchPlan = computePatchPlan(sharePointClient, source, syncableFields);
      if (patchPlan.differences.length) {
        updateCandidates.push({
          sharePoint: summarizeClient(sharePointClient),
          oneTouch: summarizeClient(source),
          differences: patchPlan.differences,
          expectedFingerprint: buildFingerprint({
            itemId: normalizeText(sharePointClient?.itemId),
            oneTouchId: normalizeText(source?.id),
            sharePoint: summarizeClient(sharePointClient),
          }),
        });
      }
      continue;
    }

    const { candidates, matchType } = findNameCandidates(sharePointClient, oneTouchClients);

    if (candidates.length === 1) {
      const source = candidates[0];
      const sourceIdKey = normalizeId(source?.id);
      consumedForCopy.add(sourceIdKey);
      copyOneTouchIdCandidates.push({
        sharePoint: summarizeClient(sharePointClient),
        oneTouch: summarizeClient(source),
        matchType,
        expectedFingerprint: buildFingerprint({
          itemId: normalizeText(sharePointClient?.itemId),
          oneTouchId: normalizeText(source?.id),
          sharePoint: summarizeClient(sharePointClient),
        }),
      });
      continue;
    }

    if (candidates.length > 1) {
      ambiguousMatches.push({
        sharePoint: summarizeClient(sharePointClient),
        oneTouchCandidates: candidates.map((client) => summarizeClient(client)),
        matchType,
        expectedFingerprint: buildFingerprint({
          itemId: normalizeText(sharePointClient?.itemId),
          sharePoint: summarizeClient(sharePointClient),
          candidateIds: candidates.map((item) => normalizeText(item?.id)).sort(),
        }),
      });
      continue;
    }

    errors.push({
      type: "missing_onetouch_by_name",
      sharePoint: summarizeClient(sharePointClient),
      message: "No OneTouch client found by name fallback.",
    });
  }

  const missingInSharePoint = [];
  for (const oneTouchClient of oneTouchClients) {
    const sourceIdKey = normalizeId(oneTouchClient?.id);
    if (!sourceIdKey) {
      continue;
    }
    if (matchedOneTouchIds.has(sourceIdKey) || consumedForCopy.has(sourceIdKey)) {
      continue;
    }
    missingInSharePoint.push({
      oneTouch: summarizeClient(oneTouchClient),
      expectedFingerprint: buildFingerprint({
        oneTouchId: normalizeText(oneTouchClient?.id),
        oneTouch: summarizeClient(oneTouchClient),
      }),
    });
  }

  return {
    copyOneTouchIdCandidates,
    missingInSharePoint,
    updateCandidates,
    ambiguousMatches,
    errors,
    totals: {
      copyOneTouchIdCandidates: copyOneTouchIdCandidates.length,
      missingInSharePoint: missingInSharePoint.length,
      updateCandidates: updateCandidates.length,
      ambiguousMatches: ambiguousMatches.length,
      errors: errors.length,
    },
    syncableFields,
  };
}

function findSharePointClientByItemId(clients, itemId) {
  const target = normalizeText(itemId);
  return clients.find((item) => normalizeText(item?.itemId) === target) || null;
}

function findOneTouchClientById(clients, id) {
  const target = normalizeId(id);
  return clients.find((item) => normalizeId(item?.id) === target) || null;
}

async function applyReconciliationAction(body) {
  const action = normalizeText(body?.action).toLowerCase();
  const sharePointItemId = normalizeText(body?.sharePointItemId);
  const oneTouchClientId = normalizeText(body?.oneTouchClientId);
  const expectedFingerprint = normalizeText(body?.expectedFingerprint);

  const { oneTouchClients, sharePointClients, sharePointFieldMap } = await loadSnapshot();
  const syncableFields = resolveSyncableFields(sharePointFieldMap);

  if (action === "copy_onetouch_id") {
    const sharePointClient = findSharePointClientByItemId(sharePointClients, sharePointItemId);
    const oneTouchClient = findOneTouchClientById(oneTouchClients, oneTouchClientId);
    if (!sharePointClient || !oneTouchClient) {
      const error = new Error("Could not resolve selected SharePoint or OneTouch client.");
      error.status = 404;
      throw error;
    }
    if (normalizeText(sharePointClient.oneTouchId)) {
      const error = new Error("SharePoint OneTouchID is already populated.");
      error.status = 409;
      throw error;
    }
    if (!fuzzyNameMatch(sharePointClient.name, oneTouchClient.name)) {
      const error = new Error("Selected OneTouch client no longer matches SharePoint name.");
      error.status = 409;
      throw error;
    }
    const currentFingerprint = buildFingerprint({
      itemId: normalizeText(sharePointClient.itemId),
      oneTouchId: normalizeText(oneTouchClient.id),
      sharePoint: summarizeClient(sharePointClient),
    });
    if (expectedFingerprint && expectedFingerprint !== currentFingerprint) {
      const error = new Error("Record changed since preview. Refresh reconciliation and retry.");
      error.status = 409;
      throw error;
    }
    await patchSharePointClient(sharePointClient.itemId, {
      oneTouchId: oneTouchClient.id,
    });
    return {
      action,
      updated: true,
      sharePointItemId: sharePointClient.itemId,
      oneTouchClientId: oneTouchClient.id,
    };
  }

  if (action === "add_missing") {
    const oneTouchClient = findOneTouchClientById(oneTouchClients, oneTouchClientId);
    if (!oneTouchClient) {
      const error = new Error("Could not resolve OneTouch client.");
      error.status = 404;
      throw error;
    }
    const alreadyExists = sharePointClients.some(
      (item) => normalizeId(item?.oneTouchId) === normalizeId(oneTouchClient.id)
    );
    if (alreadyExists) {
      const error = new Error("SharePoint record already exists for this OneTouchID.");
      error.status = 409;
      throw error;
    }
    const currentFingerprint = buildFingerprint({
      oneTouchId: normalizeText(oneTouchClient.id),
      oneTouch: summarizeClient(oneTouchClient),
    });
    if (expectedFingerprint && expectedFingerprint !== currentFingerprint) {
      const error = new Error("Source data changed since preview. Refresh reconciliation and retry.");
      error.status = 409;
      throw error;
    }
    const created = await createSharePointClient(oneTouchToSharePointFields(oneTouchClient));
    return {
      action,
      updated: true,
      sharePointItemId: normalizeText(created?.itemId),
      oneTouchClientId: normalizeText(oneTouchClient.id),
    };
  }

  if (action === "update_record") {
    const sharePointClient = findSharePointClientByItemId(sharePointClients, sharePointItemId);
    const oneTouchClient = findOneTouchClientById(oneTouchClients, oneTouchClientId);
    if (!sharePointClient || !oneTouchClient) {
      const error = new Error("Could not resolve selected SharePoint or OneTouch client.");
      error.status = 404;
      throw error;
    }
    if (normalizeId(sharePointClient.oneTouchId) !== normalizeId(oneTouchClient.id)) {
      const error = new Error("SharePoint record no longer matches the selected OneTouchID.");
      error.status = 409;
      throw error;
    }
    const currentFingerprint = buildFingerprint({
      itemId: normalizeText(sharePointClient.itemId),
      oneTouchId: normalizeText(oneTouchClient.id),
      sharePoint: summarizeClient(sharePointClient),
    });
    if (expectedFingerprint && expectedFingerprint !== currentFingerprint) {
      const error = new Error("Record changed since preview. Refresh reconciliation and retry.");
      error.status = 409;
      throw error;
    }

    const patchPlan = computePatchPlan(sharePointClient, oneTouchClient, syncableFields);
    if (!patchPlan.differences.length) {
      return {
        action,
        updated: false,
        sharePointItemId: normalizeText(sharePointClient.itemId),
        oneTouchClientId: normalizeText(oneTouchClient.id),
        skippedFields: patchPlan.skippedFields,
      };
    }

    await patchSharePointClient(sharePointClient.itemId, patchPlan.patchFields);

    return {
      action,
      updated: true,
      sharePointItemId: normalizeText(sharePointClient.itemId),
      oneTouchClientId: normalizeText(oneTouchClient.id),
      changedFields: patchPlan.differences.map((item) => item.field),
      skippedFields: patchPlan.skippedFields,
      contactMerge: {
        emailPartsUsed: dedupeEmails(
          collectNonEmptyValues([
            oneTouchClient?.emailPrimary,
            oneTouchClient?.emailSecondary,
            oneTouchClient?.email,
            oneTouchClient?.raw?.primary_email,
            oneTouchClient?.raw?.secondary_email,
            oneTouchClient?.raw?.email,
            oneTouchClient?.raw?.email_address,
          ])
        ),
        phonePartsUsed: dedupePhones(
          collectNonEmptyValues([
            oneTouchClient?.phoneMobile,
            oneTouchClient?.phoneHome,
            oneTouchClient?.phone,
            oneTouchClient?.raw?.phone_mobile,
            oneTouchClient?.raw?.phone_home,
            oneTouchClient?.raw?.phone,
            oneTouchClient?.raw?.mobile,
          ])
        ),
      },
    };
  }

  const error = new Error("Unsupported reconciliation action.");
  error.status = 400;
  throw error;
}

module.exports = {
  applyReconciliationAction,
  buildReconciliationPreview,
};
