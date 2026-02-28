const ALLOWED_PROVIDERS = new Set(["todo", "planner"]);

const WRITABLE_OVERLAY_FIELDS = new Set([
  "title",
  "workingStatus",
  "workType",
  "tags",
  "activeStartedAt",
  "lastWorkedAt",
  "energy",
  "effortMinutes",
  "impact",
  "overlayNotes",
  "pinned",
  "layout",
  "category",
]);

const OVERLAY_FIELD_MAP = {
  title: "Title",
  userUpn: "UserUPN",
  provider: "Provider",
  externalTaskId: "ExternalTaskId",
  workingStatus: "WorkingStatus",
  workType: "WorkType",
  tags: "Tags",
  activeStartedAt: "ActiveStartedAt",
  lastWorkedAt: "LastWorkedAt",
  energy: "Energy",
  effortMinutes: "EffortMinutes",
  impact: "Impact",
  overlayNotes: "OverlayNotes",
  pinned: "Pinned",
  layout: "Layout",
  category: "Category",
  lastOverlayUpdatedAt: "LastOverlayUpdatedAt",
};

function normalizeProvider(value) {
  const provider = String(value || "")
    .trim()
    .toLowerCase();

  if (!ALLOWED_PROVIDERS.has(provider)) {
    const error = new Error("Provider must be one of: todo, planner.");
    error.status = 400;
    error.code = "INVALID_PROVIDER";
    throw error;
  }

  return provider;
}

function normalizeExternalTaskId(value) {
  const id = String(value || "").trim();
  if (!id) {
    const error = new Error("externalTaskId is required.");
    error.status = 400;
    error.code = "INVALID_EXTERNAL_TASK_ID";
    throw error;
  }
  return id;
}

function buildTaskKey(provider, externalTaskId) {
  return `${normalizeProvider(provider)}|${normalizeExternalTaskId(externalTaskId)}`;
}

function toUtcIsoOrNull(value) {
  if (!value) {
    return null;
  }

  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }

  return parsed.toISOString();
}

function parseBoolean(value) {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return value !== 0;
  }
  const text = String(value || "")
    .trim()
    .toLowerCase();
  if (!text) {
    return false;
  }
  return text === "1" || text === "true" || text === "yes";
}

function parseNumberOrNull(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function parseTags(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => String(entry || "").trim()).filter(Boolean);
  }

  const text = String(value || "").trim();
  if (!text) {
    return [];
  }

  if (text.startsWith("[") || text.startsWith("{")) {
    try {
      const parsed = JSON.parse(text);
      if (Array.isArray(parsed)) {
        return parsed.map((entry) => String(entry || "").trim()).filter(Boolean);
      }
      if (parsed && Array.isArray(parsed.tags)) {
        return parsed.tags.map((entry) => String(entry || "").trim()).filter(Boolean);
      }
    } catch {
      return [];
    }
  }

  return text
    .split(",")
    .map((entry) => entry.trim())
    .filter(Boolean);
}

function parseUserUpnValue(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "string") {
    const text = value.trim().toLowerCase();
    const claimMatch = /i:0[.]f[|][^|]+[|]([a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,})/i.exec(text);
    if (claimMatch && claimMatch[1]) {
      return claimMatch[1].toLowerCase();
    }
    const emailMatch = /[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i.exec(text);
    if (emailMatch) {
      return emailMatch[0].toLowerCase();
    }
    return text;
  }

  if (typeof value === "object") {
    if (Array.isArray(value)) {
      const first = value.map((entry) => parseUserUpnValue(entry)).find(Boolean);
      return first || "";
    }
    const claimSource = String(
      value.claims ||
        value.Claims ||
        value.name ||
        value.Name ||
        ""
    )
      .trim()
      .toLowerCase();
    if (claimSource) {
      const claimMatch = /i:0[.]f[|][^|]+[|]([a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,})/i.exec(claimSource);
      if (claimMatch && claimMatch[1]) {
        return claimMatch[1].toLowerCase();
      }
      const emailMatch = /[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/i.exec(claimSource);
      if (emailMatch) {
        return emailMatch[0].toLowerCase();
      }
    }
    return String(
      value.email ||
        value.userPrincipalName ||
        value.upn ||
        value.lookupValue ||
        value.LookupValue ||
        value.displayName ||
        value.DisplayName ||
        value.value ||
        ""
    )
      .trim()
      .toLowerCase();
  }

  return String(value).trim().toLowerCase();
}

function sanitizeOverlayPatch(patch) {
  if (!patch || typeof patch !== "object" || Array.isArray(patch)) {
    const error = new Error("patch must be an object.");
    error.status = 400;
    error.code = "INVALID_OVERLAY_PATCH";
    throw error;
  }

  const sanitized = {};
  for (const [key, value] of Object.entries(patch)) {
    if (!WRITABLE_OVERLAY_FIELDS.has(key)) {
      continue;
    }

    if (key === "pinned") {
      sanitized.pinned = parseBoolean(value);
      continue;
    }

    if (key === "tags") {
      sanitized.tags = Array.isArray(value)
        ? value.map((entry) => String(entry || "").trim()).filter(Boolean)
        : parseTags(value);
      continue;
    }

    if (key === "effortMinutes") {
      const effort = parseNumberOrNull(value);
      sanitized.effortMinutes = effort;
      continue;
    }

    if (key === "activeStartedAt" || key === "lastWorkedAt") {
      sanitized[key] = value ? toUtcIsoOrNull(value) : null;
      continue;
    }

    sanitized[key] = value === null || value === undefined ? "" : String(value).trim();
  }

  return sanitized;
}

function toGraphOverlayFields(input) {
  const fields = {};

  fields[OVERLAY_FIELD_MAP.userUpn] = String(input.userUpn || "").trim().toLowerCase();
  if (input.userUpnLookupId !== null && input.userUpnLookupId !== undefined && input.userUpnLookupId !== "") {
    fields[`${OVERLAY_FIELD_MAP.userUpn}LookupId`] = Number(input.userUpnLookupId);
  }
  fields[OVERLAY_FIELD_MAP.provider] = normalizeProvider(input.provider);
  fields[OVERLAY_FIELD_MAP.externalTaskId] = normalizeExternalTaskId(input.externalTaskId);
  fields[OVERLAY_FIELD_MAP.lastOverlayUpdatedAt] = toUtcIsoOrNull(input.lastOverlayUpdatedAt) || new Date().toISOString();

  const optional = [
    "title",
    "workingStatus",
    "workType",
    "activeStartedAt",
    "lastWorkedAt",
    "energy",
    "effortMinutes",
    "impact",
    "overlayNotes",
    "pinned",
    "layout",
    "category",
  ];

  for (const key of optional) {
    if (Object.prototype.hasOwnProperty.call(input, key)) {
      fields[OVERLAY_FIELD_MAP[key]] = input[key];
    }
  }

  if (Object.prototype.hasOwnProperty.call(input, "tags")) {
    fields[OVERLAY_FIELD_MAP.tags] = JSON.stringify(Array.isArray(input.tags) ? input.tags : []);
  }

  return fields;
}

function fromOverlayFields(item) {
  const fields = item?.fields || {};
  const provider = String(fields[OVERLAY_FIELD_MAP.provider] || "")
    .trim()
    .toLowerCase();
  const externalTaskId = String(fields[OVERLAY_FIELD_MAP.externalTaskId] || "").trim();

  if (!provider || !externalTaskId) {
    return null;
  }

  let tags = [];
  const rawTags = fields[OVERLAY_FIELD_MAP.tags];
  if (typeof rawTags === "string") {
    tags = parseTags(rawTags);
  } else if (Array.isArray(rawTags)) {
    tags = parseTags(rawTags);
  }

  const overlay = {
    itemId: String(item?.id || ""),
    title: String(fields[OVERLAY_FIELD_MAP.title] || "").trim(),
    userUpn: parseUserUpnValue(fields[OVERLAY_FIELD_MAP.userUpn]),
    provider,
    externalTaskId,
    workingStatus: String(fields[OVERLAY_FIELD_MAP.workingStatus] || "").trim().toLowerCase(),
    workType: String(fields[OVERLAY_FIELD_MAP.workType] || "").trim(),
    tags,
    activeStartedAt: toUtcIsoOrNull(fields[OVERLAY_FIELD_MAP.activeStartedAt]),
    lastWorkedAt: toUtcIsoOrNull(fields[OVERLAY_FIELD_MAP.lastWorkedAt]),
    energy: fields[OVERLAY_FIELD_MAP.energy] ?? "",
    effortMinutes: parseNumberOrNull(fields[OVERLAY_FIELD_MAP.effortMinutes]),
    impact: fields[OVERLAY_FIELD_MAP.impact] ?? "",
    overlayNotes: String(fields[OVERLAY_FIELD_MAP.overlayNotes] || "").trim(),
    pinned: parseBoolean(fields[OVERLAY_FIELD_MAP.pinned]),
    layout: String(fields[OVERLAY_FIELD_MAP.layout] || "").trim(),
    category: String(fields[OVERLAY_FIELD_MAP.category] || "").trim(),
    lastOverlayUpdatedAt: toUtcIsoOrNull(fields[OVERLAY_FIELD_MAP.lastOverlayUpdatedAt]),
  };

  overlay.key = buildTaskKey(overlay.provider, overlay.externalTaskId);
  return overlay;
}

module.exports = {
  OVERLAY_FIELD_MAP,
  buildTaskKey,
  fromOverlayFields,
  normalizeExternalTaskId,
  normalizeProvider,
  sanitizeOverlayPatch,
  toGraphOverlayFields,
  toUtcIsoOrNull,
};
