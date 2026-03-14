const ACCESS_ENV_KEYS = [
  "ACCESS_FULL_EMAILS",
  "ACCESS_DIRECTOR_EMAILS",
  "ACCESS_MARKETING_EMAILS",
  "ACCESS_PHOTO_LAYOUT_EMAILS",
  "ACCESS_TIME_EMAILS",
  "ACCESS_HR_EMAILS",
  "ACCESS_CLIENTS_EMAILS",
  "ACCESS_CONSULTANT_EMAILS",
];

const ROLE_BY_PAGE_KEY = new Map(
  [
    ["clients,carers,whiteboard,simpletasks,tasks,mapping,drivetime,reports,marketing,photolayout", "admin"],
    ["marketing,photolayout", "marketing"],
    ["photolayout", "photo_layout"],
    ["mapping,drivetime", "time_only"],
    ["carers", "hr_only"],
    ["clients", "clients_only"],
    ["clients,carers", "hr_clients"],
    ["clients,mapping,drivetime", "time_clients"],
    ["carers,mapping,drivetime", "time_hr"],
    ["clients,carers,mapping,drivetime", "time_hr_clients"],
    ["consultant", "consultant"],
  ].map(([pageKey, role]) => [canonicalizePagesKey(pageKey.split(",")), role])
);

function canonicalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function canonicalizePagesKey(pages) {
  return Array.from(
    new Set(
      (Array.isArray(pages) ? pages : [])
        .map((page) => String(page || "").trim().toLowerCase())
        .filter(Boolean)
    )
  )
    .sort()
    .join(",");
}

function parseEmailList(value) {
  return String(value || "")
    .split(",")
    .map((email) => canonicalizeEmail(email))
    .filter(Boolean);
}

function hasAccessEnvConfig() {
  return ACCESS_ENV_KEYS.some((key) => String(process.env[key] || "").trim() !== "");
}

function resolveRoleFromFlags(flags) {
  if (flags.full) {
    return "admin";
  }

  if (flags.director) {
    return "director";
  }

  const pages = [];
  if (flags.marketing) {
    pages.push("marketing", "photolayout");
  }
  if (flags.photoLayout) {
    pages.push("photolayout");
  }
  if (flags.time) {
    pages.push("mapping", "drivetime");
  }
  if (flags.hr) {
    pages.push("carers");
  }
  if (flags.clients) {
    pages.push("clients");
  }
  if (flags.consultant) {
    pages.push("consultant");
  }

  const pageKey = canonicalizePagesKey(pages);
  return ROLE_BY_PAGE_KEY.get(pageKey) || null;
}

function buildAuthorizedUsersFromEnv() {
  const flagsByEmail = new Map();

  function mark(emails, flagKey) {
    for (const email of emails) {
      const existing = flagsByEmail.get(email) || {
        full: false,
        director: false,
        marketing: false,
        photoLayout: false,
        time: false,
        hr: false,
        clients: false,
        consultant: false,
      };
      existing[flagKey] = true;
      flagsByEmail.set(email, existing);
    }
  }

  mark(parseEmailList(process.env.ACCESS_FULL_EMAILS), "full");
  mark(parseEmailList(process.env.ACCESS_DIRECTOR_EMAILS), "director");
  mark(parseEmailList(process.env.ACCESS_MARKETING_EMAILS), "marketing");
  mark(parseEmailList(process.env.ACCESS_PHOTO_LAYOUT_EMAILS), "photoLayout");
  mark(parseEmailList(process.env.ACCESS_TIME_EMAILS), "time");
  mark(parseEmailList(process.env.ACCESS_HR_EMAILS), "hr");
  mark(parseEmailList(process.env.ACCESS_CLIENTS_EMAILS), "clients");
  mark(parseEmailList(process.env.ACCESS_CONSULTANT_EMAILS), "consultant");

  const map = new Map();
  for (const [email, flags] of flagsByEmail.entries()) {
    const role = resolveRoleFromFlags(flags);
    if (!role) {
      continue;
    }
    map.set(email, role);
  }

  return map;
}
function getAuthorizedUsersMap() {
  if (!hasAccessEnvConfig()) {
    return new Map();
  }
  return buildAuthorizedUsersFromEnv();
}

module.exports = {
  getAuthorizedUsersMap,
};
