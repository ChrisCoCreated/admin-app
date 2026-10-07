// Shared repository: keep this file identical in Admin and Associates.
const { ROLES, ROLE_DESCRIPTIONS, normalizeEmail, getProtectedEmail, failure, validateRoles, validateAssignment, effectiveRoles } = require("./shared-roles");
async function request(table, query = {}, options = {}) {
  const host = String(process.env.SUPABASE_URL || "").trim().replace(/\/+$/, "");
  const key = String(process.env.SUPABASE_SERVICE_ROLE_KEY || "").trim();
  if (!host || !key) throw failure("Shared role storage is not configured.");
  const url = new URL(`${host}/rest/v1/${table}`);
  for (const [name, value] of Object.entries(query)) url.searchParams.set(name, String(value));
  try {
    const response = await fetch(url, {
      ...options, signal: AbortSignal.timeout(15000),
      headers: { Accept: "application/json", apikey: key, Authorization: `Bearer ${key}`, ...options.headers },
    });
    if (!response.ok) throw failure("Shared role storage is temporarily unavailable.");
    return response.status === 204 ? null : await response.json();
  } catch (error) { throw error.status ? error : failure("Shared role storage is temporarily unavailable."); }
}
async function loadSnapshot({ requireActive = true } = {}) {
  const states = await request("app_role_management_state", { select: "access_enabled,writes_enabled", id: "eq.shared" });
  const state = Array.isArray(states) && states.length === 1 ? states[0] : null;
  if (!state || (requireActive && state.access_enabled !== true)) throw failure("Shared access is awaiting migration review and activation.");
  const assignments = Object.create(null), metadata = Object.create(null);
  for (let offset = 0; ;) {
    const rows = await request("app_role_assignments", { select: "email,roles,updated_by,updated_at", order: "email.asc", limit: 500, offset });
    if (!Array.isArray(rows)) throw failure("Invalid shared role storage response.");
    for (const row of rows) {
      const email = normalizeEmail(row.email);
      if (!email || email !== row.email || Object.hasOwn(assignments, email)) throw failure("Invalid shared role assignment.");
      try { assignments[email] = validateRoles(row.roles); } catch { throw failure("Shared role assignments need migration review."); }
      metadata[email] = { updatedBy: row.updated_by, updatedAt: row.updated_at };
    }
    if (!rows.length) break;
    offset += rows.length;
  }
  const aliases = Object.create(null);
  for (let offset = 0; ;) {
    const rows = await request("app_role_aliases", { select: "alias_email,canonical_email", order: "alias_email.asc", limit: 500, offset });
    if (!Array.isArray(rows)) throw failure("Invalid shared identity response.");
    for (const row of rows) {
      if (row.alias_email !== normalizeEmail(row.alias_email) || row.canonical_email !== normalizeEmail(row.canonical_email)
        || row.alias_email === row.canonical_email || Object.hasOwn(assignments, row.alias_email)
        || !Object.hasOwn(assignments, row.canonical_email) || Object.hasOwn(aliases, row.alias_email)) throw failure("Shared identities need migration review.");
      aliases[row.alias_email] = row.canonical_email;
    }
    if (!rows.length) break;
    offset += rows.length;
  }
  return { assignments, metadata, state, aliases };
}
function profileFor(email, snapshot) {
  const identityEmail = normalizeEmail(email);
  email = snapshot.aliases?.[identityEmail] || identityEmail;
  const roles = effectiveRoles(email, snapshot.assignments);
  const aliases = Object.entries(snapshot.aliases || {}).filter(([,canonical]) => canonical === email).map(([alias]) => alias);
  return { email, identityEmail, aliases, roles, isSuperadmin: roles.includes("superadmin") };
}
async function listAssignments(snapshot) {
  snapshot ||= await loadSnapshot();
  const emails = new Set([...Object.keys(snapshot.assignments), getProtectedEmail()]);
  return {
    roles: ROLES, descriptions: ROLE_DESCRIPTIONS, writesEnabled: snapshot.state.writes_enabled === true,
    users: [...emails].sort().map(email => ({ ...profileFor(email, snapshot), protected: email === getProtectedEmail(), ...snapshot.metadata[email] })),
  };
}
async function saveAssignment(body, actor) {
  const snapshot = await loadSnapshot();
  const actorProfile = profileFor(actor, snapshot);
  if (!actorProfile.isSuperadmin) throw failure("Superadmin access required.", 403);
  const email = snapshot.aliases[normalizeEmail(body?.email)] || body?.email;
  const assignment = validateAssignment({ ...body, email }, getProtectedEmail(), actorProfile.email);
  if (snapshot.state.writes_enabled !== true) throw failure("Role editing is paused during rollout.");
  await request("app_role_assignments", { on_conflict: "email" }, {
    method: "POST", headers: { "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=minimal" },
    body: JSON.stringify({ ...assignment, updated_by: actorProfile.email, updated_at: new Date().toISOString() }),
  });
  return assignment;
}
module.exports = { loadSnapshot, profileFor, listAssignments, saveAssignment, request };
