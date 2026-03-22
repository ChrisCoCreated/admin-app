const { supabaseRestFetch } = require("../supabase-rest");

function normalizeEmail(value) {
  return String(value || "").trim().toLowerCase();
}

function isUuid(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || ""));
}

function ensureUuid(value, label) {
  if (!isUuid(value)) {
    const error = new Error(`${label} is invalid.`);
    error.status = 400;
    error.code = "INVALID_UUID";
    throw error;
  }
  return String(value);
}

async function listProblemRowsByOwnerEmail(ownerEmail) {
  const rows = await supabaseRestFetch("problems_to_solve", {
    query: {
      select: "*",
      owner_email: `eq.${normalizeEmail(ownerEmail)}`,
      order: "updated_at.desc,created_at.desc",
    },
  });
  return Array.isArray(rows) ? rows : [];
}

async function getProblemRowByIdAndOwner(problemId, ownerEmail) {
  const rows = await supabaseRestFetch("problems_to_solve", {
    query: {
      select: "*",
      id: `eq.${ensureUuid(problemId, "Problem")}`,
      owner_email: `eq.${normalizeEmail(ownerEmail)}`,
      limit: "1",
    },
  });
  return Array.isArray(rows) ? rows[0] || null : null;
}

async function createProblemRow(row) {
  const rows = await supabaseRestFetch("problems_to_solve", {
    method: "POST",
    query: { select: "*" },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: [row],
  });
  return Array.isArray(rows) ? rows[0] || null : null;
}

async function updateProblemRow(problemId, ownerEmail, patch) {
  const rows = await supabaseRestFetch("problems_to_solve", {
    method: "PATCH",
    query: {
      select: "*",
      id: `eq.${ensureUuid(problemId, "Problem")}`,
      owner_email: `eq.${normalizeEmail(ownerEmail)}`,
    },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: patch,
  });
  return Array.isArray(rows) ? rows[0] || null : null;
}

module.exports = {
  createProblemRow,
  ensureUuid,
  getProblemRowByIdAndOwner,
  listProblemRowsByOwnerEmail,
  normalizeEmail,
  updateProblemRow,
};
