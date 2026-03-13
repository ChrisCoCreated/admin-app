const { supabaseRestFetch } = require("./supabase-rest");

const SCORECARD_TABLE = "performance_alignment_scorecards";

function normalizeReviewPeriod(value) {
  const reviewPeriod = String(value || "").trim();
  if (!reviewPeriod) {
    throw new Error("Review period is required.");
  }
  if (reviewPeriod.length > 120) {
    throw new Error("Review period is too long.");
  }
  return reviewPeriod;
}

function assertPayload(value) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("Scorecard payload must be an object.");
  }
  return value;
}

async function getScorecard(reviewPeriod) {
  const rows = await supabaseRestFetch(SCORECARD_TABLE, {
    query: {
      select: "review_period,payload,created_by,updated_by,created_at,updated_at",
      review_period: `eq.${normalizeReviewPeriod(reviewPeriod)}`,
      limit: "1",
    },
  });

  return Array.isArray(rows) ? rows[0] || null : null;
}

async function createScorecard({ reviewPeriod, payload, email }) {
  const rows = await supabaseRestFetch(SCORECARD_TABLE, {
    method: "POST",
    query: {
      select: "review_period,payload,created_by,updated_by,created_at,updated_at",
    },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: [
      {
        review_period: normalizeReviewPeriod(reviewPeriod),
        payload: assertPayload(payload),
        created_by: email || null,
        updated_by: email || null,
      },
    ],
  });

  return Array.isArray(rows) ? rows[0] || null : null;
}

async function updateScorecard({ reviewPeriod, payload, email }) {
  const rows = await supabaseRestFetch(SCORECARD_TABLE, {
    method: "PATCH",
    query: {
      select: "review_period,payload,created_by,updated_by,created_at,updated_at",
      review_period: `eq.${normalizeReviewPeriod(reviewPeriod)}`,
    },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: {
      payload: assertPayload(payload),
      updated_by: email || null,
    },
  });

  return Array.isArray(rows) ? rows[0] || null : null;
}

async function upsertScorecard({ reviewPeriod, payload, email }) {
  const existing = await getScorecard(reviewPeriod);
  if (!existing) {
    return createScorecard({ reviewPeriod, payload, email });
  }
  return updateScorecard({ reviewPeriod, payload, email });
}

module.exports = {
  getScorecard,
  normalizeReviewPeriod,
  upsertScorecard,
};
