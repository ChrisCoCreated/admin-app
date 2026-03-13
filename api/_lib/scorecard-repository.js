const { supabaseRestFetch } = require("./supabase-rest");

const ENTITY_CONFIG = {
  values: {
    table: "scorecard_values",
    reviewTable: "scorecard_value_reviews",
    idField: "value_id",
    labelField: "name",
    createFields: ["name", "description", "sort_order", "is_active"],
    reviewSelect: "id,review_id,value_id,name_snapshot,score,evidence,improvement_action,created_at,updated_at",
  },
  principles: {
    table: "scorecard_principles",
    reviewTable: "scorecard_principle_reviews",
    idField: "principle_id",
    labelField: "name",
    createFields: ["name", "description", "sort_order", "is_active"],
    reviewSelect: "id,review_id,principle_id,name_snapshot,score,example_decision,adjustment_needed,created_at,updated_at",
  },
  objectives: {
    table: "scorecard_objectives",
    reviewTable: "scorecard_objective_reviews",
    idField: "objective_id",
    labelField: "title",
    createFields: ["title", "description", "default_owner", "sort_order", "is_active"],
    reviewSelect: "id,review_id,objective_id,title_snapshot,owner,progress_score,risks,created_at,updated_at",
  },
  goals: {
    table: "scorecard_goals",
    reviewTable: "scorecard_goal_reviews",
    idField: "goal_id",
    labelField: "title",
    createFields: ["title", "description", "default_owner", "objective_id", "sort_order", "is_active"],
    reviewSelect:
      "id,review_id,goal_id,title_snapshot,objective_id_snapshot,objective_title_snapshot,owner,status,progress_score,next_action,created_at,updated_at",
  },
};

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

function normalizeEntityType(entityType) {
  const normalized = String(entityType || "").trim().toLowerCase();
  if (!ENTITY_CONFIG[normalized]) {
    throw new Error("Unknown scorecard entity type.");
  }
  return normalized;
}

function isUuid(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i.test(String(value || ""));
}

function ensureUuid(value, label) {
  if (!isUuid(value)) {
    throw new Error(`${label} is invalid.`);
  }
  return String(value);
}

function normalizeNumber(value, fallback = 0) {
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function normalizeNullableScore(value) {
  if (value === "" || value === null || value === undefined) {
    return null;
  }
  const score = Number.parseInt(value, 10);
  if (!Number.isFinite(score) || score < 1 || score > 5) {
    throw new Error("Scores must be between 1 and 5.");
  }
  return score;
}

function normalizeNullableText(value) {
  const text = String(value || "").trim();
  return text || null;
}

function buildInFilter(ids) {
  const safeIds = ids.filter((id) => isUuid(id)).map((id) => String(id));
  return safeIds.length ? `in.(${safeIds.join(",")})` : null;
}

function mapDefinitionRow(entityType, row) {
  if (entityType === "values" || entityType === "principles") {
    return {
      id: row.id,
      name: row.name || "",
      description: row.description || "",
      sortOrder: Number(row.sort_order || 0),
      isActive: row.is_active !== false,
      createdAt: row.created_at || null,
      updatedAt: row.updated_at || null,
    };
  }

  if (entityType === "objectives") {
    return {
      id: row.id,
      title: row.title || "",
      description: row.description || "",
      defaultOwner: row.default_owner || "",
      sortOrder: Number(row.sort_order || 0),
      isActive: row.is_active !== false,
      createdAt: row.created_at || null,
      updatedAt: row.updated_at || null,
    };
  }

  return {
    id: row.id,
    title: row.title || "",
    description: row.description || "",
    defaultOwner: row.default_owner || "",
    objectiveId: row.objective_id || null,
    sortOrder: Number(row.sort_order || 0),
    isActive: row.is_active !== false,
    createdAt: row.created_at || null,
    updatedAt: row.updated_at || null,
  };
}

function sortDefinitions(list, entityType) {
  const labelField = entityType === "values" || entityType === "principles" ? "name" : "title";
  return [...list].sort((a, b) => {
    const orderDelta = Number(a.sortOrder || 0) - Number(b.sortOrder || 0);
    if (orderDelta !== 0) {
      return orderDelta;
    }
    return String(a[labelField] || "").localeCompare(String(b[labelField] || ""), undefined, {
      sensitivity: "base",
    });
  });
}

async function listDefinitions(entityType, options = {}) {
  const normalizedEntityType = normalizeEntityType(entityType);
  const config = ENTITY_CONFIG[normalizedEntityType];
  if (Array.isArray(options.ids) && options.ids.length === 0) {
    return [];
  }
  const query = {
    select: "*",
    order: normalizedEntityType === "goals" ? "sort_order.asc,title.asc" : "sort_order.asc",
  };

  if (options.activeOnly) {
    query.is_active = "eq.true";
  }

  if (Array.isArray(options.ids) && options.ids.length) {
    const filter = buildInFilter(options.ids);
    if (filter) {
      query.id = filter;
    }
  }

  const rows = await supabaseRestFetch(config.table, { query });
  return sortDefinitions(
    (Array.isArray(rows) ? rows : []).map((row) => mapDefinitionRow(normalizedEntityType, row)),
    normalizedEntityType
  );
}

async function getDefinitionById(entityType, id) {
  const rows = await listDefinitions(entityType, { ids: [id] });
  return rows[0] || null;
}

function sanitizeDefinitionPayload(entityType, payload = {}, mode = "create") {
  const normalizedEntityType = normalizeEntityType(entityType);
  const config = ENTITY_CONFIG[normalizedEntityType];
  const output = {};

  for (const field of config.createFields) {
    if (mode === "update" && !(field in payload)) {
      continue;
    }

    if (field === "sort_order") {
      output.sort_order = normalizeNumber(payload.sort_order, 0);
      continue;
    }

    if (field === "is_active") {
      output.is_active = payload.is_active !== false;
      continue;
    }

    if (field === "objective_id") {
      output.objective_id = payload.objective_id ? ensureUuid(payload.objective_id, "Objective") : null;
      continue;
    }

    if (field === "default_owner") {
      output.default_owner = normalizeNullableText(payload.default_owner);
      continue;
    }

    if (field === "name" || field === "title" || field === "description") {
      output[field] = String(payload[field] || "");
    }
  }

  return output;
}

async function createDefinition(entityType, payload = {}) {
  const normalizedEntityType = normalizeEntityType(entityType);
  const config = ENTITY_CONFIG[normalizedEntityType];
  const body = sanitizeDefinitionPayload(normalizedEntityType, payload, "create");
  const rows = await supabaseRestFetch(config.table, {
    method: "POST",
    query: { select: "*" },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: [body],
  });

  return mapDefinitionRow(normalizedEntityType, Array.isArray(rows) ? rows[0] || {} : {});
}

async function updateDefinition(entityType, id, patch = {}) {
  const normalizedEntityType = normalizeEntityType(entityType);
  const config = ENTITY_CONFIG[normalizedEntityType];
  const definitionId = ensureUuid(id, "Definition");
  const body = sanitizeDefinitionPayload(normalizedEntityType, patch, "update");

  const rows = await supabaseRestFetch(config.table, {
    method: "PATCH",
    query: {
      select: "*",
      id: `eq.${definitionId}`,
    },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body,
  });

  return mapDefinitionRow(normalizedEntityType, Array.isArray(rows) ? rows[0] || {} : {});
}

async function getReviewByPeriod(reviewPeriod) {
  const rows = await supabaseRestFetch("scorecard_reviews", {
    query: {
      select: "*",
      review_period: `eq.${normalizeReviewPeriod(reviewPeriod)}`,
      limit: "1",
    },
  });

  return Array.isArray(rows) ? rows[0] || null : null;
}

async function fetchReviewRows(entityType, reviewId) {
  const normalizedEntityType = normalizeEntityType(entityType);
  const config = ENTITY_CONFIG[normalizedEntityType];
  if (!reviewId) {
    return [];
  }

  const rows = await supabaseRestFetch(config.reviewTable, {
    query: {
      select: config.reviewSelect,
      review_id: `eq.${ensureUuid(reviewId, "Review")}`,
    },
  });

  return Array.isArray(rows) ? rows : [];
}

function buildReviewMeta(reviewRow) {
  if (!reviewRow) {
    return null;
  }

  return {
    id: reviewRow.id,
    reviewPeriod: reviewRow.review_period,
    reviewType: reviewRow.review_type || null,
    createdAt: reviewRow.created_at || null,
    updatedAt: reviewRow.updated_at || null,
    createdBy: reviewRow.created_by || null,
    updatedBy: reviewRow.updated_by || null,
  };
}

function mergeReviewedDefinitions(activeDefinitions, reviewedDefinitions, entityType) {
  const byId = new Map(activeDefinitions.map((item) => [item.id, item]));
  for (const definition of reviewedDefinitions) {
    if (!byId.has(definition.id)) {
      byId.set(definition.id, definition);
    }
  }
  return sortDefinitions([...byId.values()], entityType);
}

async function getScorecardView(reviewPeriod) {
  const normalizedReviewPeriod = normalizeReviewPeriod(reviewPeriod);
  const reviewRow = await getReviewByPeriod(normalizedReviewPeriod);
  const reviewId = reviewRow?.id || null;

  const [activeValues, activePrinciples, activeObjectives, activeGoals, valueReviews, principleReviews, objectiveReviews, goalReviews] =
    await Promise.all([
      listDefinitions("values", { activeOnly: true }),
      listDefinitions("principles", { activeOnly: true }),
      listDefinitions("objectives", { activeOnly: true }),
      listDefinitions("goals", { activeOnly: true }),
      fetchReviewRows("values", reviewId),
      fetchReviewRows("principles", reviewId),
      fetchReviewRows("objectives", reviewId),
      fetchReviewRows("goals", reviewId),
    ]);

  const [reviewedValues, reviewedPrinciples, reviewedObjectives, reviewedGoals] = await Promise.all([
    listDefinitions("values", { ids: valueReviews.map((row) => row.value_id) }),
    listDefinitions("principles", { ids: principleReviews.map((row) => row.principle_id) }),
    listDefinitions("objectives", { ids: objectiveReviews.map((row) => row.objective_id) }),
    listDefinitions("goals", { ids: goalReviews.map((row) => row.goal_id) }),
  ]);

  const mergedGoals = mergeReviewedDefinitions(activeGoals, reviewedGoals, "goals");
  const goalObjectiveIds = mergedGoals.map((item) => item.objectiveId).filter(Boolean);
  const referencedObjectives = await listDefinitions("objectives", { ids: goalObjectiveIds });
  const objectives = mergeReviewedDefinitions(
    mergeReviewedDefinitions(activeObjectives, reviewedObjectives, "objectives"),
    referencedObjectives,
    "objectives"
  );
  const values = mergeReviewedDefinitions(activeValues, reviewedValues, "values");
  const principles = mergeReviewedDefinitions(activePrinciples, reviewedPrinciples, "principles");
  const goals = mergedGoals;

  const objectiveMap = new Map(objectives.map((item) => [item.id, item]));
  const valueReviewMap = new Map(valueReviews.map((row) => [row.value_id, row]));
  const principleReviewMap = new Map(principleReviews.map((row) => [row.principle_id, row]));
  const objectiveReviewMap = new Map(objectiveReviews.map((row) => [row.objective_id, row]));
  const goalReviewMap = new Map(goalReviews.map((row) => [row.goal_id, row]));

  return {
    review: buildReviewMeta(reviewRow) || {
      id: null,
      reviewPeriod: normalizedReviewPeriod,
      reviewType: null,
      createdAt: null,
      updatedAt: null,
      createdBy: null,
      updatedBy: null,
    },
    reflection: {
      goingWell: reviewRow?.going_well || "",
      needsAttention: reviewRow?.needs_attention || "",
      nextAction: reviewRow?.next_action || "",
    },
    values: values.map((item) => {
      const reviewData = valueReviewMap.get(item.id);
      return {
        ...item,
        review: reviewData
          ? {
              score: reviewData.score ?? "",
              evidence: reviewData.evidence || "",
              improvementAction: reviewData.improvement_action || "",
              nameSnapshot: reviewData.name_snapshot || "",
            }
          : null,
      };
    }),
    principles: principles.map((item) => {
      const reviewData = principleReviewMap.get(item.id);
      return {
        ...item,
        review: reviewData
          ? {
              score: reviewData.score ?? "",
              exampleDecision: reviewData.example_decision || "",
              adjustmentNeeded: reviewData.adjustment_needed || "",
              nameSnapshot: reviewData.name_snapshot || "",
            }
          : null,
      };
    }),
    objectives: objectives.map((item) => {
      const reviewData = objectiveReviewMap.get(item.id);
      return {
        ...item,
        review: reviewData
          ? {
              owner: reviewData.owner || "",
              progressScore: reviewData.progress_score ?? "",
              risks: reviewData.risks || "",
              titleSnapshot: reviewData.title_snapshot || "",
            }
          : null,
      };
    }),
    goals: goals.map((item) => {
      const reviewData = goalReviewMap.get(item.id);
      const currentObjective = item.objectiveId ? objectiveMap.get(item.objectiveId) : null;
      return {
        ...item,
        objectiveTitle: currentObjective?.title || "",
        review: reviewData
          ? {
              owner: reviewData.owner || "",
              status: reviewData.status || "",
              progressScore: reviewData.progress_score ?? "",
              nextAction: reviewData.next_action || "",
              titleSnapshot: reviewData.title_snapshot || "",
              objectiveIdSnapshot: reviewData.objective_id_snapshot || null,
              objectiveTitleSnapshot: reviewData.objective_title_snapshot || "",
            }
          : null,
      };
    }),
  };
}

async function createReview(reviewPeriod, reflection, email) {
  const rows = await supabaseRestFetch("scorecard_reviews", {
    method: "POST",
    query: { select: "*" },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: [
      {
        review_period: normalizeReviewPeriod(reviewPeriod),
        review_type: String(reflection?.reviewType || "").trim() || null,
        going_well: normalizeNullableText(reflection?.goingWell),
        needs_attention: normalizeNullableText(reflection?.needsAttention),
        next_action: normalizeNullableText(reflection?.nextAction),
        created_by: email || null,
        updated_by: email || null,
      },
    ],
  });

  return Array.isArray(rows) ? rows[0] || null : null;
}

async function updateReview(reviewId, reflection, email) {
  const rows = await supabaseRestFetch("scorecard_reviews", {
    method: "PATCH",
    query: {
      select: "*",
      id: `eq.${ensureUuid(reviewId, "Review")}`,
    },
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=representation",
    },
    body: {
      review_type: String(reflection?.reviewType || "").trim() || null,
      going_well: normalizeNullableText(reflection?.goingWell),
      needs_attention: normalizeNullableText(reflection?.needsAttention),
      next_action: normalizeNullableText(reflection?.nextAction),
      updated_by: email || null,
    },
  });

  return Array.isArray(rows) ? rows[0] || null : null;
}

function hasValueReviewContent(item) {
  return item.score != null || item.evidence || item.improvementAction;
}

function hasPrincipleReviewContent(item) {
  return item.score != null || item.exampleDecision || item.adjustmentNeeded;
}

function hasObjectiveReviewContent(item) {
  return item.progressScore != null || item.owner || item.risks;
}

function hasGoalReviewContent(item) {
  return item.progressScore != null || item.owner || item.status || item.nextAction;
}

async function deleteReviewRows(table, reviewId) {
  await supabaseRestFetch(table, {
    method: "DELETE",
    query: {
      review_id: `eq.${ensureUuid(reviewId, "Review")}`,
    },
    headers: {
      Prefer: "return=minimal",
    },
  });
}

async function insertReviewRows(table, rows) {
  if (!rows.length) {
    return;
  }
  await supabaseRestFetch(table, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Prefer: "return=minimal",
    },
    body: rows,
  });
}

async function saveScorecardReview(reviewPeriod, payload = {}, email = "") {
  const normalizedReviewPeriod = normalizeReviewPeriod(reviewPeriod);
  const reflection = payload?.reflection && typeof payload.reflection === "object" ? payload.reflection : {};
  let reviewRow = await getReviewByPeriod(normalizedReviewPeriod);
  if (!reviewRow) {
    reviewRow = await createReview(normalizedReviewPeriod, reflection, email);
  } else {
    reviewRow = await updateReview(reviewRow.id, reflection, email);
  }

  const reviewId = ensureUuid(reviewRow?.id, "Review");
  const valueInputs = Array.isArray(payload?.values) ? payload.values : [];
  const principleInputs = Array.isArray(payload?.principles) ? payload.principles : [];
  const objectiveInputs = Array.isArray(payload?.objectives) ? payload.objectives : [];
  const goalInputs = Array.isArray(payload?.goals) ? payload.goals : [];

  const [existingValueRows, existingPrincipleRows, existingObjectiveRows, existingGoalRows, valueDefinitions, principleDefinitions, objectiveDefinitions, goalDefinitions] =
    await Promise.all([
      fetchReviewRows("values", reviewId),
      fetchReviewRows("principles", reviewId),
      fetchReviewRows("objectives", reviewId),
      fetchReviewRows("goals", reviewId),
      listDefinitions("values", { ids: valueInputs.map((item) => item.id) }),
      listDefinitions("principles", { ids: principleInputs.map((item) => item.id) }),
      listDefinitions("objectives", { ids: objectiveInputs.map((item) => item.id) }),
      listDefinitions("goals", { ids: goalInputs.map((item) => item.id) }),
    ]);

  const goalObjectiveIds = goalDefinitions.map((item) => item.objectiveId).filter(Boolean);
  const goalObjectiveDefinitions = await listDefinitions("objectives", { ids: goalObjectiveIds });

  const valueMap = new Map(valueDefinitions.map((item) => [item.id, item]));
  const principleMap = new Map(principleDefinitions.map((item) => [item.id, item]));
  const objectiveMap = new Map(
    [...objectiveDefinitions, ...goalObjectiveDefinitions].map((item) => [item.id, item])
  );
  const goalMap = new Map(goalDefinitions.map((item) => [item.id, item]));

  const existingValueMap = new Map(existingValueRows.map((item) => [item.value_id, item]));
  const existingPrincipleMap = new Map(existingPrincipleRows.map((item) => [item.principle_id, item]));
  const existingObjectiveMap = new Map(existingObjectiveRows.map((item) => [item.objective_id, item]));
  const existingGoalMap = new Map(existingGoalRows.map((item) => [item.goal_id, item]));

  const valueRows = valueInputs
    .map((item) => {
      const definition = valueMap.get(ensureUuid(item.id, "Value"));
      if (!definition) {
        throw new Error("Could not find one of the selected values.");
      }
      const normalized = {
        review_id: reviewId,
        value_id: definition.id,
        name_snapshot: existingValueMap.get(definition.id)?.name_snapshot || definition.name || "",
        score: normalizeNullableScore(item.score),
        evidence: normalizeNullableText(item.evidence),
        improvement_action: normalizeNullableText(item.improvementAction),
      };
      return hasValueReviewContent(normalized) ? normalized : null;
    })
    .filter(Boolean);

  const principleRows = principleInputs
    .map((item) => {
      const definition = principleMap.get(ensureUuid(item.id, "Principle"));
      if (!definition) {
        throw new Error("Could not find one of the selected principles.");
      }
      const normalized = {
        review_id: reviewId,
        principle_id: definition.id,
        name_snapshot: existingPrincipleMap.get(definition.id)?.name_snapshot || definition.name || "",
        score: normalizeNullableScore(item.score),
        example_decision: normalizeNullableText(item.exampleDecision),
        adjustment_needed: normalizeNullableText(item.adjustmentNeeded),
      };
      return hasPrincipleReviewContent(normalized) ? normalized : null;
    })
    .filter(Boolean);

  const objectiveRows = objectiveInputs
    .map((item) => {
      const definition = objectiveMap.get(ensureUuid(item.id, "Objective"));
      if (!definition) {
        throw new Error("Could not find one of the selected objectives.");
      }
      const normalized = {
        review_id: reviewId,
        objective_id: definition.id,
        title_snapshot: existingObjectiveMap.get(definition.id)?.title_snapshot || definition.title || "",
        owner: normalizeNullableText(item.owner),
        progress_score: normalizeNullableScore(item.progressScore),
        risks: normalizeNullableText(item.risks),
      };
      return hasObjectiveReviewContent(normalized) ? normalized : null;
    })
    .filter(Boolean);

  const goalRows = goalInputs
    .map((item) => {
      const definition = goalMap.get(ensureUuid(item.id, "Goal"));
      if (!definition) {
        throw new Error("Could not find one of the selected goals.");
      }
      const linkedObjective = definition.objectiveId ? objectiveMap.get(definition.objectiveId) : null;
      const existing = existingGoalMap.get(definition.id);
      const normalized = {
        review_id: reviewId,
        goal_id: definition.id,
        title_snapshot: existing?.title_snapshot || definition.title || "",
        objective_id_snapshot: existing?.objective_id_snapshot || definition.objectiveId || null,
        objective_title_snapshot:
          existing?.objective_title_snapshot || linkedObjective?.title || item.objectiveTitleSnapshot || null,
        owner: normalizeNullableText(item.owner),
        status: normalizeNullableText(item.status),
        progress_score: normalizeNullableScore(item.progressScore),
        next_action: normalizeNullableText(item.nextAction),
      };
      return hasGoalReviewContent(normalized) ? normalized : null;
    })
    .filter(Boolean);

  await Promise.all([
    deleteReviewRows(ENTITY_CONFIG.values.reviewTable, reviewId),
    deleteReviewRows(ENTITY_CONFIG.principles.reviewTable, reviewId),
    deleteReviewRows(ENTITY_CONFIG.objectives.reviewTable, reviewId),
    deleteReviewRows(ENTITY_CONFIG.goals.reviewTable, reviewId),
  ]);

  await Promise.all([
    insertReviewRows(ENTITY_CONFIG.values.reviewTable, valueRows),
    insertReviewRows(ENTITY_CONFIG.principles.reviewTable, principleRows),
    insertReviewRows(ENTITY_CONFIG.objectives.reviewTable, objectiveRows),
    insertReviewRows(ENTITY_CONFIG.goals.reviewTable, goalRows),
  ]);

  return getScorecardView(normalizedReviewPeriod);
}

module.exports = {
  createDefinition,
  getDefinitionById,
  getScorecardView,
  listDefinitions,
  normalizeEntityType,
  normalizeReviewPeriod,
  saveScorecardReview,
  updateDefinition,
};
