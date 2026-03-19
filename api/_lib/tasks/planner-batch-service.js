const { createGraphAppClient } = require("../graph-app-client");
const { listTaskSetTemplates } = require("./task-set-source");

const DEFAULT_PLANNER_TEST_PLAN_ID = "K9YRrHpDOE-uRMWoqZaegpcAA3bl";
const DEFAULT_PLANNER_TEST_BUCKET_NAME = "Task Creation Test";

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeEmail(value) {
  return normalizeText(value).toLowerCase();
}

function toIsoAtUtcNoon(rawValue) {
  const raw = normalizeText(rawValue);
  if (!raw) {
    const now = new Date();
    return new Date(
      Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate(), 12, 0, 0, 0)
    ).toISOString();
  }

  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) {
    return `${raw}T12:00:00.000Z`;
  }

  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) {
    const error = new Error("anchorDate must be a valid ISO date or datetime.");
    error.status = 400;
    error.code = "INVALID_ANCHOR_DATE";
    throw error;
  }

  return parsed.toISOString();
}

function addDays(isoValue, days) {
  const parsed = new Date(isoValue);
  parsed.setUTCDate(parsed.getUTCDate() + days);
  return parsed.toISOString();
}

function buildValidationError(code, message) {
  return { code, message };
}

function getPlannerTestConfig() {
  return {
    planId: normalizeText(process.env.PLANNER_TEST_PLAN_ID || DEFAULT_PLANNER_TEST_PLAN_ID),
    bucketName: normalizeText(process.env.PLANNER_TEST_BUCKET_NAME || DEFAULT_PLANNER_TEST_BUCKET_NAME),
  };
}

async function resolvePlannerBucket(graphClient, planId, bucketName) {
  const buckets = await graphClient.fetchAllPages(
    `https://graph.microsoft.com/v1.0/planner/plans/${encodeURIComponent(planId)}/buckets?$top=200`
  );
  const bucket = buckets.find(
    (entry) => normalizeText(entry?.name).toLowerCase() === bucketName.toLowerCase()
  );

  if (!bucket?.id) {
    const error = new Error(`Planner bucket '${bucketName}' was not found in the configured test plan.`);
    error.status = 404;
    error.code = "PLANNER_BUCKET_NOT_FOUND";
    throw error;
  }

  return {
    bucketId: normalizeText(bucket.id),
    bucketName: normalizeText(bucket.name) || bucketName,
  };
}

async function resolveUser(graphClient, targetUser, userCache) {
  const key = normalizeEmail(targetUser);
  if (userCache.has(key)) {
    return userCache.get(key);
  }

  const payload = await graphClient.fetchJson(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(key)}?$select=id,userPrincipalName,mail,displayName`
  );
  const resolved = {
    id: normalizeText(payload?.id),
    userPrincipalName: normalizeEmail(payload?.userPrincipalName || payload?.mail || key),
    displayName: normalizeText(payload?.displayName),
  };

  if (!resolved.id) {
    const error = new Error(`Could not resolve Planner assignee '${targetUser}'.`);
    error.status = 404;
    error.code = "PLANNER_ASSIGNEE_NOT_FOUND";
    throw error;
  }

  userCache.set(key, resolved);
  return resolved;
}

function normalizeDirectTask(task, index) {
  const title = normalizeText(task?.title);
  const targetUser = normalizeEmail(task?.targetUser);
  const description = normalizeText(task?.description);
  const dueDateTimeUtc = task?.dueDateTimeUtc ? normalizeText(task.dueDateTimeUtc) : null;
  const errors = [];

  if (!title) {
    errors.push(buildValidationError("INVALID_TITLE", `Task at index ${index} is missing a title.`));
  }
  if (!targetUser) {
    errors.push(buildValidationError("TARGET_USER_REQUIRED", `Task '${title || `#${index + 1}`}' is missing targetUser.`));
  }
  if (dueDateTimeUtc) {
    const parsed = new Date(dueDateTimeUtc);
    if (Number.isNaN(parsed.getTime())) {
      errors.push(
        buildValidationError("INVALID_DUE_DATE", `Task '${title || `#${index + 1}`}' has an invalid dueDateTimeUtc.`)
      );
    }
  }

  return {
    index,
    title,
    description,
    targetUser,
    dueDateTimeUtc,
    sourceTaskSet: normalizeText(task?.sourceTaskSet),
    sourceItemId: normalizeText(task?.sourceItemId),
    sourceResponsiblePerson: normalizeText(task?.sourceResponsiblePerson),
    area: normalizeText(task?.area),
    dueDateDelay: Number.isFinite(Number(task?.dueDateDelay)) ? Number(task.dueDateDelay) : null,
    valid: errors.length === 0,
    errors,
  };
}

async function expandTaskSet(body) {
  const taskSet = normalizeText(body?.taskSet);
  if (!taskSet) {
    const error = new Error("taskSet is required when using task set expansion.");
    error.status = 400;
    error.code = "TASK_SET_REQUIRED";
    throw error;
  }

  const { templates, meta } = await listTaskSetTemplates({
    taskSet,
    area: body?.area,
  });
  if (!templates.length) {
    const error = new Error(`No task set templates found for '${taskSet}'.`);
    error.status = 404;
    error.code = "TASK_SET_NOT_FOUND";
    throw error;
  }

  const anchorDateTimeUtc = toIsoAtUtcNoon(body?.anchorDate);
  const tasks = templates.map((template, index) => {
    const dueDateDelay = Number(template?.dueDateDelay);
    return normalizeDirectTask(
      {
        title: template.title,
        description: template.description,
        targetUser: template.responsiblePerson,
        dueDateTimeUtc: Number.isFinite(dueDateDelay) && dueDateDelay >= 0 ? addDays(anchorDateTimeUtc, dueDateDelay) : null,
        sourceTaskSet: template.taskSet,
        sourceItemId: template.itemId,
        sourceResponsiblePerson: template.responsiblePerson,
        area: template.area,
        dueDateDelay: template.dueDateDelay,
      },
      index
    );
  });

  return {
    tasks,
    meta: {
      ...meta,
      anchorDateTimeUtc,
      source: "task_set",
    },
  };
}

async function expandDirectTasks(body) {
  const rawTasks = Array.isArray(body?.tasks) ? body.tasks : [];
  if (!rawTasks.length) {
    const error = new Error("tasks must be a non-empty array.");
    error.status = 400;
    error.code = "TASKS_REQUIRED";
    throw error;
  }

  return {
    tasks: rawTasks.map((task, index) => normalizeDirectTask(task, index)),
    meta: {
      source: "direct",
    },
  };
}

async function fetchPlannerTaskDetails(graphClient, taskId) {
  return graphClient.fetchJson(
    `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(taskId)}/details`
  );
}

async function updatePlannerTaskDescription(graphClient, taskId, description) {
  if (!normalizeText(description)) {
    return null;
  }

  const details = await fetchPlannerTaskDetails(graphClient, taskId);
  const etag = normalizeText(details?.["@odata.etag"]);
  if (!etag) {
    const error = new Error(`Could not resolve Planner task details ETag for task '${taskId}'.`);
    error.status = 502;
    error.code = "PLANNER_DETAILS_ETAG_MISSING";
    throw error;
  }

  return graphClient.fetchJson(
    `https://graph.microsoft.com/v1.0/planner/tasks/${encodeURIComponent(taskId)}/details`,
    {
      method: "PATCH",
      headers: {
        "Content-Type": "application/json",
        "If-Match": etag,
        Prefer: "return=representation",
      },
      body: JSON.stringify({
        description,
        previewType: "description",
      }),
    }
  );
}

async function createPlannerTask(graphClient, task, plannerContext, userCache) {
  const resolvedUser = await resolveUser(graphClient, task.targetUser, userCache);
  const createdTask = await graphClient.fetchJson("https://graph.microsoft.com/v1.0/planner/tasks", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      planId: plannerContext.planId,
      bucketId: plannerContext.bucketId,
      title: task.title,
      ...(task.dueDateTimeUtc ? { dueDateTime: task.dueDateTimeUtc } : {}),
      assignments: {
        [resolvedUser.id]: {
          "@odata.type": "#microsoft.graph.plannerAssignment",
          orderHint: " !",
        },
      },
    }),
  });

  if (task.description) {
    await updatePlannerTaskDescription(graphClient, normalizeText(createdTask?.id), task.description);
  }

  return {
    provider: "planner",
    externalTaskId: normalizeText(createdTask?.id),
    externalContainerId: plannerContext.bucketId,
    planId: plannerContext.planId,
    bucketId: plannerContext.bucketId,
    bucketName: plannerContext.bucketName,
    targetUser: resolvedUser.userPrincipalName,
    targetUserId: resolvedUser.id,
    title: normalizeText(createdTask?.title) || task.title,
    description: task.description || "",
    dueDateTimeUtc: task.dueDateTimeUtc || null,
    sourceTaskSet: task.sourceTaskSet || "",
    sourceItemId: task.sourceItemId || "",
    sourceResponsiblePerson: task.sourceResponsiblePerson || "",
    area: task.area || "",
    dueDateDelay: task.dueDateDelay,
  };
}

async function createPlannerBatch(body = {}) {
  const provider = normalizeText(body?.provider || "planner").toLowerCase();
  if (provider !== "planner") {
    const error = new Error("This batch endpoint currently supports provider='planner' only.");
    error.status = 400;
    error.code = "UNSUPPORTED_PROVIDER";
    throw error;
  }

  const expansion = Array.isArray(body?.tasks) ? await expandDirectTasks(body) : await expandTaskSet(body);
  const graphClient = createGraphAppClient();
  const config = getPlannerTestConfig();
  const plannerBucket = await resolvePlannerBucket(graphClient, config.planId, config.bucketName);
  const plannerContext = {
    planId: config.planId,
    bucketId: plannerBucket.bucketId,
    bucketName: plannerBucket.bucketName,
  };

  const validTasks = expansion.tasks.filter((task) => task.valid);
  const invalidTasks = expansion.tasks.filter((task) => !task.valid);
  const invalidErrors = invalidTasks.map((task) => ({
    index: task.index,
    title: task.title || "",
    targetUser: task.targetUser || "",
    errors: task.errors,
  }));

  if (expansion.tasks.length === 0) {
    const error = new Error("No tasks were available to expand.");
    error.status = 400;
    error.code = "NO_EXPANDED_TASKS";
    throw error;
  }

  if (body?.dryRun === true) {
    return {
      ok: invalidErrors.length === 0,
      dryRun: true,
      provider: "planner",
      tasks: validTasks.map((task) => ({
        provider: "planner",
        planId: plannerContext.planId,
        bucketId: plannerContext.bucketId,
        bucketName: plannerContext.bucketName,
        title: task.title,
        description: task.description,
        targetUser: task.targetUser,
        dueDateTimeUtc: task.dueDateTimeUtc,
        sourceTaskSet: task.sourceTaskSet,
        sourceItemId: task.sourceItemId,
        sourceResponsiblePerson: task.sourceResponsiblePerson,
        area: task.area,
        dueDateDelay: task.dueDateDelay,
      })),
      errors: invalidErrors,
      meta: {
        total: expansion.tasks.length,
        validCount: validTasks.length,
        invalidCount: invalidErrors.length,
        planId: plannerContext.planId,
        bucketId: plannerContext.bucketId,
        bucketName: plannerContext.bucketName,
        ...expansion.meta,
      },
    };
  }

  const createdTasks = [];
  const createErrors = [...invalidErrors];
  const userCache = new Map();

  for (const task of validTasks) {
    try {
      createdTasks.push(await createPlannerTask(graphClient, task, plannerContext, userCache));
    } catch (error) {
      createErrors.push({
        index: task.index,
        title: task.title,
        targetUser: task.targetUser,
        errors: [
          buildValidationError(
            String(error?.code || "PLANNER_TASK_CREATE_FAILED"),
            error?.message || "Could not create Planner task."
          ),
        ],
      });
    }
  }

  return {
    ok: createErrors.length === 0,
    dryRun: false,
    provider: "planner",
    tasks: createdTasks,
    errors: createErrors,
    meta: {
      total: expansion.tasks.length,
      createdCount: createdTasks.length,
      errorCount: createErrors.length,
      planId: plannerContext.planId,
      bucketId: plannerContext.bucketId,
      bucketName: plannerContext.bucketName,
      ...expansion.meta,
    },
  };
}

module.exports = {
  createPlannerBatch,
};
