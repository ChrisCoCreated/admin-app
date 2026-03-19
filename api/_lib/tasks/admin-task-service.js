const { createGraphAppClient } = require("../graph-app-client");
const { listTaskSetTemplates } = require("./task-set-source");

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeUserId(value) {
  return normalizeText(value).toLowerCase();
}

function toTodoDueDatePayload(value) {
  if (!value) {
    return null;
  }
  const iso = new Date(value);
  if (Number.isNaN(iso.getTime())) {
    const error = new Error("dueDateTimeUtc must be a valid datetime.");
    error.status = 400;
    error.code = "INVALID_DUE_DATE";
    throw error;
  }
  return {
    dateTime: iso.toISOString(),
    timeZone: "UTC",
  };
}

function inferSource(body) {
  if (Array.isArray(body?.tasks)) {
    return "direct";
  }
  if (body?.taskSet) {
    return "task_set";
  }
  return normalizeText(body?.source).toLowerCase();
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

function addDays(isoDateTimeUtc, days) {
  const parsed = new Date(isoDateTimeUtc);
  parsed.setUTCDate(parsed.getUTCDate() + days);
  return parsed.toISOString();
}

function normalizeTaskInput(task, index, defaultTargetUser) {
  if (!task || typeof task !== "object") {
    const error = new Error(`Task at index ${index} must be an object.`);
    error.status = 400;
    error.code = "INVALID_TASK";
    throw error;
  }

  const title = normalizeText(task.title);
  if (!title) {
    const error = new Error(`Task at index ${index} is missing a title.`);
    error.status = 400;
    error.code = "INVALID_TITLE";
    throw error;
  }

  const targetUser = normalizeUserId(task.targetUser || defaultTargetUser);
  if (!targetUser) {
    const error = new Error(`Task '${title}' is missing targetUser.`);
    error.status = 400;
    error.code = "TARGET_USER_REQUIRED";
    throw error;
  }

  return {
    title,
    description: normalizeText(task.description),
    dueDateTimeUtc: normalizeText(task.dueDateTimeUtc),
    targetUser,
    sourceTaskSet: normalizeText(task.sourceTaskSet),
    sourceItemId: normalizeText(task.sourceItemId),
    sourceResponsiblePerson: normalizeText(task.sourceResponsiblePerson),
    area: normalizeText(task.area),
    dueDateDelay: Number.isFinite(Number(task.dueDateDelay)) ? Number(task.dueDateDelay) : null,
  };
}

async function resolveDefaultTodoListIdForUser(graphClient, targetUser) {
  const lists = await graphClient.fetchAllPages(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(targetUser)}/todo/lists?$top=100`
  );

  if (!Array.isArray(lists) || lists.length === 0) {
    const error = new Error(`No Microsoft To Do list available for user '${targetUser}'.`);
    error.status = 404;
    error.code = "TODO_LIST_NOT_FOUND";
    throw error;
  }

  const preferred =
    lists.find((entry) => normalizeText(entry?.wellknownListName).toLowerCase() === "defaultlist") ||
    lists.find((entry) => normalizeText(entry?.displayName).toLowerCase() === "tasks") ||
    lists[0];
  const listId = normalizeText(preferred?.id);
  if (!listId) {
    const error = new Error(`Could not resolve Microsoft To Do list id for user '${targetUser}'.`);
    error.status = 404;
    error.code = "TODO_LIST_NOT_FOUND";
    throw error;
  }
  return listId;
}

async function createTodoTaskForUser(graphClient, task, cachedListIdsByUser) {
  const targetUser = task.targetUser;
  let listId = cachedListIdsByUser.get(targetUser);
  if (!listId) {
    listId = await resolveDefaultTodoListIdForUser(graphClient, targetUser);
    cachedListIdsByUser.set(targetUser, listId);
  }

  const createdTask = await graphClient.fetchJson(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(targetUser)}/todo/lists/${encodeURIComponent(listId)}/tasks`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        title: task.title,
        ...(task.description
          ? {
              body: {
                contentType: "text",
                content: task.description,
              },
            }
          : {}),
        ...(task.dueDateTimeUtc ? { dueDateTime: toTodoDueDatePayload(task.dueDateTimeUtc) } : {}),
      }),
    }
  );

  return {
    provider: "todo",
    externalTaskId: normalizeText(createdTask?.id),
    externalContainerId: listId,
    title: normalizeText(createdTask?.title) || task.title,
    dueDateTimeUtc: normalizeText(createdTask?.dueDateTime?.dateTime) || task.dueDateTimeUtc || null,
    targetUser,
    description: task.description || "",
    sourceTaskSet: task.sourceTaskSet || "",
    sourceItemId: task.sourceItemId || "",
    sourceResponsiblePerson: task.sourceResponsiblePerson || "",
    area: task.area || "",
    dueDateDelay: task.dueDateDelay,
  };
}

async function createTasksForUserRequest(body = {}) {
  const graphClient = createGraphAppClient();
  const dryRun = body?.dryRun === true;
  const source = inferSource(body);
  let normalizedTasks = [];

  if (source === "direct") {
    const tasks = Array.isArray(body?.tasks) ? body.tasks : [];
    if (!tasks.length) {
      const error = new Error("tasks must be a non-empty array.");
      error.status = 400;
      error.code = "TASKS_REQUIRED";
      throw error;
    }
    normalizedTasks = tasks.map((task, index) => normalizeTaskInput(task, index, body?.targetUser));
  } else if (source === "task_set") {
    const taskSet = normalizeText(body?.taskSet);
    if (!taskSet) {
      const error = new Error("taskSet is required when source is task_set.");
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
    normalizedTasks = templates.map((template, index) =>
      normalizeTaskInput(
        {
          title: template.title,
          description: template.description,
          targetUser: body?.targetUser || template.responsiblePerson,
          dueDateTimeUtc:
            Number(template.dueDateDelay) >= 0 ? addDays(anchorDateTimeUtc, Number(template.dueDateDelay)) : "",
          sourceTaskSet: template.taskSet,
          sourceItemId: template.itemId,
          sourceResponsiblePerson: template.responsiblePerson,
          area: template.area,
          dueDateDelay: template.dueDateDelay,
        },
        index,
        body?.targetUser
      )
    );

    body._taskSetMeta = {
      ...meta,
      anchorDateTimeUtc,
    };
  } else {
    const error = new Error("Unsupported task creation source. Use direct tasks or taskSet.");
    error.status = 400;
    error.code = "UNSUPPORTED_SOURCE";
    throw error;
  }

  const cachedListIdsByUser = new Map();
  if (dryRun) {
    return {
      ok: true,
      dryRun: true,
      tasks: normalizedTasks,
      meta: {
        total: normalizedTasks.length,
        distinctTargetUsers: Array.from(new Set(normalizedTasks.map((task) => task.targetUser))).length,
        source,
        ...(body?._taskSetMeta || {}),
      },
    };
  }

  const createdTasks = [];
  for (const task of normalizedTasks) {
    createdTasks.push(await createTodoTaskForUser(graphClient, task, cachedListIdsByUser));
  }

  return {
    ok: true,
    dryRun: false,
    tasks: createdTasks,
    meta: {
      total: createdTasks.length,
      distinctTargetUsers: Array.from(new Set(createdTasks.map((task) => task.targetUser))).length,
      source,
      ...(body?._taskSetMeta || {}),
    },
  };
}

module.exports = {
  createTasksForUserRequest,
};
