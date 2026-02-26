const { toUtcIsoOrNull } = require("./unified-model");

const TODO_LISTS_PAGE_SIZE = 100;
const TODO_TASKS_PAGE_SIZE = 200;
const TODO_LIST_CONCURRENCY = 4;

function isInvalidRequestError(error) {
  const code = String(error?.code || "").toLowerCase();
  return Number(error?.status) === 400 && code.includes("invalidrequest");
}

function mapLimit(items, concurrency, mapper) {
  return new Promise((resolve, reject) => {
    const results = new Array(items.length);
    let nextIndex = 0;
    let active = 0;
    let done = 0;

    function launchNext() {
      if (done >= items.length) {
        resolve(results);
        return;
      }

      while (active < concurrency && nextIndex < items.length) {
        const currentIndex = nextIndex;
        nextIndex += 1;
        active += 1;

        Promise.resolve(mapper(items[currentIndex], currentIndex))
          .then((value) => {
            results[currentIndex] = value;
            active -= 1;
            done += 1;
            launchNext();
          })
          .catch(reject);
      }
    }

    if (!items.length) {
      resolve([]);
      return;
    }

    launchNext();
  });
}

function normalizeTodoTask(task, listId) {
  const status = String(task?.status || "").trim().toLowerCase();
  const completedDateTimeUtc = toUtcIsoOrNull(task?.completedDateTime?.dateTime);

  return {
    provider: "todo",
    externalTaskId: String(task?.id || "").trim(),
    externalContainerId: String(listId || "").trim(),
    title: String(task?.title || "").trim(),
    dueDateTimeUtc: toUtcIsoOrNull(task?.dueDateTime?.dateTime),
    isCompleted: status === "completed" || Boolean(completedDateTimeUtc),
    completedDateTimeUtc,
    source: {
      rawId: String(task?.id || "").trim(),
      webUrl: String(task?.webLink || "").trim() || undefined,
      etag: String(task?.["@odata.etag"] || "").trim() || undefined,
    },
  };
}

async function fetchTodoTasks(graphClient) {
  let lists = [];
  try {
    const listsUrl = `https://graph.microsoft.com/v1.0/me/todo/lists?$select=id,displayName&$top=${TODO_LISTS_PAGE_SIZE}`;
    lists = await graphClient.fetchAllPages(listsUrl);
  } catch (error) {
    if (!isInvalidRequestError(error)) {
      throw error;
    }
    // Some tenants/mailboxes reject selected field shape; retry with the broad endpoint.
    const fallbackListsUrl = `https://graph.microsoft.com/v1.0/me/todo/lists?$top=${TODO_LISTS_PAGE_SIZE}`;
    lists = await graphClient.fetchAllPages(fallbackListsUrl);
  }

  const taskBatches = await mapLimit(lists, TODO_LIST_CONCURRENCY, async (list) => {
    const listId = String(list?.id || "").trim();
    if (!listId) {
      return [];
    }

    let tasks = [];
    try {
      const tasksUrl =
        `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}` +
        `/tasks?$select=id,title,status,dueDateTime,completedDateTime,bodyLastModifiedDateTime,webLink&$top=${TODO_TASKS_PAGE_SIZE}`;
      tasks = await graphClient.fetchAllPages(tasksUrl);
    } catch (error) {
      if (!isInvalidRequestError(error)) {
        throw error;
      }
      const fallbackTasksUrl =
        `https://graph.microsoft.com/v1.0/me/todo/lists/${encodeURIComponent(listId)}` +
        `/tasks?$top=${TODO_TASKS_PAGE_SIZE}`;
      tasks = await graphClient.fetchAllPages(fallbackTasksUrl);
    }

    return tasks
      .map((task) => normalizeTodoTask(task, listId))
      .filter((task) => Boolean(task.externalTaskId));
  });

  return taskBatches.flat();
}

module.exports = {
  fetchTodoTasks,
};
