const { toUtcIsoOrNull } = require("./unified-model");

const PLANNER_PAGE_SIZE = 200;

function normalizePlannerTask(task) {
  const completedDateTimeUtc = toUtcIsoOrNull(task?.completedDateTime);
  const percentComplete = Number(task?.percentComplete || 0);

  return {
    provider: "planner",
    externalTaskId: String(task?.id || "").trim(),
    externalContainerId: String(task?.planId || "").trim(),
    title: String(task?.title || "").trim(),
    dueDateTimeUtc: toUtcIsoOrNull(task?.dueDateTime),
    isCompleted: Boolean(completedDateTimeUtc) || percentComplete >= 100,
    completedDateTimeUtc,
    source: {
      rawId: String(task?.id || "").trim(),
      webUrl: undefined,
      etag: String(task?.["@odata.etag"] || "").trim() || undefined,
    },
  };
}

async function fetchPlannerTasks(graphClient) {
  const url =
    `https://graph.microsoft.com/v1.0/me/planner/tasks` +
    `?$select=id,title,planId,dueDateTime,completedDateTime,percentComplete&$top=${PLANNER_PAGE_SIZE}`;

  const tasks = await graphClient.fetchAllPages(url);
  return tasks
    .map((task) => normalizePlannerTask(task))
    .filter((task) => Boolean(task.externalTaskId));
}

module.exports = {
  fetchPlannerTasks,
};
