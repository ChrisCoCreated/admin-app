const { buildTaskKey } = require("./unified-model");

function mergeTasksWithOverlays(tasks, overlaysByKey) {
  let overlayMatchedCount = 0;
  const taskKeySet = new Set();

  const merged = tasks.map((task) => {
    const key = buildTaskKey(task.provider, task.externalTaskId);
    taskKeySet.add(key);
    const overlay = overlaysByKey.get(key);

    if (!overlay) {
      return {
        ...task,
      };
    }

    overlayMatchedCount += 1;
    return {
      ...task,
      overlay: {
        itemId: overlay.itemId,
        workingStatus: overlay.workingStatus,
        workType: overlay.workType,
        tags: overlay.tags,
        activeStartedAt: overlay.activeStartedAt,
        lastWorkedAt: overlay.lastWorkedAt,
        energy: overlay.energy,
        effortMinutes: overlay.effortMinutes,
        impact: overlay.impact,
        overlayNotes: overlay.overlayNotes,
        pinned: overlay.pinned,
        layout: overlay.layout,
        category: overlay.category,
        lastOverlayUpdatedAt: overlay.lastOverlayUpdatedAt,
      },
    };
  });

  let overlayOrphanCount = 0;
  for (const key of overlaysByKey.keys()) {
    if (!taskKeySet.has(key)) {
      overlayOrphanCount += 1;
    }
  }

  return {
    tasks: merged,
    overlayMatchedCount,
    overlayOrphanCount,
  };
}

function compareNullableDateAsc(a, b) {
  if (!a && !b) {
    return 0;
  }
  if (!a) {
    return 1;
  }
  if (!b) {
    return -1;
  }

  const aTime = Date.parse(a);
  const bTime = Date.parse(b);

  if (!Number.isFinite(aTime) && !Number.isFinite(bTime)) {
    return 0;
  }
  if (!Number.isFinite(aTime)) {
    return 1;
  }
  if (!Number.isFinite(bTime)) {
    return -1;
  }

  return aTime - bTime;
}

function sortUnifiedTasks(tasks) {
  return [...tasks].sort((a, b) => {
    const aPinned = a?.overlay?.pinned === true;
    const bPinned = b?.overlay?.pinned === true;
    if (aPinned !== bPinned) {
      return aPinned ? -1 : 1;
    }

    const aActive = String(a?.overlay?.workingStatus || "").toLowerCase() === "active";
    const bActive = String(b?.overlay?.workingStatus || "").toLowerCase() === "active";
    if (aActive !== bActive) {
      return aActive ? -1 : 1;
    }

    const dueCmp = compareNullableDateAsc(a?.dueDateTimeUtc, b?.dueDateTimeUtc);
    if (dueCmp !== 0) {
      return dueCmp;
    }

    const titleCmp = String(a?.title || "").localeCompare(String(b?.title || ""), undefined, {
      sensitivity: "base",
    });
    if (titleCmp !== 0) {
      return titleCmp;
    }

    return String(a?.externalTaskId || "").localeCompare(String(b?.externalTaskId || ""));
  });
}

module.exports = {
  mergeTasksWithOverlays,
  sortUnifiedTasks,
};
