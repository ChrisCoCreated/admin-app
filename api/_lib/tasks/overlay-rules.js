const { toUtcIsoOrNull } = require("./unified-model");

function applyOverlayBehaviorRules(existingOverlay, patch, nowIso) {
  const next = {
    ...patch,
  };

  const nextStatus = String(next.workingStatus || "")
    .trim()
    .toLowerCase();

  if (nextStatus === "active") {
    next.lastWorkedAt = nowIso;
    const existingActiveStartedAt = existingOverlay?.activeStartedAt || null;
    if (!existingActiveStartedAt && !next.activeStartedAt) {
      next.activeStartedAt = nowIso;
    } else if (next.activeStartedAt) {
      next.activeStartedAt = toUtcIsoOrNull(next.activeStartedAt);
    }
  }

  if (nextStatus === "parked") {
    next.lastWorkedAt = nowIso;
  }

  return next;
}

module.exports = {
  applyOverlayBehaviorRules,
};
