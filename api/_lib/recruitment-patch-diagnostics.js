function getTextLengths(fields = {}) {
  return Object.fromEntries(
    Object.entries(fields)
      .filter(([, value]) => typeof value === "string")
      .map(([key, value]) => [key, value.length])
  );
}

function getBooleanFields(fields = {}) {
  return Object.fromEntries(Object.entries(fields).filter(([, value]) => typeof value === "boolean"));
}

function getPresence(fields = {}) {
  return Object.fromEntries(Object.entries(fields).map(([key, value]) => [key, value !== undefined && value !== null && value !== ""]));
}

function logRecruitmentPatchFailure(endpoint, error, details = {}) {
  const fields = details.fields && typeof details.fields === "object" ? details.fields : {};
  console.warn("[recruitment-save] Graph write failed", {
    endpoint,
    operation: details.operation || "write",
    itemId: details.itemId || "",
    status: error?.status || "",
    graphCode: error?.code || "",
    graphMessage: error?.message || "",
    fieldNames: Object.keys(fields),
    textLengths: getTextLengths(fields),
    booleanFields: getBooleanFields(fields),
    present: getPresence(fields),
  });
}

module.exports = {
  logRecruitmentPatchFailure,
};
