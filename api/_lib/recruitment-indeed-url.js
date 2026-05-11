function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeIndeedUrl(value) {
  return normalizeText(value);
}

module.exports = {
  normalizeIndeedUrl,
};
