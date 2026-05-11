const SINGLE_LINE_TEXT_MAX_LENGTH = 255;

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeIndeedUrl(value) {
  const raw = normalizeText(value);
  if (!raw) {
    return "";
  }

  let normalized = raw;
  try {
    const parsed = new URL(raw);
    if (parsed.protocol === "http:" || parsed.protocol === "https:") {
      parsed.search = "";
      parsed.hash = "";
      normalized = parsed.toString();
    }
  } catch {
    normalized = raw;
  }

  if (normalized.length <= SINGLE_LINE_TEXT_MAX_LENGTH) {
    return normalized;
  }
  return normalized.slice(0, SINGLE_LINE_TEXT_MAX_LENGTH);
}

module.exports = {
  normalizeIndeedUrl,
};
