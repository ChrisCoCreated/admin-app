const ALLOWED_BLOCK_TAGS = new Set(["P", "BR", "STRONG", "B", "EM", "I", "UL", "OL", "LI", "A"]);

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeHref(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  const lower = text.toLowerCase();
  if (lower.startsWith("http://") || lower.startsWith("https://") || lower.startsWith("mailto:")) {
    return text;
  }

  return "";
}

function sanitizeLightHtml(input) {
  const source = String(input || "").trim();
  if (!source) {
    return "<p></p>";
  }

  const stripped = source
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, "")
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, "");

  const wrapper = { innerHTML: stripped };
  // Lightweight sanitizer for simple rich text persisted in SharePoint.
  const tokenized = String(wrapper.innerHTML || "").split(/(<[^>]+>)/g);

  const stack = [];
  let output = "";

  for (const token of tokenized) {
    if (!token) {
      continue;
    }

    if (!token.startsWith("<")) {
      output += escapeHtml(token);
      continue;
    }

    const closeMatch = /^<\s*\/\s*([a-z0-9]+)\s*>$/i.exec(token);
    if (closeMatch) {
      const tag = closeMatch[1].toUpperCase();
      if (!ALLOWED_BLOCK_TAGS.has(tag) || tag === "BR") {
        continue;
      }
      while (stack.length > 0) {
        const openTag = stack.pop();
        output += `</${openTag.toLowerCase()}>`;
        if (openTag === tag) {
          break;
        }
      }
      continue;
    }

    const openMatch = /^<\s*([a-z0-9]+)([^>]*)>$/i.exec(token);
    if (!openMatch) {
      continue;
    }

    const tag = openMatch[1].toUpperCase();
    if (!ALLOWED_BLOCK_TAGS.has(tag)) {
      continue;
    }

    if (tag === "BR") {
      output += "<br>";
      continue;
    }

    if (tag === "A") {
      const hrefMatch = /\shref\s*=\s*("([^"]*)"|'([^']*)'|([^\s>]+))/i.exec(openMatch[2] || "");
      const href = normalizeHref(hrefMatch?.[2] || hrefMatch?.[3] || hrefMatch?.[4] || "");
      if (!href) {
        continue;
      }
      output += `<a href="${escapeHtml(href)}" target="_blank" rel="noreferrer noopener">`;
      stack.push(tag);
      continue;
    }

    output += `<${tag.toLowerCase()}>`;
    stack.push(tag);
  }

  while (stack.length > 0) {
    const tag = stack.pop();
    output += `</${tag.toLowerCase()}>`;
  }

  const normalized = output.trim();
  if (!normalized) {
    return "<p></p>";
  }

  if (!/^<(p|ul|ol|h1|h2|h3|li|strong|b|em|i|a|br)/i.test(normalized)) {
    return `<p>${normalized}</p>`;
  }

  return normalized;
}

module.exports = {
  sanitizeLightHtml,
};
