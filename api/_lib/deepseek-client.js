const DEFAULT_BASE_URL = "https://api.deepseek.com";
const DEFAULT_CHAT_MODEL = "deepseek-v4-flash";
const DEFAULT_REASONING_EFFORT = "high";
const SUPPORTED_MODELS = new Set(["deepseek-v4-flash", "deepseek-v4-pro"]);
const LEGACY_MODEL_ALIASES = new Map([
  ["deepseek-chat", { model: "deepseek-v4-flash", thinking: "disabled" }],
  ["deepseek-reasoner", { model: "deepseek-v4-flash", thinking: "enabled" }],
]);

function normalizeText(value) {
  return String(value || "").trim();
}

function normalizeBaseUrl(value) {
  return normalizeText(value).replace(/\/+$/, "") || DEFAULT_BASE_URL;
}

function normalizeModel(value) {
  const requested = normalizeText(value) || DEFAULT_CHAT_MODEL;
  const legacy = LEGACY_MODEL_ALIASES.get(requested);
  if (legacy) {
    return legacy;
  }
  if (!SUPPORTED_MODELS.has(requested)) {
    const error = new Error(
      `Unsupported DeepSeek model "${requested}". Use deepseek-v4-flash or deepseek-v4-pro.`
    );
    error.status = 400;
    throw error;
  }
  return { model: requested, thinking: null };
}

function normalizeThinking(value, fallback = null) {
  if (value === undefined || value === null || value === "") {
    return fallback;
  }
  if (typeof value === "boolean") {
    return value ? "enabled" : "disabled";
  }

  const normalized = normalizeText(value).toLowerCase();
  if (["enabled", "enable", "thinking", "think", "true", "1", "on", "yes"].includes(normalized)) {
    return "enabled";
  }
  if (["disabled", "disable", "no_think", "nothink", "non-thinking", "false", "0", "off", "no"].includes(normalized)) {
    return "disabled";
  }

  const error = new Error('Invalid thinking mode. Use "enabled" or "disabled".');
  error.status = 400;
  throw error;
}

function normalizeReasoningEffort(value) {
  const normalized = normalizeText(value).toLowerCase();
  if (!normalized || ["low", "medium", "high"].includes(normalized)) {
    return "high";
  }
  if (["xhigh", "max"].includes(normalized)) {
    return "max";
  }

  const error = new Error('Invalid reasoning_effort. Use "high" or "max".');
  error.status = 400;
  throw error;
}

function normalizeMaxTokens(value) {
  if (value === undefined || value === null || value === "") {
    return null;
  }
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed <= 0) {
    const error = new Error("max_tokens must be a positive integer.");
    error.status = 400;
    throw error;
  }
  return parsed;
}

function normalizeMessage(message, index) {
  const role = normalizeText(message?.role).toLowerCase();
  if (!["system", "user", "assistant", "tool"].includes(role)) {
    const error = new Error(`Message ${index + 1} has an invalid role.`);
    error.status = 400;
    throw error;
  }

  const normalized = { role };
  if (typeof message?.content === "string") {
    normalized.content = message.content;
  } else if (role !== "assistant") {
    const error = new Error(`Message ${index + 1} must include string content.`);
    error.status = 400;
    throw error;
  }

  const name = normalizeText(message?.name);
  if (name) {
    normalized.name = name;
  }

  if (role === "assistant" && Array.isArray(message?.tool_calls) && message.tool_calls.length > 0) {
    normalized.tool_calls = message.tool_calls;
  }

  if (role === "tool") {
    const toolCallId = normalizeText(message?.tool_call_id);
    if (!toolCallId) {
      const error = new Error(`Tool message ${index + 1} must include tool_call_id.`);
      error.status = 400;
      throw error;
    }
    normalized.tool_call_id = toolCallId;
  }

  return normalized;
}

function normalizeMessages(messages) {
  if (!Array.isArray(messages) || messages.length === 0) {
    const error = new Error("messages must be a non-empty array.");
    error.status = 400;
    throw error;
  }
  return messages.map(normalizeMessage);
}

function extractResponseText(payload) {
  const primary = normalizeText(payload?.choices?.[0]?.message?.content);
  if (primary) {
    return primary;
  }

  if (typeof payload?.output_text === "string" && payload.output_text.trim()) {
    return payload.output_text.trim();
  }

  const outputs = Array.isArray(payload?.output) ? payload.output : [];
  for (const output of outputs) {
    const content = Array.isArray(output?.content) ? output.content : [];
    for (const item of content) {
      if (typeof item?.text === "string" && item.text.trim()) {
        return item.text.trim();
      }
    }
  }

  return "";
}

async function createChatCompletion(options = {}) {
  const apiKey = normalizeText(options.apiKey || process.env.DEEPSEEK_API_KEY);
  if (!apiKey) {
    const error = new Error("Server missing DEEPSEEK_API_KEY for DeepSeek requests.");
    error.status = 500;
    throw error;
  }

  const normalizedModel = normalizeModel(options.model || process.env.DEEPSEEK_MODEL);
  const thinking = normalizeThinking(options.thinking, normalizedModel.thinking);
  const body = {
    model: normalizedModel.model,
    messages: normalizeMessages(options.messages),
  };

  const maxTokens = normalizeMaxTokens(options.maxTokens ?? options.max_tokens);
  if (maxTokens != null) {
    body.max_tokens = maxTokens;
  }

  if (options.responseFormat && typeof options.responseFormat === "object") {
    body.response_format = options.responseFormat;
  }

  if (thinking) {
    body.thinking = { type: thinking };
  }

  if (thinking === "enabled") {
    body.reasoning_effort = normalizeReasoningEffort(options.reasoningEffort || options.reasoning_effort || DEFAULT_REASONING_EFFORT);
  } else {
    for (const [bodyKey, optionKey] of [
      ["temperature", "temperature"],
      ["top_p", "topP"],
      ["presence_penalty", "presencePenalty"],
      ["frequency_penalty", "frequencyPenalty"],
    ]) {
      const value = options[optionKey];
      if (typeof value === "number" && Number.isFinite(value)) {
        body[bodyKey] = value;
      }
    }
  }

  const baseUrl = normalizeBaseUrl(process.env.DEEPSEEK_BASE_URL);
  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify(body),
  });

  const payload = await response.json().catch(() => null);
  if (!response.ok) {
    const detail = payload?.error?.message || `DeepSeek request failed (${response.status}).`;
    const error = new Error(detail);
    error.status = response.status;
    error.payload = payload;
    throw error;
  }

  return {
    payload,
    text: extractResponseText(payload),
    content: payload?.choices?.[0]?.message?.content || "",
    reasoningContent: payload?.choices?.[0]?.message?.reasoning_content || "",
    usage: payload?.usage || null,
    model: payload?.model || normalizedModel.model,
    thinking: thinking || null,
    reasoningEffort: thinking === "enabled" ? body.reasoning_effort : null,
  };
}

module.exports = {
  DEFAULT_CHAT_MODEL,
  createChatCompletion,
  extractResponseText,
  normalizeMaxTokens,
  normalizeMessages,
  normalizeReasoningEffort,
  normalizeText,
  normalizeThinking,
};
