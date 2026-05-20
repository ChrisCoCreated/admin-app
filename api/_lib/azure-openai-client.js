const { OpenAI } = require("openai");
const {
  extractResponseText,
  normalizeMaxTokens,
  normalizeMessages,
  normalizeReasoningEffort,
  normalizeText,
  normalizeThinking,
} = require("./deepseek-client");

const SDK_TIMEOUT_MS = 2147483647;
const FAST_MODEL_NAMES = new Set(["deepseek-v4-flash", "deepseek-chat", "flash", "fast"]);
const PRIMARY_MODEL_NAMES = new Set(["deepseek-v4-pro", "deepseek-reasoner", "pro", "primary"]);

function missingConfigError(message) {
  const error = new Error(message);
  error.status = 500;
  return error;
}

function normalizeEndpoint(value) {
  const endpoint = normalizeText(value || process.env.AZURE_OPENAI_ENDPOINT);
  if (!endpoint) {
    throw missingConfigError("Server missing AZURE_OPENAI_ENDPOINT for Azure OpenAI requests.");
  }
  return endpoint.replace(/\/+$/, "");
}

function endpointHost(endpoint) {
  try {
    return new URL(endpoint).host || null;
  } catch (_error) {
    return normalizeText(endpoint).replace(/^https?:\/\//i, "").split("/")[0] || null;
  }
}

function normalizeApiVersion(value) {
  return normalizeText(value || process.env.AZURE_OPENAI_API_VERSION) || null;
}

function normalizeApiKey(value) {
  const apiKey = normalizeText(value || process.env.AZURE_OPENAI_API_KEY);
  if (!apiKey) {
    throw missingConfigError("Server missing AZURE_OPENAI_API_KEY for Azure OpenAI requests.");
  }
  return apiKey;
}

function normalizeDeployment(value) {
  const deployment = normalizeText(value || process.env.AZURE_OPENAI_DEPLOYMENT_NAME);
  if (!deployment) {
    throw missingConfigError(
      "Server missing Azure deployment config. Set AZURE_OPENAI_DEPLOYMENT_NAME, or tiered AZURE_OPENAI_DEPLOYMENT_PRIMARY / AZURE_OPENAI_DEPLOYMENT_FAST."
    );
  }
  return deployment;
}

function resolveDeploymentName(options = {}) {
  const explicitDeployment = normalizeText(options.deployment);
  if (explicitDeployment) {
    return explicitDeployment;
  }

  const requestedModel = normalizeText(options.model).toLowerCase();
  if (FAST_MODEL_NAMES.has(requestedModel)) {
    const fastDeployment = normalizeText(process.env.AZURE_OPENAI_DEPLOYMENT_FAST);
    if (fastDeployment) {
      return fastDeployment;
    }
  }

  if (PRIMARY_MODEL_NAMES.has(requestedModel)) {
    const primaryDeployment = normalizeText(process.env.AZURE_OPENAI_DEPLOYMENT_PRIMARY);
    if (primaryDeployment) {
      return primaryDeployment;
    }
  }

  const defaultPrimaryDeployment = normalizeText(process.env.AZURE_OPENAI_DEPLOYMENT_PRIMARY);
  if (defaultPrimaryDeployment) {
    return defaultPrimaryDeployment;
  }

  return normalizeDeployment(options.deployment);
}

function buildAzureRouteMetadata({ endpoint, apiVersion, deployment, requestedModel }) {
  return {
    provider: "azure_openai",
    requestedModel: normalizeText(requestedModel) || null,
    deployment,
    apiVersion,
    endpointHost: endpointHost(endpoint),
  };
}

function buildFoundryBaseUrl(endpoint) {
  const normalized = normalizeEndpoint(endpoint);
  if (/\/openai\/v1\/?$/i.test(normalized)) {
    return normalized.replace(/\/+$/, "");
  }
  return `${normalized.replace(/\/+$/, "")}/openai/v1`;
}

function resolveAzureRouteMetadata(options = {}) {
  const deployment = resolveDeploymentName(options);
  const endpoint = normalizeEndpoint(options.endpoint);
  return buildAzureRouteMetadata({
    endpoint,
    apiVersion: null,
    deployment,
    requestedModel: options.model,
  });
}

function mapThinkingToReasoningEffort(thinking, effort) {
  if (thinking === "disabled") {
    return "none";
  }

  const normalized = normalizeText(effort).toLowerCase();
  if (!normalized) {
    return "medium";
  }
  if (["low", "medium", "high"].includes(normalized)) {
    return normalized;
  }
  if (["max", "xhigh"].includes(normalized)) {
    return "high";
  }

  const error = new Error('Invalid Azure reasoning effort. Use "low", "medium", "high", or "max".');
  error.status = 400;
  throw error;
}

function supportsAzureReasoning(options = {}, aiRoute = {}) {
  const requestedModel = normalizeText(options.model || aiRoute.requestedModel).toLowerCase();
  if (FAST_MODEL_NAMES.has(requestedModel)) {
    return false;
  }
  return true;
}

function buildClient(options = {}) {
  if (options.client) {
    return options.client;
  }
  if (typeof options.clientFactory === "function") {
    return options.clientFactory();
  }

  return new OpenAI({
    baseURL: buildFoundryBaseUrl(options.endpoint),
    apiKey: normalizeApiKey(options.apiKey),
    maxRetries: 0,
    timeout: SDK_TIMEOUT_MS,
  });
}

function buildResponsesInput(messages) {
  const normalizedMessages = normalizeMessages(messages);
  const instructions = normalizedMessages
    .filter((message) => message.role === "system" || message.role === "developer")
    .map((message) => normalizeText(message.content))
    .filter(Boolean)
    .join("\n\n");
  const input = normalizedMessages
    .filter((message) => message.role !== "system" && message.role !== "developer")
    .map((message) => ({
      role: message.role === "assistant" ? "assistant" : "user",
      content: normalizeText(message.content),
    }));

  return {
    instructions: instructions || undefined,
    input: input.length > 0 ? input : normalizeText(normalizedMessages[0]?.content),
  };
}

function extractResponsesText(payload) {
  const outputText = normalizeText(payload?.output_text);
  if (outputText) {
    return outputText;
  }

  const parts = [];
  for (const item of payload?.output || []) {
    if (item?.type !== "message") {
      continue;
    }
    for (const content of item.content || []) {
      if (content?.type === "output_text" && typeof content.text === "string") {
        parts.push(content.text);
      }
    }
  }
  return parts.join("").trim();
}

async function createChatCompletion(options = {}) {
  const aiRoute = resolveAzureRouteMetadata(options);
  const deployment = aiRoute.deployment;
  const thinking = normalizeThinking(options.thinking, null);
  const responsesInput = buildResponsesInput(options.messages);
  const body = {
    model: deployment,
    input: responsesInput.input,
  };
  if (responsesInput.instructions) {
    body.instructions = responsesInput.instructions;
  }

  const maxTokens = normalizeMaxTokens(options.maxTokens ?? options.max_tokens);
  if (maxTokens != null) {
    body.max_output_tokens = maxTokens;
  }

  if (options.responseFormat && typeof options.responseFormat === "object") {
    body.text = { format: options.responseFormat };
  }

  if (thinking) {
    const effort = mapThinkingToReasoningEffort(thinking, options.reasoningEffort || options.reasoning_effort);
    if (effort !== "none" && supportsAzureReasoning(options, aiRoute)) {
      body.reasoning = { effort };
    }
  }

  if (thinking !== "enabled") {
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

  const client = buildClient({
    ...options,
    endpoint: options.endpoint,
    deployment,
  });
  let payload;
  let requestId;
  try {
    const request = client.responses.create(body, {
      maxRetries: 0,
      timeout: SDK_TIMEOUT_MS,
    });
    const response = await request.withResponse();
    payload = response.data;
    requestId = response.request_id;
  } catch (error) {
    error.aiRoute = aiRoute;
    throw error;
  }

  return {
    payload,
    text: extractResponsesText(payload) || extractResponseText(payload),
    content: extractResponsesText(payload) || "",
    reasoningContent: "",
    usage: payload?.usage || null,
    model: payload?.model || deployment,
    thinking: thinking || null,
    reasoningEffort: body.reasoning?.effort || null,
    requestId,
    aiRoute,
  };
}

module.exports = {
  createChatCompletion,
  resolveAzureRouteMetadata,
};
