const { AzureOpenAI } = require("openai");
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

function normalizeApiVersion(value) {
  const apiVersion = normalizeText(value || process.env.AZURE_OPENAI_API_VERSION);
  if (!apiVersion) {
    throw missingConfigError("Server missing AZURE_OPENAI_API_VERSION for Azure OpenAI requests.");
  }
  return apiVersion;
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

function mapThinkingToReasoningEffort(thinking, effort) {
  if (thinking === "disabled") {
    return "none";
  }

  const normalized = normalizeReasoningEffort(effort);
  return normalized === "max" ? "high" : normalized;
}

function buildClient(options = {}) {
  if (options.client) {
    return options.client;
  }
  if (typeof options.clientFactory === "function") {
    return options.clientFactory();
  }

  return new AzureOpenAI({
    endpoint: normalizeEndpoint(options.endpoint),
    apiKey: normalizeApiKey(options.apiKey),
    apiVersion: normalizeApiVersion(options.apiVersion || options.api_version),
    deployment: normalizeDeployment(options.deployment),
    maxRetries: 0,
    timeout: SDK_TIMEOUT_MS,
  });
}

async function createChatCompletion(options = {}) {
  const deployment = resolveDeploymentName(options);
  const thinking = normalizeThinking(options.thinking, null);
  const body = {
    model: deployment,
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
    body.reasoning_effort = mapThinkingToReasoningEffort(
      thinking,
      options.reasoningEffort || options.reasoning_effort
    );
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
    deployment,
  });
  const request = client.chat.completions.create(body, {
    maxRetries: 0,
    timeout: SDK_TIMEOUT_MS,
  });
  const { data: payload, request_id: requestId } = await request.withResponse();

  return {
    payload,
    text: extractResponseText(payload),
    content: payload?.choices?.[0]?.message?.content || "",
    reasoningContent: "",
    usage: payload?.usage || null,
    model: payload?.model || deployment,
    thinking: thinking || null,
    reasoningEffort: thinking ? body.reasoning_effort : null,
    requestId,
  };
}

module.exports = {
  createChatCompletion,
};
