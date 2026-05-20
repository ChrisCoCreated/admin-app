const { createChatCompletion: createDeepSeekChatCompletion, normalizeText } = require("./deepseek-client");
const { createChatCompletion: createAzureOpenAiChatCompletion } = require("./azure-openai-client");

const DEFAULT_AI_PROVIDER = "deepseek";
const SUPPORTED_AI_PROVIDERS = new Set(["deepseek", "azure_openai"]);

function getConfiguredAiProvider() {
  const provider = normalizeText(process.env.AI_PROVIDER).toLowerCase();
  return provider || DEFAULT_AI_PROVIDER;
}

function createConfigError(message) {
  const error = new Error(message);
  error.status = 500;
  return error;
}

function validateConfiguredAiProvider() {
  const provider = getConfiguredAiProvider();
  if (!SUPPORTED_AI_PROVIDERS.has(provider)) {
    throw createConfigError(`Unsupported AI_PROVIDER "${provider}". Use deepseek or azure_openai.`);
  }

  if (provider !== "azure_openai") {
    return provider;
  }

  const requiredKeys = [
    "AZURE_OPENAI_ENDPOINT",
    "AZURE_OPENAI_API_KEY",
    "AZURE_OPENAI_API_VERSION",
  ];
  const missing = requiredKeys.filter((key) => !normalizeText(process.env[key]));
  if (missing.length > 0) {
    throw createConfigError(`Missing Azure OpenAI configuration: ${missing.join(", ")}.`);
  }

  const hasSingleDeployment = Boolean(normalizeText(process.env.AZURE_OPENAI_DEPLOYMENT_NAME));
  const hasTieredDeployment =
    Boolean(normalizeText(process.env.AZURE_OPENAI_DEPLOYMENT_PRIMARY)) &&
    Boolean(normalizeText(process.env.AZURE_OPENAI_DEPLOYMENT_FAST));
  if (!hasSingleDeployment && !hasTieredDeployment) {
    throw createConfigError(
      "Missing Azure OpenAI deployment configuration: set AZURE_OPENAI_DEPLOYMENT_NAME, or both AZURE_OPENAI_DEPLOYMENT_PRIMARY and AZURE_OPENAI_DEPLOYMENT_FAST."
    );
  }

  return provider;
}

function logAiProviderError(provider, error) {
  const requestId = normalizeText(error?.requestId || error?.requestID);
  console.error("[ai-provider] request failed", {
    provider,
    requestId: requestId || null,
    status: Number(error?.status) || null,
    code: normalizeText(error?.code) || null,
  });
}

function mapAzureOpenAiError(error) {
  const status = Number(error?.status) || 502;
  const mapped = new Error("Azure OpenAI request failed.");
  mapped.status = status;
  mapped.requestId = normalizeText(error?.requestId || error?.requestID) || null;
  mapped.code = normalizeText(error?.code) || null;

  if (status === 400 || status === 422) {
    mapped.message = "Azure OpenAI rejected the AI request.";
    return mapped;
  }
  if (status === 401 || status === 403) {
    mapped.message = "Azure OpenAI authentication failed.";
    return mapped;
  }
  if (status === 404) {
    mapped.message = "Azure OpenAI deployment not found.";
    return mapped;
  }
  if (status === 429) {
    mapped.message = "Azure OpenAI rate limit exceeded. Try again soon.";
    return mapped;
  }
  if (status >= 500) {
    mapped.message = "Azure OpenAI is temporarily unavailable.";
    return mapped;
  }
  if (!error?.status) {
    mapped.message = "Could not reach Azure OpenAI.";
    return mapped;
  }

  return mapped;
}

async function createChatCompletion(options = {}) {
  const provider = validateConfiguredAiProvider();
  if (provider === "azure_openai") {
    try {
      return await createAzureOpenAiChatCompletion(options);
    } catch (error) {
      logAiProviderError(provider, error);
      throw mapAzureOpenAiError(error);
    }
  }

  return createDeepSeekChatCompletion(options);
}

module.exports = {
  DEFAULT_AI_PROVIDER,
  createChatCompletion,
  getConfiguredAiProvider,
  validateConfiguredAiProvider,
};
