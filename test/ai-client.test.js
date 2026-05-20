const test = require("node:test");
const assert = require("node:assert/strict");

const AI_ENV_KEYS = [
  "AI_PROVIDER",
  "DEEPSEEK_API_KEY",
  "DEEPSEEK_MODEL",
  "DEEPSEEK_BASE_URL",
  "AZURE_OPENAI_ENDPOINT",
  "AZURE_OPENAI_API_KEY",
  "AZURE_OPENAI_API_VERSION",
  "AZURE_OPENAI_DEPLOYMENT_NAME",
  "AZURE_OPENAI_DEPLOYMENT_PRIMARY",
  "AZURE_OPENAI_DEPLOYMENT_FAST",
];

const ORIGINAL_ENV = Object.fromEntries(AI_ENV_KEYS.map((key) => [key, process.env[key]]));
const ORIGINAL_FETCH = global.fetch;
const ORIGINAL_CONSOLE_ERROR = console.error;

function restoreEnv() {
  for (const key of AI_ENV_KEYS) {
    if (ORIGINAL_ENV[key] === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = ORIGINAL_ENV[key];
    }
  }
}

function loadAiClient() {
  const aiClientPath = require.resolve("../api/_lib/ai-client");
  const azureClientPath = require.resolve("../api/_lib/azure-openai-client");
  const deepseekClientPath = require.resolve("../api/_lib/deepseek-client");
  delete require.cache[aiClientPath];
  delete require.cache[azureClientPath];
  delete require.cache[deepseekClientPath];
  return require("../api/_lib/ai-client");
}

test.afterEach(() => {
  restoreEnv();
  global.fetch = ORIGINAL_FETCH;
  console.error = ORIGINAL_CONSOLE_ERROR;
});

test("defaults to DeepSeek when AI_PROVIDER is unset", async () => {
  delete process.env.AI_PROVIDER;
  process.env.DEEPSEEK_API_KEY = "deepseek-test-key";

  let requestBody = null;
  global.fetch = async (_url, init) => {
    requestBody = JSON.parse(String(init.body || "{}"));
    return new Response(
      JSON.stringify({
        choices: [{ message: { content: "hello from deepseek" } }],
        usage: { prompt_tokens: 1, completion_tokens: 1 },
        model: "deepseek-v4-flash",
      }),
      {
        status: 200,
        headers: {
          "Content-Type": "application/json",
        },
      }
    );
  };

  const { createChatCompletion } = loadAiClient();
  const result = await createChatCompletion({
    messages: [{ role: "user", content: "Hi" }],
    thinking: "disabled",
  });

  assert.equal(result.content, "hello from deepseek");
  assert.equal(requestBody.model, "deepseek-v4-flash");
  assert.equal(requestBody.thinking.type, "disabled");
});

test("validates Azure configuration only when Azure is selected", async () => {
  process.env.AI_PROVIDER = "deepseek";
  const deepseekClient = loadAiClient();
  assert.equal(deepseekClient.validateConfiguredAiProvider(), "deepseek");

  process.env.AI_PROVIDER = "azure_openai";
  delete process.env.AZURE_OPENAI_ENDPOINT;
  delete process.env.AZURE_OPENAI_API_KEY;
  delete process.env.AZURE_OPENAI_API_VERSION;
  delete process.env.AZURE_OPENAI_DEPLOYMENT_NAME;
  delete process.env.AZURE_OPENAI_DEPLOYMENT_PRIMARY;
  delete process.env.AZURE_OPENAI_DEPLOYMENT_FAST;

  const azureClient = loadAiClient();
  assert.throws(
    () => azureClient.validateConfiguredAiProvider(),
    /Missing Azure OpenAI configuration|Missing Azure OpenAI deployment configuration/
  );
});

test("routes Azure requests through the configured deployment with a mocked OpenAI client", async () => {
  process.env.AI_PROVIDER = "azure_openai";
  process.env.AZURE_OPENAI_ENDPOINT = "https://example-resource.openai.azure.com";
  process.env.AZURE_OPENAI_API_KEY = "azure-test-key";
  process.env.AZURE_OPENAI_API_VERSION = "2024-10-21";
  process.env.AZURE_OPENAI_DEPLOYMENT_PRIMARY = "ai-primary-prod";
  process.env.AZURE_OPENAI_DEPLOYMENT_FAST = "ai-fast-prod";

  let capturedBody = null;
  const stubClient = {
    chat: {
      completions: {
        create(body) {
          capturedBody = body;
          return {
            withResponse: async () => ({
              data: {
                choices: [{ message: { content: "hello from azure" } }],
                usage: { prompt_tokens: 2, completion_tokens: 3 },
                model: "ai-primary-prod",
              },
              request_id: "req_test_123",
            }),
          };
        },
      },
    },
  };

  const { createChatCompletion } = loadAiClient();
  const result = await createChatCompletion({
    model: "deepseek-v4-pro",
    thinking: "enabled",
    reasoningEffort: "max",
    responseFormat: { type: "json_object" },
    messages: [{ role: "user", content: "Hi" }],
    clientFactory: () => stubClient,
  });

  assert.equal(capturedBody.model, "ai-primary-prod");
  assert.equal(capturedBody.reasoning_effort, "high");
  assert.deepEqual(capturedBody.response_format, { type: "json_object" });
  assert.equal(result.content, "hello from azure");
  assert.equal(result.model, "ai-primary-prod");
  assert.deepEqual(result.aiRoute, {
    provider: "azure_openai",
    requestedModel: "deepseek-v4-pro",
    deployment: "ai-primary-prod",
    apiVersion: "2024-10-21",
    endpointHost: "example-resource.openai.azure.com",
  });
});

test("routes Azure flash requests through the fast deployment env", async () => {
  process.env.AI_PROVIDER = "azure_openai";
  process.env.AZURE_OPENAI_ENDPOINT = "https://example-resource.openai.azure.com";
  process.env.AZURE_OPENAI_API_KEY = "azure-test-key";
  process.env.AZURE_OPENAI_API_VERSION = "2024-10-21";
  process.env.AZURE_OPENAI_DEPLOYMENT_PRIMARY = "ai-primary-prod";
  process.env.AZURE_OPENAI_DEPLOYMENT_FAST = "ai-fast-prod";

  let capturedBody = null;
  const stubClient = {
    chat: {
      completions: {
        create(body) {
          capturedBody = body;
          return {
            withResponse: async () => ({
              data: {
                choices: [{ message: { content: "fast azure" } }],
                usage: { prompt_tokens: 1, completion_tokens: 1 },
                model: "ai-fast-prod",
              },
              request_id: "req_test_fast",
            }),
          };
        },
      },
    },
  };

  const { createChatCompletion } = loadAiClient();
  const result = await createChatCompletion({
    model: "deepseek-v4-flash",
    thinking: "disabled",
    messages: [{ role: "user", content: "Hi" }],
    clientFactory: () => stubClient,
  });

  assert.equal(capturedBody.model, "ai-fast-prod");
  assert.equal(result.model, "ai-fast-prod");
  assert.equal(result.content, "fast azure");
  assert.deepEqual(result.aiRoute, {
    provider: "azure_openai",
    requestedModel: "deepseek-v4-flash",
    deployment: "ai-fast-prod",
    apiVersion: "2024-10-21",
    endpointHost: "example-resource.openai.azure.com",
  });
});

test("allows request provider to override AI_PROVIDER", async () => {
  process.env.AI_PROVIDER = "deepseek";
  process.env.AZURE_OPENAI_ENDPOINT = "https://example-resource.openai.azure.com";
  process.env.AZURE_OPENAI_API_KEY = "azure-test-key";
  process.env.AZURE_OPENAI_API_VERSION = "2024-10-21";
  process.env.AZURE_OPENAI_DEPLOYMENT_PRIMARY = "ai-primary-prod";
  process.env.AZURE_OPENAI_DEPLOYMENT_FAST = "ai-fast-prod";

  let capturedBody = null;
  const stubClient = {
    chat: {
      completions: {
        create(body) {
          capturedBody = body;
          return {
            withResponse: async () => ({
              data: {
                choices: [{ message: { content: "selected azure" } }],
                usage: { prompt_tokens: 1, completion_tokens: 1 },
                model: "ai-primary-prod",
              },
              request_id: "req_provider_override",
            }),
          };
        },
      },
    },
  };

  const { createChatCompletion } = loadAiClient();
  const result = await createChatCompletion({
    provider: "azure_openai",
    model: "primary",
    thinking: "disabled",
    messages: [{ role: "user", content: "Hi" }],
    clientFactory: () => stubClient,
  });

  assert.equal(capturedBody.model, "ai-primary-prod");
  assert.equal(result.content, "selected azure");
});

test("maps Azure SDK errors to the existing internal error shape and logs request metadata only", async () => {
  process.env.AI_PROVIDER = "azure_openai";
  process.env.AZURE_OPENAI_ENDPOINT = "https://example-resource.openai.azure.com";
  process.env.AZURE_OPENAI_API_KEY = "azure-test-key";
  process.env.AZURE_OPENAI_API_VERSION = "2024-10-21";
  process.env.AZURE_OPENAI_DEPLOYMENT_PRIMARY = "ai-primary-prod";
  process.env.AZURE_OPENAI_DEPLOYMENT_FAST = "ai-fast-prod";

  const logged = [];
  console.error = (...args) => {
    logged.push(args);
  };

  const stubClient = {
    chat: {
      completions: {
        create() {
          return {
            withResponse: async () => {
              const error = new Error("Raw Azure provider detail that should stay server-side");
              error.status = 401;
              error.requestID = "req_auth_401";
              error.code = "unauthorized";
              throw error;
            },
          };
        },
      },
    },
  };

  const { createChatCompletion } = loadAiClient();

  await assert.rejects(
    () =>
      createChatCompletion({
        messages: [{ role: "user", content: "Hi" }],
        clientFactory: () => stubClient,
      }),
    (error) => {
      assert.equal(error.message, "Azure OpenAI authentication failed.");
      assert.equal(error.status, 401);
      return true;
    }
  );

  assert.equal(logged.length, 1);
  assert.equal(logged[0][0], "[ai-provider] request failed");
  assert.deepEqual(logged[0][1], {
    provider: "azure_openai",
    requestId: "req_auth_401",
    status: 401,
    code: "unauthorized",
    requestedModel: null,
    deployment: "ai-primary-prod",
    apiVersion: "2024-10-21",
    endpointHost: "example-resource.openai.azure.com",
  });
});

test("maps Azure DeploymentNotFound and logs safe deployment routing metadata", async () => {
  process.env.AI_PROVIDER = "azure_openai";
  process.env.AZURE_OPENAI_ENDPOINT = "https://example-resource.openai.azure.com";
  process.env.AZURE_OPENAI_API_KEY = "azure-test-key";
  process.env.AZURE_OPENAI_API_VERSION = "2024-10-21";
  process.env.AZURE_OPENAI_DEPLOYMENT_PRIMARY = "ai-primary-prod";
  process.env.AZURE_OPENAI_DEPLOYMENT_FAST = "ai-fast-prod";

  const logged = [];
  console.error = (...args) => {
    logged.push(args);
  };

  const stubClient = {
    chat: {
      completions: {
        create(body) {
          assert.equal(body.model, "ai-fast-prod");
          assert.equal(body.messages[0].content, "Sensitive prompt text");
          return {
            withResponse: async () => {
              const error = new Error("Deployment missing");
              error.status = 404;
              error.code = "DeploymentNotFound";
              throw error;
            },
          };
        },
      },
    },
  };

  const { createChatCompletion } = loadAiClient();

  await assert.rejects(
    () =>
      createChatCompletion({
        provider: "azure_openai",
        model: "fast",
        messages: [{ role: "user", content: "Sensitive prompt text" }],
        clientFactory: () => stubClient,
      }),
    (error) => {
      assert.equal(error.message, "Azure OpenAI deployment not found.");
      assert.equal(error.status, 404);
      return true;
    }
  );

  assert.equal(logged.length, 1);
  assert.equal(logged[0][0], "[ai-provider] request failed");
  assert.deepEqual(logged[0][1], {
    provider: "azure_openai",
    requestId: null,
    status: 404,
    code: "DeploymentNotFound",
    requestedModel: "fast",
    deployment: "ai-fast-prod",
    apiVersion: "2024-10-21",
    endpointHost: "example-resource.openai.azure.com",
  });
  assert.equal(JSON.stringify(logged).includes("Sensitive prompt text"), false);
});
