const { requireApiAuth } = require("../_lib/require-api-auth");
const { createChatCompletion, normalizeText } = require("../_lib/deepseek-client");

function toArray(value) {
  return Array.isArray(value) ? value : [];
}

module.exports = async (req, res) => {
  if (req.method !== "POST") {
    res.status(405).json({
      error: {
        code: "METHOD_NOT_ALLOWED",
        message: "Method Not Allowed",
      },
    });
    return;
  }

  if (!(await requireApiAuth(req, res))) {
    return;
  }

  const body = req.body && typeof req.body === "object" ? req.body : {};

  try {
    const completion = await createChatCompletion({
      model: body.model,
      thinking: body.thinking,
      reasoningEffort: body.reasoningEffort || body.reasoning_effort,
      maxTokens: body.maxTokens ?? body.max_tokens,
      temperature: body.temperature,
      topP: body.topP ?? body.top_p,
      presencePenalty: body.presencePenalty ?? body.presence_penalty,
      frequencyPenalty: body.frequencyPenalty ?? body.frequency_penalty,
      responseFormat: body.responseFormat || body.response_format,
      messages: toArray(body.messages),
    });

    res.status(200).json({
      success: true,
      model: completion.model,
      thinking: completion.thinking,
      reasoningEffort: completion.reasoningEffort,
      content: normalizeText(completion.content),
      reasoningContent: normalizeText(completion.reasoningContent),
      usage: completion.usage,
      raw: body.includeRaw === true ? completion.payload : undefined,
    });
  } catch (error) {
    res.status(Number(error?.status) || 500).json({
      error: {
        code: "AI_CHAT_FAILED",
        message: error?.message || "Could not complete the DeepSeek chat request.",
      },
    });
  }
};
