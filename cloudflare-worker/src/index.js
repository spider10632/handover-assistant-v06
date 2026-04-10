const JSON_BASE_HEADERS = {
  "content-type": "application/json; charset=utf-8",
  "cache-control": "no-store",
};
const MAX_TRANSLATE_ITEMS = 60;
const MAX_TRANSLATE_TEXT_LENGTH = 2000;
const TRANSLATE_RETRY_STATUSES = new Set([429, 500, 502, 503, 504]);
const TRANSLATE_MODEL_FALLBACKS = ["gemini-2.5-flash-lite", "gemini-2.5-flash"];
const TRANSLATE_PROVIDER_DEFAULT_ORDER = ["azure", "deepl", "google-cloud", "google-public", "gemini"];

export default {
  async fetch(request, env) {
    const cors = buildCorsHeaders();

    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: {
          ...cors,
          "access-control-allow-methods": "GET, PUT, POST, OPTIONS",
          "access-control-allow-headers": "content-type",
          "access-control-max-age": "86400",
        },
      });
    }

    const url = new URL(request.url);
    const path = url.pathname;

    if (path === "/health" && request.method === "GET") {
      return jsonResponse({ ok: true, service: "handover-cloud-api" }, 200, cors);
    }

    const translateMatch = path.match(/^\/v1\/translate\/([a-z0-9_-]+)$/i);
    if (translateMatch) {
      const serverId = normalizeServerId(translateMatch[1]);
      if (!serverId) {
        return jsonResponse({ ok: false, error: "INVALID_SERVER_ID" }, 400, cors);
      }
      const allowedUsers = parseAllowedUsers(env.ALLOWED_USERS);
      if (allowedUsers.size > 0 && !allowedUsers.has(serverId)) {
        return jsonResponse({ ok: false, error: "FORBIDDEN_SERVER" }, 403, cors);
      }
      if (request.method !== "POST") {
        return jsonResponse({ ok: false, error: "METHOD_NOT_ALLOWED" }, 405, {
          ...cors,
          allow: "POST, OPTIONS",
        });
      }
      return handleTranslate(request, env, serverId, cors);
    }

    const match = path.match(/^\/v1\/state\/([a-z0-9_-]+)$/i);
    if (!match) {
      return jsonResponse({ ok: false, error: "NOT_FOUND" }, 404, cors);
    }

    if (!env || !env.DB) {
      return jsonResponse({ ok: false, error: "D1_NOT_CONFIGURED" }, 500, cors);
    }

    const serverId = normalizeServerId(match[1]);
    if (!serverId) {
      return jsonResponse({ ok: false, error: "INVALID_SERVER_ID" }, 400, cors);
    }

    const allowedUsers = parseAllowedUsers(env.ALLOWED_USERS);
    if (allowedUsers.size > 0 && !allowedUsers.has(serverId)) {
      return jsonResponse({ ok: false, error: "FORBIDDEN_SERVER" }, 403, cors);
    }

    if (request.method === "GET") {
      return handleGetState(env, serverId, cors);
    }

    if (request.method === "PUT") {
      return handlePutState(request, env, serverId, cors);
    }

    return jsonResponse({ ok: false, error: "METHOD_NOT_ALLOWED" }, 405, {
      ...cors,
      allow: "GET, PUT, POST, OPTIONS",
    });
  },
};

async function handleGetState(env, serverId, cors) {
  const row = await env.DB.prepare("SELECT payload, updated_at FROM handover_state WHERE server_id = ?")
    .bind(serverId)
    .first();

  if (!row) {
    return jsonResponse({ ok: false, error: "NOT_FOUND" }, 404, cors);
  }

  const payload = safeJsonParse(row.payload);
  if (!payload) {
    return jsonResponse({ ok: false, error: "CORRUPTED_PAYLOAD" }, 500, cors);
  }

  return jsonResponse(
    {
      ok: true,
      payload,
      updatedAt: row.updated_at || null,
    },
    200,
    cors,
  );
}

async function handlePutState(request, env, serverId, cors) {
  let payload;
  try {
    payload = await request.json();
  } catch (error) {
    return jsonResponse({ ok: false, error: "INVALID_JSON" }, 400, cors);
  }

  if (!isValidBackupPayload(payload, serverId)) {
    return jsonResponse({ ok: false, error: "INVALID_PAYLOAD" }, 400, cors);
  }

  const updatedAt = new Date().toISOString();
  const bodyText = JSON.stringify(payload);

  await env.DB.prepare(
    "INSERT INTO handover_state (server_id, payload, updated_at) VALUES (?, ?, ?) ON CONFLICT(server_id) DO UPDATE SET payload = excluded.payload, updated_at = excluded.updated_at",
  )
    .bind(serverId, bodyText, updatedAt)
    .run();

  return jsonResponse({ ok: true, updatedAt }, 200, cors);
}

async function handleTranslate(request, env, serverId, cors) {
  if (!env) {
    return jsonResponse({ ok: false, error: "TRANSLATION_NOT_CONFIGURED" }, 503, cors);
  }

  let body;
  try {
    body = await request.json();
  } catch (error) {
    return jsonResponse({ ok: false, error: "INVALID_JSON" }, 400, cors);
  }

  const targetLang = normalizeTargetLang(body && body.targetLang);
  if (!targetLang) {
    return jsonResponse({ ok: false, error: "INVALID_TARGET_LANG" }, 400, cors);
  }

  const texts = Array.isArray(body && body.texts) ? body.texts : [];
  if (texts.length === 0 || texts.length > MAX_TRANSLATE_ITEMS) {
    return jsonResponse({ ok: false, error: "INVALID_TEXT_COUNT" }, 400, cors);
  }

  const normalizedTexts = [];
  for (const item of texts) {
    const text = String(item == null ? "" : item);
    if (text.length > MAX_TRANSLATE_TEXT_LENGTH) {
      return jsonResponse({ ok: false, error: "TEXT_TOO_LONG" }, 400, cors);
    }
    normalizedTexts.push(text);
  }

  const nonEmpty = normalizedTexts.filter((item) => String(item || "").trim().length > 0);
  if (nonEmpty.length === 0) {
    return jsonResponse(
      {
        ok: true,
        serverId,
        targetLang,
        translations: normalizedTexts,
      },
      200,
      cors,
    );
  }

  const order = buildTranslateProviderOrder(env);
  const failures = [];
  let translated = null;
  let provider = "";

  for (const providerName of order) {
    if (!isTranslateProviderReady(providerName, env)) {
      failures.push(providerName + ":NOT_CONFIGURED");
      continue;
    }
    try {
      translated = await translateTextsByProvider(providerName, env, targetLang, normalizedTexts);
      provider = providerName;
      break;
    } catch (error) {
      const status = Number(error && error.status ? error.status : 0);
      const message = String(error && error.message ? error.message : "translate failed");
      failures.push(providerName + ":" + (status || "ERR") + ":" + message.slice(0, 220));
    }
  }

  if (!translated || !provider) {
    return jsonResponse(
      {
        ok: false,
        error: "TRANSLATION_FAILED",
        message: "all providers failed",
        details: failures,
      },
      502,
      cors,
    );
  }

  return jsonResponse(
    {
      ok: true,
      serverId,
      targetLang,
      provider,
      translations: translated,
    },
    200,
    cors,
  );
}

function buildTranslateProviderOrder(env) {
  const configured = String(env.TRANSLATE_PROVIDER_ORDER || "")
    .split(",")
    .map((item) => String(item || "").trim().toLowerCase())
    .filter(Boolean);
  const source = configured.length > 0 ? configured : TRANSLATE_PROVIDER_DEFAULT_ORDER;
  const valid = source.filter((item) => {
    return item === "azure" || item === "deepl" || item === "google-cloud" || item === "gemini" || item === "google-public";
  });
  return Array.from(new Set(valid.length > 0 ? valid : TRANSLATE_PROVIDER_DEFAULT_ORDER));
}

function isTranslateProviderReady(provider, env) {
  if (provider === "azure") {
    return Boolean(String(env.AZURE_TRANSLATOR_KEY || "").trim());
  }
  if (provider === "deepl") {
    return Boolean(String(env.DEEPL_API_KEY || "").trim());
  }
  if (provider === "google-cloud") {
    return Boolean(String(env.GOOGLE_TRANSLATE_API_KEY || "").trim());
  }
  if (provider === "gemini") {
    return Boolean(String(env.GEMINI_API_KEY || "").trim());
  }
  if (provider === "google-public") {
    return true;
  }
  return false;
}

async function translateTextsByProvider(provider, env, targetLang, texts) {
  if (provider === "azure") {
    return translateTextsWithAzure(env, targetLang, texts);
  }
  if (provider === "deepl") {
    return translateTextsWithDeepL(env, targetLang, texts);
  }
  if (provider === "google-cloud") {
    return translateTextsWithGoogleCloud(env, targetLang, texts);
  }
  if (provider === "gemini") {
    return translateTextsWithGemini(env, targetLang, texts);
  }
  if (provider === "google-public") {
    return translateTextsWithGooglePublic(targetLang, texts);
  }
  throw new Error("unknown provider: " + provider);
}

function parseAllowedUsers(raw) {
  const value = String(raw || "").trim();
  if (!value) {
    return new Set();
  }
  return new Set(
    value
      .split(",")
      .map(function (item) {
        return normalizeServerId(item);
      })
      .filter(Boolean),
  );
}

function normalizeServerId(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]/g, "");
}

function normalizeTargetLang(value) {
  const text = String(value || "")
    .trim()
    .toLowerCase();
  if (text === "en" || text === "zh") {
    return text;
  }
  return "";
}

async function translateTextsWithGemini(env, targetLang, texts) {
  const apiKey = String(env.GEMINI_API_KEY || "").trim();
  if (!apiKey) {
    throw new Error("missing api key");
  }
  const targetLabel = targetLang === "en" ? "English" : "Traditional Chinese (Taiwan)";
  const models = buildTranslateModelList(env);

  const requestPayload = {
    systemInstruction: {
      parts: [
        {
          text:
            "You are a professional concierge handover translator. Translate fully and faithfully. Never summarize, shorten, omit, reorder, or rewrite facts. Keep all names, room numbers, dates, times, amounts, symbols, separators, and line breaks. Return JSON only.",
        },
      ],
    },
    contents: [
      {
        role: "user",
        parts: [
          {
            text: JSON.stringify({
              task: "Translate each text to the target language with full fidelity.",
              target_language: targetLabel,
              output_format: {
                translations: ["same length as input texts array"],
              },
              strict_rules: [
                "Output array length must equal input length.",
                "Each output item must correspond to the same input index.",
                "Do not summarize or shorten.",
                "Preserve all factual details and order.",
                "If unsure about a token, keep it as-is.",
              ],
              texts: texts,
            }),
          },
        ],
      },
    ],
    generationConfig: {
      temperature: 0,
      responseMimeType: "application/json",
    },
  };

  let lastError = null;
  for (const model of models) {
    for (let attempt = 1; attempt <= 3; attempt += 1) {
      try {
        const translations = await requestGeminiTranslations(apiKey, model, requestPayload, texts.length);
        return translations.map((item, index) => {
          const value = String(item == null ? "" : item).trim();
          return value || String(texts[index] || "");
        });
      } catch (error) {
        lastError = error;
        const status = Number(error && error.status ? error.status : 0);
        const message = String(error && error.message ? error.message : "");
        if (status === 400 && /API_KEY_INVALID|INVALID_ARGUMENT/i.test(message)) {
          throw error;
        }
        if (attempt >= 3 || !TRANSLATE_RETRY_STATUSES.has(status)) {
          break;
        }
        const backoffMs = 250 * attempt * attempt;
        await sleep(backoffMs);
      }
    }
  }
  throw lastError || new Error("translation failed");
}

function buildTranslateModelList(env) {
  const preferred = String(env.GEMINI_MODEL || "gemini-2.5-flash-lite")
    .trim()
    .toLowerCase();
  const list = [preferred, ...TRANSLATE_MODEL_FALLBACKS]
    .map((item) => String(item || "").trim().toLowerCase())
    .filter(Boolean);
  return Array.from(new Set(list));
}

async function requestGeminiTranslations(apiKey, model, requestPayload, expectedLength) {
  const endpoint =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(model) +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "content-type": "application/json",
    },
    body: JSON.stringify(requestPayload),
  });

  if (!response.ok) {
    const detail = await response.text();
    const error = new Error("gemini http " + response.status + " " + detail);
    error.status = response.status;
    throw error;
  }

  const data = await response.json();
  const text = extractGeminiText(data);
  if (!text) {
    throw new Error("empty gemini response");
  }

  const parsed = parseJsonLoose(text);
  const translations = Array.isArray(parsed && parsed.translations) ? parsed.translations : [];
  if (translations.length !== expectedLength) {
    throw new Error("translation length mismatch");
  }
  return translations;
}

function parseJsonLoose(raw) {
  const text = String(raw || "").trim();
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text);
  } catch (error) {
    // continue
  }

  const noFence = text.replace(/^```json\s*/i, "").replace(/^```\s*/i, "").replace(/\s*```$/, "").trim();
  if (noFence && noFence !== text) {
    try {
      return JSON.parse(noFence);
    } catch (error) {
      // continue
    }
  }

  const firstBrace = text.indexOf("{");
  const lastBrace = text.lastIndexOf("}");
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    const block = text.slice(firstBrace, lastBrace + 1);
    try {
      return JSON.parse(block);
    } catch (error) {
      return null;
    }
  }
  return null;
}

async function runTranslateWithRetries(jobName, handler) {
  let lastError = null;
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    try {
      return await handler(attempt);
    } catch (error) {
      lastError = error;
      const status = Number(error && error.status ? error.status : 0);
      if (attempt >= 3 || !TRANSLATE_RETRY_STATUSES.has(status)) {
        break;
      }
      await sleep(250 * attempt * attempt);
    }
  }
  const detail = String(lastError && lastError.message ? lastError.message : "translate failed");
  throw new Error(jobName + " failed: " + detail);
}

function normalizeProviderTranslatedList(translations, sourceTexts) {
  const source = Array.isArray(sourceTexts) ? sourceTexts : [];
  const list = Array.isArray(translations) ? translations : [];
  if (list.length !== source.length) {
    throw new Error("translation length mismatch");
  }
  return list.map((item, index) => {
    const value = String(item == null ? "" : item).trim();
    return value || String(source[index] || "");
  });
}

async function translateTextsWithAzure(env, targetLang, texts) {
  const key = String(env.AZURE_TRANSLATOR_KEY || "").trim();
  if (!key) {
    throw new Error("azure key missing");
  }
  const region = String(env.AZURE_TRANSLATOR_REGION || "").trim();
  const endpointBase = String(env.AZURE_TRANSLATOR_ENDPOINT || "https://api.cognitive.microsofttranslator.com").trim();
  const endpoint =
    endpointBase.replace(/\/+$/, "") +
    "/translate?api-version=3.0&to=" +
    encodeURIComponent(targetLang === "zh" ? "zh-Hant" : "en");

  return runTranslateWithRetries("azure", async () => {
    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "ocp-apim-subscription-key": key,
        ...(region ? { "ocp-apim-subscription-region": region } : {}),
      },
      body: JSON.stringify(
        texts.map((text) => {
          return { Text: String(text || "") };
        }),
      ),
    });
    if (!response.ok) {
      const detail = await response.text();
      const error = new Error("azure http " + response.status + " " + detail);
      error.status = response.status;
      throw error;
    }
    const payload = await response.json();
    if (!Array.isArray(payload)) {
      throw new Error("azure invalid response");
    }
    const translated = payload.map((item) => {
      const list = item && Array.isArray(item.translations) ? item.translations : [];
      const first = list[0];
      return first && typeof first.text === "string" ? first.text : "";
    });
    return normalizeProviderTranslatedList(translated, texts);
  });
}

async function translateTextsWithDeepL(env, targetLang, texts) {
  const key = String(env.DEEPL_API_KEY || "").trim();
  if (!key) {
    throw new Error("deepl key missing");
  }
  const endpoint = String(env.DEEPL_API_URL || "https://api-free.deepl.com/v2/translate").trim();
  const target = targetLang === "zh" ? "ZH" : "EN";

  return runTranslateWithRetries("deepl", async () => {
    const body = new URLSearchParams();
    body.set("target_lang", target);
    for (const text of texts) {
      body.append("text", String(text || ""));
    }

    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        Authorization: "DeepL-Auth-Key " + key,
        "content-type": "application/x-www-form-urlencoded",
      },
      body: body.toString(),
    });
    if (!response.ok) {
      const detail = await response.text();
      const error = new Error("deepl http " + response.status + " " + detail);
      error.status = response.status;
      throw error;
    }
    const payload = await response.json();
    const list = payload && Array.isArray(payload.translations) ? payload.translations : [];
    const translated = list.map((item) => {
      return item && typeof item.text === "string" ? item.text : "";
    });
    return normalizeProviderTranslatedList(translated, texts);
  });
}

async function translateTextsWithGoogleCloud(env, targetLang, texts) {
  const key = String(env.GOOGLE_TRANSLATE_API_KEY || "").trim();
  if (!key) {
    throw new Error("google cloud key missing");
  }
  const endpoint = "https://translation.googleapis.com/language/translate/v2?key=" + encodeURIComponent(key);
  const target = normalizeUiTargetLang(targetLang);

  return runTranslateWithRetries("google-cloud", async () => {
    const response = await fetch(endpoint, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: JSON.stringify({
        q: texts,
        target,
        format: "text",
      }),
    });
    if (!response.ok) {
      const detail = await response.text();
      const error = new Error("google cloud http " + response.status + " " + detail);
      error.status = response.status;
      throw error;
    }
    const payload = await response.json();
    const list =
      payload &&
      payload.data &&
      Array.isArray(payload.data.translations)
        ? payload.data.translations
        : [];
    const translated = list.map((item) => {
      return decodeHtmlEntity(String(item && item.translatedText ? item.translatedText : ""));
    });
    return normalizeProviderTranslatedList(translated, texts);
  });
}

async function translateTextsWithGooglePublic(targetLang, texts) {
  const tl = normalizeUiTargetLang(targetLang);
  const results = [];
  for (const source of texts) {
    const text = String(source || "");
    if (!text.trim()) {
      results.push(text);
      continue;
    }
    const url =
      "https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=" +
      encodeURIComponent(tl) +
      "&dt=t&q=" +
      encodeURIComponent(text);
    const response = await fetch(url, {
      method: "GET",
      headers: {
        accept: "application/json",
      },
    });
    if (!response.ok) {
      const detail = await response.text();
      throw new Error("google public translate http " + response.status + " " + detail);
    }
    const payload = await response.json();
    const translated = extractGooglePublicTranslatedText(payload);
    results.push(translated || text);
  }
  return results;
}

function normalizeUiTargetLang(targetLang) {
  const normalized = normalizeTargetLang(targetLang);
  if (normalized === "zh") {
    return "zh-TW";
  }
  return "en";
}

function decodeHtmlEntity(input) {
  const text = String(input || "");
  if (!text.includes("&")) {
    return text;
  }
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x2F;/g, "/");
}

function extractGooglePublicTranslatedText(payload) {
  const root = Array.isArray(payload) ? payload : null;
  if (!root || !Array.isArray(root[0])) {
    return "";
  }
  const sentenceParts = root[0];
  return sentenceParts
    .map((entry) => {
      return Array.isArray(entry) ? String(entry[0] || "") : "";
    })
    .join("")
    .trim();
}

function extractGeminiText(data) {
  const candidates = data && Array.isArray(data.candidates) ? data.candidates : [];
  const first = candidates[0];
  const parts = first && first.content && Array.isArray(first.content.parts) ? first.content.parts : [];
  return parts
    .map((part) => {
      return typeof part.text === "string" ? part.text : "";
    })
    .join("")
    .trim();
}

function sleep(ms) {
  const wait = Number(ms) > 0 ? Number(ms) : 0;
  return new Promise((resolve) => setTimeout(resolve, wait));
}

function buildCorsHeaders() {
  return {
    "access-control-allow-origin": "*",
    "access-control-expose-headers": "content-type",
  };
}

function jsonResponse(data, status, extraHeaders) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      ...JSON_BASE_HEADERS,
      ...(extraHeaders || {}),
    },
  });
}

function safeJsonParse(raw) {
  if (!raw || typeof raw !== "string") {
    return null;
  }
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function isValidBackupPayload(payload, serverId) {
  if (!payload || typeof payload !== "object" || Array.isArray(payload)) {
    return false;
  }
  if (normalizeServerId(payload.serverId) !== serverId) {
    return false;
  }
  if (payload.tasks !== undefined && !Array.isArray(payload.tasks)) {
    return false;
  }
  if (payload.todayOverview !== undefined && (typeof payload.todayOverview !== "object" || payload.todayOverview === null)) {
    return false;
  }
  if (payload.deletedTaskIds !== undefined && (typeof payload.deletedTaskIds !== "object" || payload.deletedTaskIds === null)) {
    return false;
  }
  return true;
}
