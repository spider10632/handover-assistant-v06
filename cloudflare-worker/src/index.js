const JSON_BASE_HEADERS = {
  "content-type": "application/json; charset=utf-8",
  "cache-control": "no-store",
};
const MAX_TRANSLATE_ITEMS = 60;
const MAX_TRANSLATE_TEXT_LENGTH = 2000;

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
  if (!env || !env.GEMINI_API_KEY) {
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

  let translated;
  try {
    translated = await translateTextsWithGemini(env, targetLang, normalizedTexts);
  } catch (error) {
    return jsonResponse(
      {
        ok: false,
        error: "TRANSLATION_FAILED",
        message: String(error && error.message ? error.message : "translate failed"),
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
      translations: translated,
    },
    200,
    cors,
  );
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
  const model = String(env.GEMINI_MODEL || "gemini-2.5-flash-lite").trim() || "gemini-2.5-flash-lite";
  const apiKey = String(env.GEMINI_API_KEY || "").trim();
  if (!apiKey) {
    throw new Error("missing api key");
  }
  const endpoint =
    "https://generativelanguage.googleapis.com/v1beta/models/" +
    encodeURIComponent(model) +
    ":generateContent?key=" +
    encodeURIComponent(apiKey);
  const targetLabel = targetLang === "en" ? "English" : "Traditional Chinese (Taiwan)";

  const requestPayload = {
    systemInstruction: {
      parts: [
        {
          text:
            "You are a professional concierge handover translator. Keep original facts exactly, including names, room numbers, dates, and amounts. Return JSON only.",
        },
      ],
    },
    contents: [
      {
        role: "user",
        parts: [
          {
            text: JSON.stringify({
              task: "Translate each text to the target language.",
              target_language: targetLabel,
              output_format: {
                translations: ["same length as input texts array"],
              },
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

  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "content-type": "application/json",
    },
    body: JSON.stringify(requestPayload),
  });

  if (!response.ok) {
    const detail = await response.text();
    throw new Error("gemini http " + response.status + " " + detail);
  }

  const data = await response.json();
  const text = extractGeminiText(data);
  if (!text) {
    throw new Error("empty gemini response");
  }

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (error) {
    throw new Error("invalid gemini json");
  }
  const translations = Array.isArray(parsed && parsed.translations) ? parsed.translations : [];
  if (translations.length !== texts.length) {
    throw new Error("translation length mismatch");
  }
  return translations.map((item, index) => {
    const value = String(item == null ? "" : item).trim();
    return value || String(texts[index] || "");
  });
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
