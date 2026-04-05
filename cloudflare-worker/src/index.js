const JSON_BASE_HEADERS = {
  "content-type": "application/json; charset=utf-8",
  "cache-control": "no-store",
};

export default {
  async fetch(request, env) {
    const cors = buildCorsHeaders();

    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: {
          ...cors,
          "access-control-allow-methods": "GET, PUT, OPTIONS",
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
      allow: "GET, PUT, OPTIONS",
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
