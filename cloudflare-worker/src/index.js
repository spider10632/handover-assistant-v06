const JSON_BASE_HEADERS = {
  "content-type": "application/json; charset=utf-8",
  "cache-control": "no-store",
};
const MAX_TRANSLATE_ITEMS = 60;
const MAX_TRANSLATE_TEXT_LENGTH = 2000;
const TRANSLATE_RETRY_STATUSES = new Set([429, 500, 502, 503, 504]);
const TRANSLATE_MODEL_FALLBACKS = ["gemini-2.5-flash-lite", "gemini-2.5-flash"];
const TRANSLATE_PROVIDER_DEFAULT_ORDER = ["azure", "deepl", "google-cloud", "google-public", "gemini"];
const SESSION_TTL_DAYS = 30;
const AUDIT_DEFAULT_LIMIT = 200;
const AUDIT_MAX_LIMIT = 500;
const ACCOUNT_ROLES = new Set(["admin", "manager", "staff"]);

export default {
  async fetch(request, env) {
    const cors = buildCorsHeaders();

    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: {
          ...cors,
          "access-control-allow-methods": "GET, PUT, POST, PATCH, DELETE, OPTIONS",
          "access-control-allow-headers": "content-type, authorization",
          "access-control-max-age": "86400",
        },
      });
    }

    const url = new URL(request.url);
    const path = url.pathname;
    const allowedUsers = parseAllowedUsers(env && env.ALLOWED_USERS);
    const hasDb = Boolean(env && env.DB);

    if (path === "/health" && request.method === "GET") {
      return jsonResponse({ ok: true, service: "handover-cloud-api" }, 200, cors);
    }

    if (path.startsWith("/v2/")) {
      if (!hasDb) {
        return jsonResponse({ ok: false, error: "D1_NOT_CONFIGURED" }, 500, cors);
      }
      return handleV2Request(request, env, path, allowedUsers, cors);
    }

    const translateMatch = path.match(/^\/v1\/translate\/([a-z0-9_-]+)$/i);
    if (translateMatch) {
      const serverId = normalizeServerId(translateMatch[1]);
      if (!serverId) {
        return jsonResponse({ ok: false, error: "INVALID_SERVER_ID" }, 400, cors);
      }
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

    if (!hasDb) {
      return jsonResponse({ ok: false, error: "D1_NOT_CONFIGURED" }, 500, cors);
    }

    const serverId = normalizeServerId(match[1]);
    if (!serverId) {
      return jsonResponse({ ok: false, error: "INVALID_SERVER_ID" }, 400, cors);
    }

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

async function handleV2Request(request, env, path, allowedUsers, cors) {
  if (path === "/v2/auth/login") {
    if (request.method !== "POST") {
      return jsonResponse({ ok: false, error: "METHOD_NOT_ALLOWED" }, 405, {
        ...cors,
        allow: "POST, OPTIONS",
      });
    }
    return handleV2Login(request, env, allowedUsers, cors);
  }

  if (path === "/v2/auth/me") {
    if (request.method !== "GET") {
      return jsonResponse({ ok: false, error: "METHOD_NOT_ALLOWED" }, 405, {
        ...cors,
        allow: "GET, OPTIONS",
      });
    }
    const auth = await requireSession(request, env, null, cors);
    if (!auth.ok) {
      return auth.response;
    }
    return jsonResponse(
      {
        ok: true,
        serverId: auth.session.serverId,
        username: auth.session.username,
        role: auth.session.role,
        expiresAt: auth.session.expiresAt,
      },
      200,
      cors,
    );
  }

  if (path === "/v2/auth/logout") {
    if (request.method !== "POST") {
      return jsonResponse({ ok: false, error: "METHOD_NOT_ALLOWED" }, 405, {
        ...cors,
        allow: "POST, OPTIONS",
      });
    }
    const auth = await requireSession(request, env, null, cors);
    if (!auth.ok) {
      return auth.response;
    }
    await env.DB.prepare("UPDATE auth_sessions SET revoked_at = ? WHERE token_hash = ?")
      .bind(new Date().toISOString(), auth.session.tokenHash)
      .run();
    return jsonResponse({ ok: true }, 200, cors);
  }

  const stateMatch = path.match(/^\/v2\/state\/([a-z0-9_-]+)$/i);
  if (stateMatch) {
    const serverId = normalizeServerId(stateMatch[1]);
    if (!serverId) {
      return jsonResponse({ ok: false, error: "INVALID_SERVER_ID" }, 400, cors);
    }
    if (allowedUsers.size > 0 && !allowedUsers.has(serverId)) {
      return jsonResponse({ ok: false, error: "FORBIDDEN_SERVER" }, 403, cors);
    }
    const auth = await requireSession(request, env, serverId, cors);
    if (!auth.ok) {
      return auth.response;
    }
    if (request.method === "GET") {
      return handleGetState(env, serverId, cors);
    }
    if (request.method === "PUT") {
      return handlePutState(request, env, serverId, cors, auth.session);
    }
    return jsonResponse({ ok: false, error: "METHOD_NOT_ALLOWED" }, 405, {
      ...cors,
      allow: "GET, PUT, OPTIONS",
    });
  }

  const accountsMatch = path.match(/^\/v2\/admin\/([a-z0-9_-]+)\/accounts(?:\/([a-z0-9_-]+))?$/i);
  if (accountsMatch) {
    const serverId = normalizeServerId(accountsMatch[1]);
    const targetUsername = normalizeAccountName(accountsMatch[2] || "");
    if (!serverId) {
      return jsonResponse({ ok: false, error: "INVALID_SERVER_ID" }, 400, cors);
    }
    if (allowedUsers.size > 0 && !allowedUsers.has(serverId)) {
      return jsonResponse({ ok: false, error: "FORBIDDEN_SERVER" }, 403, cors);
    }
    const auth = await requireAdminSession(request, env, serverId, cors);
    if (!auth.ok) {
      return auth.response;
    }
    if (request.method === "GET" && !targetUsername) {
      return handleAdminListAccounts(env, serverId, cors);
    }
    if (request.method === "POST" && !targetUsername) {
      return handleAdminCreateAccount(request, env, serverId, auth.session, cors);
    }
    if (request.method === "PATCH" && targetUsername) {
      return handleAdminUpdateAccount(request, env, serverId, targetUsername, auth.session, cors);
    }
    if (request.method === "DELETE" && targetUsername) {
      return handleAdminDeleteAccount(env, serverId, targetUsername, auth.session, cors);
    }
    return jsonResponse({ ok: false, error: "METHOD_NOT_ALLOWED" }, 405, {
      ...cors,
      allow: "GET, POST, PATCH, DELETE, OPTIONS",
    });
  }

  const auditMatch = path.match(/^\/v2\/admin\/([a-z0-9_-]+)\/audit$/i);
  if (auditMatch) {
    const serverId = normalizeServerId(auditMatch[1]);
    if (!serverId) {
      return jsonResponse({ ok: false, error: "INVALID_SERVER_ID" }, 400, cors);
    }
    if (allowedUsers.size > 0 && !allowedUsers.has(serverId)) {
      return jsonResponse({ ok: false, error: "FORBIDDEN_SERVER" }, 403, cors);
    }
    const auth = await requireAdminSession(request, env, serverId, cors);
    if (!auth.ok) {
      return auth.response;
    }
    if (request.method !== "GET") {
      return jsonResponse({ ok: false, error: "METHOD_NOT_ALLOWED" }, 405, {
        ...cors,
        allow: "GET, OPTIONS",
      });
    }
    const url = new URL(request.url);
    return handleAdminListAudit(env, serverId, url.searchParams, cors);
  }

  return jsonResponse({ ok: false, error: "NOT_FOUND" }, 404, cors);
}

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

async function handlePutState(request, env, serverId, cors, actorSession) {
  let payload;
  try {
    payload = await request.json();
  } catch (error) {
    return jsonResponse({ ok: false, error: "INVALID_JSON" }, 400, cors);
  }

  if (!isValidBackupPayload(payload, serverId)) {
    return jsonResponse({ ok: false, error: "INVALID_PAYLOAD" }, 400, cors);
  }

  let previousPayload = null;
  if (actorSession) {
    const previousRow = await env.DB.prepare("SELECT payload FROM handover_state WHERE server_id = ?")
      .bind(serverId)
      .first();
    previousPayload = safeJsonParse(previousRow && previousRow.payload ? previousRow.payload : "");
  }

  const updatedAt = new Date().toISOString();
  const bodyText = JSON.stringify(payload);

  await env.DB.prepare(
    "INSERT INTO handover_state (server_id, payload, updated_at) VALUES (?, ?, ?) ON CONFLICT(server_id) DO UPDATE SET payload = excluded.payload, updated_at = excluded.updated_at",
  )
    .bind(serverId, bodyText, updatedAt)
    .run();

  if (actorSession) {
    try {
      await writeAuditForStateChange(env, serverId, actorSession, previousPayload, payload, updatedAt);
    } catch (error) {
      // Avoid blocking state write on audit issue.
      console.error("writeAuditForStateChange failed", error);
    }
  }

  return jsonResponse({ ok: true, updatedAt }, 200, cors);
}

async function handleV2Login(request, env, allowedUsers, cors) {
  let body;
  try {
    body = await request.json();
  } catch (error) {
    return jsonResponse({ ok: false, error: "INVALID_JSON" }, 400, cors);
  }

  const serverId = normalizeServerId(body && body.serverId);
  const password = String(body && body.password ? body.password : "");
  if (!serverId) {
    return jsonResponse({ ok: false, error: "INVALID_SERVER_ID" }, 400, cors);
  }
  if (allowedUsers.size > 0 && !allowedUsers.has(serverId)) {
    return jsonResponse({ ok: false, error: "FORBIDDEN_SERVER" }, 403, cors);
  }
  if (password.length < 4) {
    return jsonResponse({ ok: false, error: "INVALID_PASSWORD" }, 400, cors);
  }

  await ensureBootstrapAccounts(env, serverId);

  const passwordHash = await hashLoginPassword(serverId, password);
  const row = await env.DB.prepare(
    "SELECT username, role, enabled FROM auth_accounts WHERE server_id = ? AND password_hash = ? LIMIT 1",
  )
    .bind(serverId, passwordHash)
    .first();

  if (!row) {
    return jsonResponse({ ok: false, error: "INVALID_CREDENTIALS" }, 401, cors);
  }
  if (!toDbBoolean(row.enabled)) {
    return jsonResponse({ ok: false, error: "ACCOUNT_DISABLED" }, 403, cors);
  }

  const username = normalizeAccountName(row.username);
  const role = normalizeAccountRole(row.role);
  if (!username || !role) {
    return jsonResponse({ ok: false, error: "INVALID_ACCOUNT" }, 500, cors);
  }

  const token = createSessionToken();
  const tokenHash = await sha256Hex(token);
  const now = new Date();
  const createdAt = now.toISOString();
  const expiresAt = new Date(now.getTime() + SESSION_TTL_DAYS * 24 * 60 * 60 * 1000).toISOString();

  await env.DB.prepare(
    "INSERT INTO auth_sessions (token_hash, server_id, username, role, created_at, expires_at, revoked_at) VALUES (?, ?, ?, ?, ?, ?, NULL)",
  )
    .bind(tokenHash, serverId, username, role, createdAt, expiresAt)
    .run();

  await insertAuditLog(env, {
    serverId,
    actorUsername: username,
    actorRole: role,
    action: "auth_login",
    targetType: "session",
    targetId: username,
    summary: "login",
    details: JSON.stringify({ username, role }),
    createdAt,
  });

  return jsonResponse(
    {
      ok: true,
      token,
      serverId,
      username,
      role,
      expiresAt,
    },
    200,
    cors,
  );
}

async function requireSession(request, env, serverId, cors) {
  const token = getBearerToken(request);
  if (!token) {
    return {
      ok: false,
      response: jsonResponse({ ok: false, error: "MISSING_AUTH_TOKEN" }, 401, cors),
    };
  }

  const tokenHash = await sha256Hex(token);
  const row = await env.DB.prepare(
    "SELECT server_id, username, role, expires_at, revoked_at FROM auth_sessions WHERE token_hash = ? LIMIT 1",
  )
    .bind(tokenHash)
    .first();
  if (!row) {
    return {
      ok: false,
      response: jsonResponse({ ok: false, error: "INVALID_AUTH_TOKEN" }, 401, cors),
    };
  }
  if (row.revoked_at) {
    return {
      ok: false,
      response: jsonResponse({ ok: false, error: "SESSION_REVOKED" }, 401, cors),
    };
  }

  const nowMs = Date.now();
  const expiresMs = Date.parse(String(row.expires_at || ""));
  if (!Number.isFinite(expiresMs) || expiresMs <= nowMs) {
    return {
      ok: false,
      response: jsonResponse({ ok: false, error: "SESSION_EXPIRED" }, 401, cors),
    };
  }

  const resolvedServer = normalizeServerId(row.server_id);
  if (!resolvedServer) {
    return {
      ok: false,
      response: jsonResponse({ ok: false, error: "INVALID_SESSION" }, 401, cors),
    };
  }
  if (serverId && resolvedServer !== normalizeServerId(serverId)) {
    return {
      ok: false,
      response: jsonResponse({ ok: false, error: "SESSION_SERVER_MISMATCH" }, 403, cors),
    };
  }

  const username = normalizeAccountName(row.username);
  const role = normalizeAccountRole(row.role);
  if (!username || !role) {
    return {
      ok: false,
      response: jsonResponse({ ok: false, error: "INVALID_SESSION" }, 401, cors),
    };
  }

  return {
    ok: true,
    session: {
      tokenHash,
      token,
      serverId: resolvedServer,
      username,
      role,
      expiresAt: String(row.expires_at || ""),
    },
  };
}

async function requireAdminSession(request, env, serverId, cors) {
  const auth = await requireSession(request, env, serverId, cors);
  if (!auth.ok) {
    return auth;
  }
  if (auth.session.role !== "admin") {
    return {
      ok: false,
      response: jsonResponse({ ok: false, error: "FORBIDDEN_ROLE" }, 403, cors),
    };
  }
  return auth;
}

async function handleAdminListAccounts(env, serverId, cors) {
  await ensureBootstrapAccounts(env, serverId);
  const result = await env.DB.prepare(
    "SELECT username, role, enabled, created_at, updated_at FROM auth_accounts WHERE server_id = ? ORDER BY username ASC",
  )
    .bind(serverId)
    .all();
  const rows = Array.isArray(result && result.results) ? result.results : [];
  return jsonResponse(
    {
      ok: true,
      serverId,
      accounts: rows.map((row) => ({
        username: normalizeAccountName(row.username),
        role: normalizeAccountRole(row.role),
        enabled: toDbBoolean(row.enabled),
        createdAt: String(row.created_at || ""),
        updatedAt: String(row.updated_at || ""),
      })),
    },
    200,
    cors,
  );
}

async function handleAdminCreateAccount(request, env, serverId, actorSession, cors) {
  let body;
  try {
    body = await request.json();
  } catch (error) {
    return jsonResponse({ ok: false, error: "INVALID_JSON" }, 400, cors);
  }

  const username = normalizeAccountName(body && body.username);
  const role = normalizeAccountRole(body && body.role);
  const password = String(body && body.password ? body.password : "");
  const enabled = body && body.enabled === false ? 0 : 1;

  if (!username) {
    return jsonResponse({ ok: false, error: "INVALID_USERNAME" }, 400, cors);
  }
  if (!role) {
    return jsonResponse({ ok: false, error: "INVALID_ROLE" }, 400, cors);
  }
  if (password.length < 4) {
    return jsonResponse({ ok: false, error: "INVALID_PASSWORD" }, 400, cors);
  }

  const existing = await env.DB.prepare(
    "SELECT username FROM auth_accounts WHERE server_id = ? AND username = ? LIMIT 1",
  )
    .bind(serverId, username)
    .first();
  if (existing) {
    return jsonResponse({ ok: false, error: "ACCOUNT_EXISTS" }, 409, cors);
  }

  const passwordHash = await hashLoginPassword(serverId, password);
  const nowIso = new Date().toISOString();

  await env.DB.prepare(
    "INSERT INTO auth_accounts (server_id, username, role, password_hash, enabled, created_at, updated_at, created_by, updated_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
  )
    .bind(
      serverId,
      username,
      role,
      passwordHash,
      enabled,
      nowIso,
      nowIso,
      actorSession.username,
      actorSession.username,
    )
    .run();

  await insertAuditLog(env, {
    serverId,
    actorUsername: actorSession.username,
    actorRole: actorSession.role,
    action: "account_created",
    targetType: "account",
    targetId: username,
    summary: role,
    details: JSON.stringify({ username, role, enabled: Boolean(enabled) }),
    createdAt: nowIso,
  });

  return jsonResponse(
    {
      ok: true,
      account: {
        username,
        role,
        enabled: Boolean(enabled),
      },
    },
    200,
    cors,
  );
}

async function handleAdminUpdateAccount(request, env, serverId, targetUsername, actorSession, cors) {
  if (!targetUsername) {
    return jsonResponse({ ok: false, error: "INVALID_USERNAME" }, 400, cors);
  }

  let body;
  try {
    body = await request.json();
  } catch (error) {
    return jsonResponse({ ok: false, error: "INVALID_JSON" }, 400, cors);
  }

  const fields = [];
  const binds = [];
  const details = {};

  if (body && Object.prototype.hasOwnProperty.call(body, "role")) {
    const role = normalizeAccountRole(body.role);
    if (!role) {
      return jsonResponse({ ok: false, error: "INVALID_ROLE" }, 400, cors);
    }
    fields.push("role = ?");
    binds.push(role);
    details.role = role;
  }

  if (body && Object.prototype.hasOwnProperty.call(body, "enabled")) {
    const enabled = body.enabled ? 1 : 0;
    if (targetUsername === actorSession.username && enabled === 0) {
      return jsonResponse({ ok: false, error: "CANNOT_DISABLE_SELF" }, 400, cors);
    }
    fields.push("enabled = ?");
    binds.push(enabled);
    details.enabled = Boolean(enabled);
  }

  if (body && Object.prototype.hasOwnProperty.call(body, "password")) {
    const password = String(body.password || "");
    if (password.length < 4) {
      return jsonResponse({ ok: false, error: "INVALID_PASSWORD" }, 400, cors);
    }
    const passwordHash = await hashLoginPassword(serverId, password);
    fields.push("password_hash = ?");
    binds.push(passwordHash);
    details.passwordReset = true;
  }

  if (fields.length === 0) {
    return jsonResponse({ ok: false, error: "NO_UPDATES" }, 400, cors);
  }

  const nowIso = new Date().toISOString();
  fields.push("updated_at = ?");
  fields.push("updated_by = ?");
  binds.push(nowIso, actorSession.username);

  binds.push(serverId, targetUsername);

  const result = await env.DB.prepare(
    "UPDATE auth_accounts SET " + fields.join(", ") + " WHERE server_id = ? AND username = ?",
  )
    .bind(...binds)
    .run();
  if (!result || !result.meta || Number(result.meta.changes || 0) <= 0) {
    return jsonResponse({ ok: false, error: "ACCOUNT_NOT_FOUND" }, 404, cors);
  }

  await insertAuditLog(env, {
    serverId,
    actorUsername: actorSession.username,
    actorRole: actorSession.role,
    action: "account_updated",
    targetType: "account",
    targetId: targetUsername,
    summary: "account update",
    details: JSON.stringify(details),
    createdAt: nowIso,
  });

  return jsonResponse({ ok: true }, 200, cors);
}

async function handleAdminDeleteAccount(env, serverId, targetUsername, actorSession, cors) {
  if (!targetUsername) {
    return jsonResponse({ ok: false, error: "INVALID_USERNAME" }, 400, cors);
  }
  if (targetUsername === actorSession.username) {
    return jsonResponse({ ok: false, error: "CANNOT_DELETE_SELF" }, 400, cors);
  }

  const nowIso = new Date().toISOString();
  const result = await env.DB.prepare("DELETE FROM auth_accounts WHERE server_id = ? AND username = ?")
    .bind(serverId, targetUsername)
    .run();
  if (!result || !result.meta || Number(result.meta.changes || 0) <= 0) {
    return jsonResponse({ ok: false, error: "ACCOUNT_NOT_FOUND" }, 404, cors);
  }

  await env.DB.prepare("UPDATE auth_sessions SET revoked_at = ? WHERE server_id = ? AND username = ? AND revoked_at IS NULL")
    .bind(nowIso, serverId, targetUsername)
    .run();

  await insertAuditLog(env, {
    serverId,
    actorUsername: actorSession.username,
    actorRole: actorSession.role,
    action: "account_deleted",
    targetType: "account",
    targetId: targetUsername,
    summary: "account deleted",
    details: JSON.stringify({ username: targetUsername }),
    createdAt: nowIso,
  });

  return jsonResponse({ ok: true }, 200, cors);
}

async function handleAdminListAudit(env, serverId, searchParams, cors) {
  const rawLimit = Number(searchParams.get("limit"));
  const rawOffset = Number(searchParams.get("offset"));
  const limit = Number.isFinite(rawLimit) ? Math.max(1, Math.min(AUDIT_MAX_LIMIT, Math.floor(rawLimit))) : AUDIT_DEFAULT_LIMIT;
  const offset = Number.isFinite(rawOffset) ? Math.max(0, Math.floor(rawOffset)) : 0;

  const result = await env.DB.prepare(
    "SELECT id, actor_username, actor_role, action, target_type, target_id, summary, details, created_at FROM audit_logs WHERE server_id = ? ORDER BY id DESC LIMIT ? OFFSET ?",
  )
    .bind(serverId, limit, offset)
    .all();

  const rows = Array.isArray(result && result.results) ? result.results : [];
  return jsonResponse(
    {
      ok: true,
      serverId,
      limit,
      offset,
      logs: rows.map((row) => ({
        id: Number(row.id || 0),
        actorUsername: String(row.actor_username || ""),
        actorRole: String(row.actor_role || ""),
        action: String(row.action || ""),
        targetType: String(row.target_type || ""),
        targetId: String(row.target_id || ""),
        summary: String(row.summary || ""),
        details: safeJsonParse(String(row.details || "")) || String(row.details || ""),
        createdAt: String(row.created_at || ""),
      })),
    },
    200,
    cors,
  );
}

async function ensureBootstrapAccounts(env, serverId) {
  if (serverId !== "test") {
    return;
  }
  const countRow = await env.DB.prepare("SELECT COUNT(1) AS total FROM auth_accounts WHERE server_id = ?")
    .bind(serverId)
    .first();
  const total = Number(countRow && countRow.total ? countRow.total : 0);
  if (total > 0) {
    return;
  }

  const nowIso = new Date().toISOString();
  const seedAccounts = [
    { username: "testorder", role: "admin", password: String(env.TEST_ADMIN_PASSWORD || "testorder") },
    { username: "testmanager", role: "manager", password: String(env.TEST_MANAGER_PASSWORD || "testmanager") },
    { username: "teststaff", role: "staff", password: String(env.TEST_STAFF_PASSWORD || "teststaff") },
  ];

  for (const account of seedAccounts) {
    const passwordHash = await hashLoginPassword(serverId, account.password);
    await env.DB.prepare(
      "INSERT INTO auth_accounts (server_id, username, role, password_hash, enabled, created_at, updated_at, created_by, updated_by) VALUES (?, ?, ?, ?, 1, ?, ?, 'bootstrap', 'bootstrap')",
    )
      .bind(serverId, account.username, account.role, passwordHash, nowIso, nowIso)
      .run();
  }
}

async function writeAuditForStateChange(env, serverId, actorSession, previousPayload, nextPayload, createdAt) {
  const prevTasks = Array.isArray(previousPayload && previousPayload.tasks) ? previousPayload.tasks : [];
  const nextTasks = Array.isArray(nextPayload && nextPayload.tasks) ? nextPayload.tasks : [];
  const prevMap = buildTaskMap(prevTasks);
  const nextMap = buildTaskMap(nextTasks);
  const allIds = new Set([...Object.keys(prevMap), ...Object.keys(nextMap)]);

  for (const taskId of allIds) {
    const prevTask = prevMap[taskId] || null;
    const nextTask = nextMap[taskId] || null;

    if (!prevTask && nextTask) {
      await insertAuditLog(env, {
        serverId,
        actorUsername: actorSession.username,
        actorRole: actorSession.role,
        action: "task_created",
        targetType: "task",
        targetId: taskId,
        summary: normalizeTaskTitle(nextTask),
        details: JSON.stringify({ next: buildAuditTaskSnapshot(nextTask) }),
        createdAt,
      });
      continue;
    }

    if (prevTask && !nextTask) {
      await insertAuditLog(env, {
        serverId,
        actorUsername: actorSession.username,
        actorRole: actorSession.role,
        action: "task_deleted",
        targetType: "task",
        targetId: taskId,
        summary: normalizeTaskTitle(prevTask),
        details: JSON.stringify({ previous: buildAuditTaskSnapshot(prevTask) }),
        createdAt,
      });
      continue;
    }

    if (!prevTask || !nextTask) {
      continue;
    }

    const prevFingerprint = buildTaskFingerprint(prevTask);
    const nextFingerprint = buildTaskFingerprint(nextTask);
    if (prevFingerprint === nextFingerprint) {
      continue;
    }

    const statusAction = resolveTaskStatusAuditAction(prevTask, nextTask);
    if (statusAction) {
      await insertAuditLog(env, {
        serverId,
        actorUsername: actorSession.username,
        actorRole: actorSession.role,
        action: statusAction,
        targetType: "task",
        targetId: taskId,
        summary: normalizeTaskTitle(nextTask),
        details: JSON.stringify({ previousStatus: prevTask.status || "", nextStatus: nextTask.status || "" }),
        createdAt,
      });
    }

    const pinAction = resolveTaskPinAuditAction(prevTask, nextTask);
    if (pinAction) {
      await insertAuditLog(env, {
        serverId,
        actorUsername: actorSession.username,
        actorRole: actorSession.role,
        action: pinAction,
        targetType: "task",
        targetId: taskId,
        summary: normalizeTaskTitle(nextTask),
        details: JSON.stringify({ previousPinned: Boolean(prevTask.pinned), nextPinned: Boolean(nextTask.pinned) }),
        createdAt,
      });
    }

    await insertAuditLog(env, {
      serverId,
      actorUsername: actorSession.username,
      actorRole: actorSession.role,
      action: "task_updated",
      targetType: "task",
      targetId: taskId,
      summary: normalizeTaskTitle(nextTask),
      details: JSON.stringify({
        previous: buildAuditTaskSnapshot(prevTask),
        next: buildAuditTaskSnapshot(nextTask),
      }),
      createdAt,
    });
  }

  const prevOverview = safeStableStringify(previousPayload && previousPayload.todayOverview ? previousPayload.todayOverview : {});
  const nextOverview = safeStableStringify(nextPayload && nextPayload.todayOverview ? nextPayload.todayOverview : {});
  if (prevOverview !== nextOverview) {
    await insertAuditLog(env, {
      serverId,
      actorUsername: actorSession.username,
      actorRole: actorSession.role,
      action: "today_overview_updated",
      targetType: "today_overview",
      targetId: serverId,
      summary: "today overview",
      details: JSON.stringify({
        previous: previousPayload && previousPayload.todayOverview ? previousPayload.todayOverview : {},
        next: nextPayload && nextPayload.todayOverview ? nextPayload.todayOverview : {},
      }),
      createdAt,
    });
  }
}

function resolveTaskStatusAuditAction(previousTask, nextTask) {
  const previousStatus = String(previousTask && previousTask.status ? previousTask.status : "");
  const nextStatus = String(nextTask && nextTask.status ? nextTask.status : "");
  if (previousStatus === nextStatus) {
    return "";
  }
  if (nextStatus === "done") {
    return "task_completed";
  }
  if (previousStatus === "done" && nextStatus === "pending") {
    return "task_reopened";
  }
  return "task_status_changed";
}

function resolveTaskPinAuditAction(previousTask, nextTask) {
  const previousPinned = Boolean(previousTask && previousTask.pinned);
  const nextPinned = Boolean(nextTask && nextTask.pinned);
  if (previousPinned === nextPinned) {
    return "";
  }
  return nextPinned ? "task_pinned" : "task_unpinned";
}

function buildTaskMap(tasks) {
  const map = {};
  const list = Array.isArray(tasks) ? tasks : [];
  for (const task of list) {
    if (!task || typeof task !== "object") {
      continue;
    }
    const taskId = String(task.id || "").trim();
    if (!taskId) {
      continue;
    }
    map[taskId] = task;
  }
  return map;
}

function buildTaskFingerprint(task) {
  return safeStableStringify(buildAuditTaskSnapshot(task));
}

function buildAuditTaskSnapshot(task) {
  const source = task && typeof task === "object" ? task : {};
  return {
    id: String(source.id || ""),
    title: String(source.title || ""),
    category: String(source.category || ""),
    subcategory: String(source.subcategory || ""),
    owner: String(source.owner || ""),
    completedBy: String(source.completedBy || ""),
    status: String(source.status || ""),
    pinned: Boolean(source.pinned),
    allDay: Boolean(source.allDay),
    startAt: source.startAt || null,
    endAt: source.endAt || null,
    description: String(source.description || ""),
  };
}

function normalizeTaskTitle(task) {
  const title = task && typeof task === "object" ? String(task.title || "").trim() : "";
  return title || "-";
}

async function insertAuditLog(env, entry) {
  const serverId = normalizeServerId(entry && entry.serverId);
  const actorUsername = normalizeAccountName(entry && entry.actorUsername);
  const actorRole = normalizeAccountRole(entry && entry.actorRole);
  const action = String(entry && entry.action ? entry.action : "").trim();
  const targetType = String(entry && entry.targetType ? entry.targetType : "").trim();
  if (!serverId || !actorUsername || !actorRole || !action || !targetType) {
    return;
  }
  const targetId = String(entry && entry.targetId ? entry.targetId : "").trim();
  const summary = String(entry && entry.summary ? entry.summary : "").slice(0, 400);
  const details = String(entry && entry.details ? entry.details : "").slice(0, 12000);
  const createdAt = String(entry && entry.createdAt ? entry.createdAt : "").trim() || new Date().toISOString();

  await env.DB.prepare(
    "INSERT INTO audit_logs (server_id, actor_username, actor_role, action, target_type, target_id, summary, details, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
  )
    .bind(serverId, actorUsername, actorRole, action, targetType, targetId, summary, details, createdAt)
    .run();
}

function toDbBoolean(value) {
  return Number(value) === 1 || value === true || String(value).trim() === "1";
}

function normalizeAccountRole(value) {
  const role = String(value || "")
    .trim()
    .toLowerCase();
  return ACCOUNT_ROLES.has(role) ? role : "";
}

function normalizeAccountName(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]/g, "");
}

function getBearerToken(request) {
  const raw = String(request.headers.get("authorization") || "").trim();
  if (!/^bearer\s+/i.test(raw)) {
    return "";
  }
  return raw.replace(/^bearer\s+/i, "").trim();
}

function createSessionToken() {
  const bytes = new Uint8Array(24);
  crypto.getRandomValues(bytes);
  return toBase64Url(bytes);
}

function toBase64Url(bytes) {
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

async function hashLoginPassword(serverId, password) {
  return sha256Hex("server:" + normalizeServerId(serverId) + "::password:" + String(password || ""));
}

async function sha256Hex(value) {
  const data = new TextEncoder().encode(String(value || ""));
  const digest = await crypto.subtle.digest("SHA-256", data);
  const bytes = new Uint8Array(digest);
  let output = "";
  for (const byte of bytes) {
    output += byte.toString(16).padStart(2, "0");
  }
  return output;
}

function safeStableStringify(value) {
  return JSON.stringify(sortJsonValue(value));
}

function sortJsonValue(value) {
  if (Array.isArray(value)) {
    return value.map(sortJsonValue);
  }
  if (!value || typeof value !== "object") {
    return value;
  }
  const output = {};
  Object.keys(value)
    .sort()
    .forEach((key) => {
      output[key] = sortJsonValue(value[key]);
    });
  return output;
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
