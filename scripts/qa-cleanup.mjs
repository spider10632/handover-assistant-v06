#!/usr/bin/env node

const DEFAULT_BASE_URL = "https://handover-cloud-api.spider10632.workers.dev";

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const serverId = String(args.server || "test").trim().toLowerCase();
  const prefix = String(args.prefix || "QA_TEST_").trim();
  const baseUrl = normalizeBaseUrl(args.base || process.env.CLOUD_API_BASE || DEFAULT_BASE_URL);

  if (!serverId) {
    throw new Error("Missing --server");
  }
  if (!prefix) {
    throw new Error("Missing --prefix");
  }
  if (!baseUrl) {
    throw new Error("Missing --base or CLOUD_API_BASE");
  }

  const state = await fetchCloudState(baseUrl, serverId);
  if (!state || !state.payload || !Array.isArray(state.payload.tasks)) {
    console.log("[qa-cleanup] no task list found, nothing to clean.");
    return;
  }

  const originalTasks = state.payload.tasks;
  const nextTasks = originalTasks.filter((task) => !isQaTask(task, prefix));
  const removed = originalTasks.length - nextTasks.length;

  if (removed <= 0) {
    console.log("[qa-cleanup] no matching QA tasks.");
    return;
  }

  const nextPayload = {
    ...state.payload,
    tasks: nextTasks,
    updatedAt: new Date().toISOString(),
  };
  await putCloudState(baseUrl, serverId, nextPayload);
  console.log("[qa-cleanup] removed", removed, "task(s) with prefix", prefix);
}

function parseArgs(argv) {
  const output = {};
  for (let i = 0; i < argv.length; i += 1) {
    const token = String(argv[i] || "");
    if (!token.startsWith("--")) {
      continue;
    }
    const key = token.slice(2);
    const value = argv[i + 1] && !String(argv[i + 1]).startsWith("--") ? argv[i + 1] : "";
    output[key] = value;
    if (value) {
      i += 1;
    }
  }
  return output;
}

function normalizeBaseUrl(value) {
  return String(value || "").trim().replace(/\/+$/, "");
}

function isQaTask(task, prefix) {
  if (!task || typeof task !== "object") {
    return false;
  }
  const needle = String(prefix || "");
  const title = String(task.title || "");
  const description = String(task.description || "");
  return title.includes(needle) || description.includes(needle);
}

async function fetchCloudState(baseUrl, serverId) {
  const url = baseUrl + "/v1/state/" + encodeURIComponent(serverId);
  const response = await fetch(url, {
    method: "GET",
    headers: {
      accept: "application/json",
    },
  });
  if (response.status === 404) {
    return null;
  }
  if (!response.ok) {
    const detail = await response.text();
    throw new Error("GET state failed: " + response.status + " " + detail);
  }
  return response.json();
}

async function putCloudState(baseUrl, serverId, payload) {
  const url = baseUrl + "/v1/state/" + encodeURIComponent(serverId);
  const response = await fetch(url, {
    method: "PUT",
    headers: {
      "content-type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const detail = await response.text();
    throw new Error("PUT state failed: " + response.status + " " + detail);
  }
  return response.json();
}

main().catch((error) => {
  console.error("[qa-cleanup] failed:", error && error.message ? error.message : error);
  process.exitCode = 1;
});
