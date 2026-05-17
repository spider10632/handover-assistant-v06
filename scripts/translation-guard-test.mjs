#!/usr/bin/env node

const DEFAULT_BASE_URL = "https://handover-cloud-api.spider10632.workers.dev";

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const serverId = String(args.server || "test").trim().toLowerCase();
  const baseUrl = normalizeBaseUrl(args.base || process.env.CLOUD_API_BASE || DEFAULT_BASE_URL);
  const qaPrefix = String(args.prefix || "QA_TEST_").trim();
  if (!serverId || !baseUrl || !qaPrefix) {
    throw new Error("Usage: node scripts/translation-guard-test.mjs --server test --base <url> --prefix QA_TEST_");
  }

  const runPrefix = qaPrefix + Date.now();
  const cleanupPrefix = qaPrefix;

  try {
    await insertQaMarkerTasks(baseUrl, serverId, runPrefix);
    await runTranslationGuardCases(baseUrl, serverId);
    console.log("[translation-guard-test] all assertions passed.");
  } finally {
    try {
      const removed = await cleanupQaTasks(baseUrl, serverId, cleanupPrefix);
      console.log("[translation-guard-test] cleanup removed", removed, "QA task(s).");
      const remain = await countQaTasks(baseUrl, serverId, cleanupPrefix);
      if (remain > 0) {
        throw new Error("cleanup incomplete, remaining QA tasks: " + remain);
      }
    } catch (cleanupError) {
      console.error("[translation-guard-test] cleanup failed:", cleanupError && cleanupError.message ? cleanupError.message : cleanupError);
      process.exitCode = 1;
    }
  }
}

async function runTranslationGuardCases(baseUrl, serverId) {
  const cases = [
    {
      targetLang: "en",
      source: "客人不在房間，請先跟櫃檯拿鑰匙。",
      mustContain: ["guest", "front desk"],
      forbid: [/\bif\s+you\b/i, /^\s*you\b/i],
    },
    {
      targetLang: "en",
      source: "請協助客人下行李，完成後回報櫃檯。",
      mustContain: ["guest", "luggage pickup", "front desk"],
      forbid: [/\byour\b/i],
    },
    {
      targetLang: "en",
      source: "此交接事項由夜班處理，請勿刪除原始紀錄。",
      mustContain: ["handover"],
      forbid: [],
    },
    {
      targetLang: "zh",
      source: "If the guest is not in the room, please get the key from the front desk first.",
      mustContain: ["客人", "櫃檯"],
      forbid: [/^\s*你/, /如果你/],
    },
    {
      targetLang: "zh",
      source: "Please complete the handover log before 18:00.",
      mustContain: ["交接"],
      forbid: [],
    },
    {
      targetLang: "zh",
      source: "Guest requests luggage pickup at 20:00.",
      mustContain: ["客人", "取行李"],
      forbid: [],
    },
  ];

  const grouped = groupCasesByTarget(cases);
  for (const targetLang of Object.keys(grouped)) {
    const caseList = grouped[targetLang];
    const sources = caseList.map((item) => item.source);
    const response = await requestTranslate(baseUrl, serverId, targetLang, sources);
    const translations = Array.isArray(response.translations) ? response.translations : [];
    const quality = Array.isArray(response.quality) ? response.quality : [];
    if (translations.length !== caseList.length) {
      throw new Error("translation length mismatch for " + targetLang);
    }

    caseList.forEach((testCase, index) => {
      const translated = String(translations[index] || "").trim();
      const qualityInfo = quality[index] && typeof quality[index] === "object" ? quality[index] : { ok: true };
      assertTranslationCase(testCase, translated, qualityInfo);
      console.log(
        "[translation-guard-test]",
        targetLang.toUpperCase(),
        "OK:",
        shorten(testCase.source, 40),
        "=>",
        shorten(translated, 70),
      );
    });
  }
}

function assertTranslationCase(testCase, translated, qualityInfo) {
  if (!translated) {
    throw new Error("empty translation: " + testCase.source);
  }
  if (qualityInfo.ok === false) {
    throw new Error("quality guard rejected case: " + testCase.source + " reason=" + String(qualityInfo.reason || ""));
  }
  for (const token of testCase.mustContain || []) {
    if (!containsToken(translated, token)) {
      throw new Error("missing token [" + token + "] in translation: " + translated);
    }
  }
  for (const pattern of testCase.forbid || []) {
    if (pattern.test(translated)) {
      throw new Error("forbidden pattern " + pattern + " in translation: " + translated);
    }
  }
}

function containsToken(text, token) {
  const value = String(text || "");
  const key = String(token || "");
  if (!key) {
    return true;
  }
  if (/[A-Za-z]/.test(key)) {
    return value.toLowerCase().includes(key.toLowerCase());
  }
  return value.includes(key);
}

function groupCasesByTarget(cases) {
  const grouped = {};
  for (const item of cases) {
    const target = String(item.targetLang || "").trim().toLowerCase();
    if (!target) {
      continue;
    }
    if (!grouped[target]) {
      grouped[target] = [];
    }
    grouped[target].push(item);
  }
  return grouped;
}

function shorten(text, max) {
  const value = String(text || "");
  const size = Number(max) > 0 ? Number(max) : 40;
  return value.length > size ? value.slice(0, size - 1) + "…" : value;
}

async function requestTranslate(baseUrl, serverId, targetLang, texts) {
  const url = baseUrl + "/v1/translate/" + encodeURIComponent(serverId);
  const response = await fetch(url, {
    method: "POST",
    headers: {
      "content-type": "application/json",
    },
    body: JSON.stringify({
      targetLang,
      texts,
    }),
  });
  const raw = await response.text();
  let parsed = null;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    // ignore
  }
  if (!response.ok) {
    throw new Error("translate request failed: " + response.status + " " + raw);
  }
  if (!parsed || !Array.isArray(parsed.translations)) {
    throw new Error("invalid translate response");
  }
  return parsed;
}

async function insertQaMarkerTasks(baseUrl, serverId, runPrefix) {
  const state = await fetchCloudState(baseUrl, serverId);
  if (!state || !state.payload || typeof state.payload !== "object") {
    return;
  }
  const payload = state.payload;
  const tasks = Array.isArray(payload.tasks) ? payload.tasks.slice() : [];
  const now = new Date();
  for (let i = 0; i < 2; i += 1) {
    const at = new Date(now.getTime() + i * 60000).toISOString();
    tasks.push({
      id: buildQaId(runPrefix, i),
      category: "公告",
      subcategory: "",
      title: runPrefix + "_TASK_" + (i + 1),
      owner: "qa",
      assignee: "",
      completedBy: "",
      description: runPrefix + "_DESCRIPTION_" + (i + 1),
      translations: {},
      startAt: at,
      endAt: null,
      dueAt: at,
      allDay: false,
      status: "pending",
      pinned: false,
      remindedAt: null,
      assigneeRemind5At: null,
      assigneeRemind10At: null,
      createdAt: at,
      updatedAt: at,
    });
  }
  await putCloudState(baseUrl, serverId, {
    ...payload,
    tasks,
    updatedAt: new Date().toISOString(),
  });
}

function buildQaId(prefix, index) {
  return String(prefix || "QA_TEST") + "_ID_" + String(index || 0);
}

async function cleanupQaTasks(baseUrl, serverId, prefix) {
  const state = await fetchCloudState(baseUrl, serverId);
  if (!state || !state.payload || !Array.isArray(state.payload.tasks)) {
    return 0;
  }
  const tasks = state.payload.tasks;
  const nextTasks = tasks.filter((task) => !isQaTask(task, prefix));
  const removed = tasks.length - nextTasks.length;
  if (removed <= 0) {
    return 0;
  }
  await putCloudState(baseUrl, serverId, {
    ...state.payload,
    tasks: nextTasks,
    updatedAt: new Date().toISOString(),
  });
  return removed;
}

async function countQaTasks(baseUrl, serverId, prefix) {
  const state = await fetchCloudState(baseUrl, serverId);
  if (!state || !state.payload || !Array.isArray(state.payload.tasks)) {
    return 0;
  }
  return state.payload.tasks.filter((task) => isQaTask(task, prefix)).length;
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

main().catch((error) => {
  console.error("[translation-guard-test] failed:", error && error.message ? error.message : error);
  process.exitCode = 1;
});
