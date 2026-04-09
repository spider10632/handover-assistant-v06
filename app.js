(function () {
  "use strict";

  const STORAGE_KEY = "handover_tasks_v1";
  const TODAY_OVERVIEW_KEY = "handover_today_overview_v1";
  const DELETED_TASK_IDS_KEY = "handover_deleted_task_ids_v1";
  const USER_SERVER_MAP = Object.freeze({
    caesarmetro: {
      serverId: "caesarmetro",
      displayName: "caesarmetro",
      cloudApiBase: "",
      legacyKvdbBucket: "7cWAAYbjUk95gHksRVkTp3",
      legacyKvdbKey: "handover-assistant-v07-caesarmetro",
    },
    test: {
      serverId: "test",
      displayName: "test",
      cloudApiBase: "",
      legacyKvdbBucket: "",
      legacyKvdbKey: "",
    },
  });
  const ALLOW_DYNAMIC_SERVER_LOGIN = true;
  const DEFAULT_SERVER_ID = "caesarmetro";
  const CLOUD_API_BASE_OVERRIDE =
    typeof window !== "undefined" && typeof window.HANDOVER_CLOUD_API_BASE === "string"
      ? window.HANDOVER_CLOUD_API_BASE.trim()
      : "";
  const CLOUD_REQUEST_TIMEOUT_MS = 12000;
  const CLOUD_PUSH_DEBOUNCE_MS = 1200;
  const LEGACY_MIGRATION_DONE_KEY = "handover_legacy_kvdb_migration_done_v1";
  const BACKUP_TYPE = "handover-backup";
  const BACKUP_VERSION = "0.9";
  const REMINDER_CHECK_MS = 30 * 1000;
  const COUNTDOWN_REFRESH_MS = 1000;
  const TOAST_MS = 3000;
  const CATEGORIES = ["廣場", "包裹代收", "車輛安排", "大廳", "會議室", "團桌", "客房", "預訂", "餐飲部", "待回覆信件", "郵寄", "行政", "公告", "遺留物"];
  const SUBCATEGORY_MAP = {
    廣場: ["保留車位", "其他"],
    包裹代收: ["團體", "散客", "其他"],
    車輛安排: ["禮賓車", "計程車", "其他"],
    大廳: ["行李寄放", "其他"],
    客房: ["下行李", "房務相關事項", "送房", "佈置", "其他"],
    預訂: ["餐廳", "車票", "其他"],
    行政: ["叫貨領貨", "人事相關", "其他"],
    遺留物: ["待寄", "待取"],
  };
  const EXPORT_FONT_EAST_ASIA = "DFKai-SB";
  const EXPORT_FONT_LATIN = "DFKai-SB";
  const EXPORT_DEFAULT_COLOR = "1F2A2A";
  const EXPORT_CATEGORY_COLORS = Object.freeze({
    廣場: "B42318",
    包裹代收: "1D4ED8",
    車輛安排: "0F766E",
    大廳: "A16207",
    會議室: "6D28D9",
    團桌: "C2410C",
    客房: "0E7490",
    預訂: "7C3AED",
    餐飲部: "9A3412",
    待回覆信件: "1E40AF",
    郵寄: "7A5A00",
    行政: "7C3F00",
    公告: "334155",
    遺留物: "6B4F2A",
  });

  const state = {
    tasks: [],
    queryDate: "",
    queryStatus: "all",
    queryCategory: "all",
    queryKeyword: "",
    taskListFilter: "all",
    todayOverview: {
      checkin: "",
      checkout: "",
      occupancy: "",
      dateKey: "",
    },
    currentDayKey: "",
    todayCategory: "all",
    formAllDay: false,
    editingTaskId: null,
    reminderTimer: null,
    countdownTimer: null,
    currentServerId: null,
    currentServerConfig: null,
    cloudInitDone: false,
    cloudPushTimer: null,
    cloudPushInFlight: false,
    cloudPushQueued: false,
    cloudMutePush: false,
    cloudLastErrorAt: 0,
    deletedTaskIds: {},
    toastTimer: null,
    initialized: false,
  };

  const els = {};

  document.addEventListener("DOMContentLoaded", bootstrap);

  function bootstrap() {
    cacheElements();
    initAccessGate();
  }

  function init() {
    if (state.initialized) {
      return;
    }
    ensureServerContext();
    state.initialized = true;
    state.tasks = loadTasks();
    state.deletedTaskIds = loadDeletedTaskIds();
    {
      const filtered = filterTasksByDeletedMap(state.tasks, state.deletedTaskIds);
      state.tasks = filtered.tasks;
      state.deletedTaskIds = filtered.deletedTaskIds;
    }
    saveTasks();
    saveDeletedTaskIds();
    state.todayOverview = loadTodayOverview();
    state.currentDayKey = getTodayDateKey();
    bindEvents();
    syncCollapsiblePanels();
    setupQueryCategoryOptions();
    if (els.queryCategory) {
      els.queryCategory.value = state.queryCategory;
    }
    if (els.queryKeyword) {
      els.queryKeyword.value = state.queryKeyword;
    }
    if (els.taskListFilter) {
      state.taskListFilter = normalizeTaskListFilter(els.taskListFilter.value);
      els.taskListFilter.value = state.taskListFilter;
    }
    state.todayCategory = "all";
    if (els.todayCategoryFilter) {
      els.todayCategoryFilter.value = "all";
    }
    updateSubcategoryOptions();
    updateFormLockState();
    setDefaultDueTime();
    renderTodayOverviewBar();
    renderAll();
    startReminderLoop();
    startCountdownLoop();
    initCloudSync();
  }

  function initAccessGate() {
    if (!els.passwordGate || !els.passwordForm || !els.passwordInput) {
      init();
      return;
    }
    document.body.classList.add("gate-locked");
    els.passwordGate.classList.remove("hidden");
    els.passwordForm.addEventListener("submit", handlePasswordSubmit);
    els.passwordInput.addEventListener("input", clearPasswordError);
    setTimeout(function () {
      els.passwordInput.focus();
    }, 40);
  }

  function normalizeServerInput(value) {
    return String(value || "").trim().toLowerCase();
  }

  function resolveServerByInput(value) {
    const key = normalizeServerInput(value);
    if (!key) {
      return null;
    }
    const server = USER_SERVER_MAP[key];
    if (server && typeof server === "object") {
      return server;
    }
    return buildDynamicServerConfig(key);
  }

  function buildDynamicServerConfig(value) {
    if (!ALLOW_DYNAMIC_SERVER_LOGIN) {
      return null;
    }
    const key = normalizeServerInput(value);
    if (!/^[a-z0-9][a-z0-9_-]{2,31}$/.test(key)) {
      return null;
    }
    return {
      serverId: key,
      displayName: key,
      cloudApiBase: "",
      legacyKvdbBucket: "",
      legacyKvdbKey: "",
    };
  }

  function ensureServerContext() {
    if (state.currentServerId) {
      return;
    }
    const fallback = resolveServerByInput(DEFAULT_SERVER_ID);
    state.currentServerConfig = fallback;
    state.currentServerId = fallback && fallback.serverId ? fallback.serverId : DEFAULT_SERVER_ID;
  }

  function getScopedStorageKey(baseKey) {
    const serverId = state.currentServerId || DEFAULT_SERVER_ID;
    return baseKey + "__" + serverId;
  }

  function getCurrentServerConfig() {
    const serverId = state.currentServerId || DEFAULT_SERVER_ID;
    if (
      state.currentServerConfig &&
      typeof state.currentServerConfig === "object" &&
      state.currentServerConfig.serverId === serverId
    ) {
      return state.currentServerConfig;
    }
    const mapped = USER_SERVER_MAP[serverId];
    if (mapped && typeof mapped === "object") {
      state.currentServerConfig = mapped;
      return mapped;
    }
    const dynamic = buildDynamicServerConfig(serverId);
    if (dynamic) {
      state.currentServerConfig = dynamic;
      return dynamic;
    }
    return null;
  }

  function normalizeCloudApiBase(baseUrl) {
    return String(baseUrl || "")
      .trim()
      .replace(/\/+$/, "");
  }

  function getCloudApiBase(server) {
    const overrideBase = normalizeCloudApiBase(CLOUD_API_BASE_OVERRIDE);
    if (overrideBase) {
      return overrideBase;
    }
    return normalizeCloudApiBase(server && server.cloudApiBase);
  }

  function buildCloudStateUrl(baseUrl, serverId) {
    const base = normalizeCloudApiBase(baseUrl);
    if (!base || !serverId) {
      return "";
    }
    return base + "/v1/state/" + encodeURIComponent(serverId);
  }

  function getCurrentCloudUrl() {
    const server = getCurrentServerConfig();
    if (!server) {
      return "";
    }
    const cloudBase = getCloudApiBase(server);
    return buildCloudStateUrl(cloudBase, server.serverId);
  }

  function getLegacyKvdbUrl(server) {
    if (!server || !server.legacyKvdbBucket || !server.legacyKvdbKey) {
      return "";
    }
    return (
      "https://kvdb.io/" +
      encodeURIComponent(server.legacyKvdbBucket) +
      "/" +
      encodeURIComponent(server.legacyKvdbKey)
    );
  }

  function getLegacyMigrationFlagKey(serverId) {
    const id = String(serverId || DEFAULT_SERVER_ID).trim() || DEFAULT_SERVER_ID;
    return LEGACY_MIGRATION_DONE_KEY + "__" + id;
  }

  function hasMigratedLegacyKvdb(serverId) {
    try {
      return localStorage.getItem(getLegacyMigrationFlagKey(serverId)) === "1";
    } catch (error) {
      return false;
    }
  }

  function markLegacyMigrationDone(serverId) {
    try {
      localStorage.setItem(getLegacyMigrationFlagKey(serverId), "1");
    } catch (error) {
      // ignore quota/private mode write failure
    }
  }

  function extractCloudPayload(data) {
    if (!data || typeof data !== "object") {
      return null;
    }
    if (data.payload && typeof data.payload === "object") {
      return data.payload;
    }
    if (data.data && typeof data.data === "object") {
      return data.data;
    }
    return data;
  }

  async function initCloudSync() {
    const url = getCurrentCloudUrl();
    if (!url) {
      state.cloudInitDone = true;
      showToast("尚未設定 Cloudflare API，暫用本機資料。");
      return;
    }
    try {
      const foundCloudData = await pullCloudBackupAndMerge();
      if (!foundCloudData) {
        await migrateLegacyKvdbToCloud();
      }
      state.cloudInitDone = true;
      await pushCloudBackupNow();
      showToast("已連線雲端資料庫。");
    } catch (error) {
      console.error("initCloudSync error", error);
      state.cloudInitDone = true;
      scheduleCloudPush(0);
      if (Date.now() - state.cloudLastErrorAt > 10000) {
        state.cloudLastErrorAt = Date.now();
        showToast("雲端連線失敗，暫用本機資料。");
      }
    }
  }

  async function migrateLegacyKvdbToCloud() {
    const server = getCurrentServerConfig();
    const currentServerId = state.currentServerId || DEFAULT_SERVER_ID;
    if (hasMigratedLegacyKvdb(currentServerId)) {
      return false;
    }
    const legacyUrl = getLegacyKvdbUrl(server);
    if (!legacyUrl) {
      markLegacyMigrationDone(currentServerId);
      return false;
    }
    const response = await fetchWithTimeout(legacyUrl, { method: "GET", cache: "no-store" }, CLOUD_REQUEST_TIMEOUT_MS);
    if (response.status === 404) {
      markLegacyMigrationDone(currentServerId);
      return false;
    }
    if (!response.ok) {
      throw new Error("legacy kvdb pull failed: " + response.status);
    }
    const raw = await response.text();
    if (!raw) {
      markLegacyMigrationDone(currentServerId);
      return false;
    }
    const parsed = JSON.parse(raw);
    const imported = parseBackupPayload(extractCloudPayload(parsed));
    if (!imported || imported.serverId !== currentServerId) {
      markLegacyMigrationDone(currentServerId);
      return false;
    }
    const merged = mergeBackupState(
      {
        tasks: state.tasks,
        todayOverview: state.todayOverview,
        deletedTaskIds: state.deletedTaskIds,
      },
      imported,
    );
    applyImportedBackup(merged, true);
    markLegacyMigrationDone(currentServerId);
    showToast("已完成 kvdb 舊資料搬移。");
    return true;
  }

  async function pullCloudBackupAndMerge() {
    const url = getCurrentCloudUrl();
    if (!url) {
      return false;
    }
    const response = await fetchWithTimeout(url, { method: "GET", cache: "no-store" }, CLOUD_REQUEST_TIMEOUT_MS);
    if (response.status === 404) {
      return false;
    }
    if (!response.ok) {
      throw new Error("cloud pull failed: " + response.status);
    }
    const raw = await response.text();
    if (!raw) {
      return false;
    }
    const parsed = JSON.parse(raw);
    const imported = parseBackupPayload(extractCloudPayload(parsed));
    if (!imported) {
      return false;
    }
    const currentServerId = state.currentServerId || DEFAULT_SERVER_ID;
    if (imported.serverId !== currentServerId) {
      return false;
    }
    const merged = mergeBackupState(
      {
        tasks: state.tasks,
        todayOverview: state.todayOverview,
        deletedTaskIds: state.deletedTaskIds,
      },
        imported,
      );
    applyImportedBackup(merged, true);
    return true;
  }

  function scheduleCloudPush(delayMs) {
    const url = getCurrentCloudUrl();
    if (!url || !state.cloudInitDone || state.cloudMutePush) {
      return;
    }
    const wait = typeof delayMs === "number" ? Math.max(0, delayMs) : CLOUD_PUSH_DEBOUNCE_MS;
    if (state.cloudPushTimer) {
      clearTimeout(state.cloudPushTimer);
    }
    state.cloudPushTimer = setTimeout(function () {
      state.cloudPushTimer = null;
      pushCloudBackupNow();
    }, wait);
  }

  async function pushCloudBackupNow() {
    const url = getCurrentCloudUrl();
    if (!url || !state.cloudInitDone) {
      return false;
    }
    if (state.cloudPushInFlight) {
      state.cloudPushQueued = true;
      return false;
    }
    state.cloudPushInFlight = true;
    try {
      let merged = {
        tasks: state.tasks,
        todayOverview: state.todayOverview,
        deletedTaskIds: state.deletedTaskIds,
      };
      const currentServerId = state.currentServerId || DEFAULT_SERVER_ID;

      const existingResponse = await fetchWithTimeout(url, { method: "GET", cache: "no-store" }, CLOUD_REQUEST_TIMEOUT_MS);
      if (existingResponse.ok) {
        const existingRaw = await existingResponse.text();
        if (existingRaw) {
          const existingParsed = JSON.parse(existingRaw);
          const imported = parseBackupPayload(extractCloudPayload(existingParsed));
          if (imported && imported.serverId === currentServerId) {
            merged = mergeBackupState(merged, imported, { preferBaseOverview: true });
          }
        }
      } else if (existingResponse.status !== 404) {
        throw new Error("cloud read-before-write failed: " + existingResponse.status);
      }

      state.cloudMutePush = true;
      try {
        applyImportedBackup(merged, true);
      } finally {
        state.cloudMutePush = false;
      }

      const payload = {
        type: BACKUP_TYPE,
        version: BACKUP_VERSION,
        serverId: currentServerId,
        exportedAt: new Date().toISOString(),
        tasks: merged.tasks.slice(),
        deletedTaskIds: normalizeDeletedTaskIds(merged.deletedTaskIds),
        todayOverview: {
          checkin: normalizeTodayOverviewValue(merged.todayOverview && merged.todayOverview.checkin),
          checkout: normalizeTodayOverviewValue(merged.todayOverview && merged.todayOverview.checkout),
          occupancy: normalizeOccupancyRateValue(merged.todayOverview && merged.todayOverview.occupancy),
          dateKey: normalizeDateInputFromAny(merged.todayOverview && merged.todayOverview.dateKey) || getTodayDateKey(),
        },
      };

      const response = await fetchWithTimeout(
        url,
        {
          method: "PUT",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload),
        },
        CLOUD_REQUEST_TIMEOUT_MS,
      );
      if (!response.ok) {
        throw new Error("cloud push failed: " + response.status);
      }
      return true;
    } catch (error) {
      console.error("pushCloudBackupNow error", error);
      if (Date.now() - state.cloudLastErrorAt > 10000) {
        state.cloudLastErrorAt = Date.now();
        showToast("雲端儲存失敗，請稍後重試。");
      }
      return false;
    } finally {
      state.cloudPushInFlight = false;
      if (state.cloudPushQueued) {
        state.cloudPushQueued = false;
        scheduleCloudPush(400);
      }
    }
  }

  async function fetchWithTimeout(url, options, timeoutMs) {
    const timeout = Number.isFinite(timeoutMs) ? timeoutMs : CLOUD_REQUEST_TIMEOUT_MS;
    if (typeof AbortController !== "function") {
      return fetch(url, options || {});
    }
    const controller = new AbortController();
    const timer = setTimeout(function () {
      controller.abort();
    }, timeout);
    try {
      const requestOptions = Object.assign({}, options || {}, { signal: controller.signal });
      return await fetch(url, requestOptions);
    } finally {
      clearTimeout(timer);
    }
  }

  function handlePasswordSubmit(event) {
    event.preventDefault();
    const entered = normalizeServerInput(els.passwordInput.value || "");
    const server = resolveServerByInput(entered);
    if (!server) {
      showPasswordError("使用者錯誤（請用英數、-、_，至少 3 碼）。");
      return;
    }
    state.currentServerId = server.serverId;
    state.currentServerConfig = server;
    unlockAccessGate();
    init();
    showToast("已進入 " + server.displayName + " 伺服器。");
  }

  function unlockAccessGate() {
    document.body.classList.remove("gate-locked");
    if (els.passwordGate) {
      els.passwordGate.classList.add("hidden");
    }
    if (els.passwordInput) {
      els.passwordInput.value = "";
    }
    clearPasswordError();
  }

  function showPasswordError(message) {
    if (!els.passwordError) {
      return;
    }
    els.passwordError.textContent = message || "使用者錯誤。";
    els.passwordError.classList.remove("hidden");
  }

  function clearPasswordError() {
    if (!els.passwordError) {
      return;
    }
    els.passwordError.textContent = "";
    els.passwordError.classList.add("hidden");
  }

  function cacheElements() {
    els.passwordGate = document.getElementById("password-gate");
    els.passwordForm = document.getElementById("password-form");
    els.passwordInput = document.getElementById("password-input");
    els.passwordError = document.getElementById("password-error");
    els.taskForm = document.getElementById("task-form");
    els.taskCategory = document.getElementById("task-category");
    els.taskSubcategory = document.getElementById("task-subcategory");
    els.categoryTip = document.getElementById("category-tip");
    els.taskTitle = document.getElementById("task-title");
    els.taskOwner = document.getElementById("task-owner");
    els.taskStartDate = document.getElementById("task-start-date");
    els.taskEndDate = document.getElementById("task-end-date");
    els.taskStartAt = document.getElementById("task-start-at");
    els.taskEndAt = document.getElementById("task-end-at");
    els.allDayBtn = document.getElementById("all-day-btn");
    els.taskPinned = document.getElementById("task-pinned");
    els.taskDescription = document.getElementById("task-description");
    els.addTaskBtn = document.getElementById("add-task-btn");
    els.clearFormBtn = document.getElementById("clear-form-btn");
    els.cancelEditBtn = document.getElementById("cancel-edit-btn");
    els.queryDate = document.getElementById("query-date");
    els.queryStatus = document.getElementById("query-status");
    els.queryCategory = document.getElementById("query-category");
    els.queryKeyword = document.getElementById("query-keyword");
    els.searchBtn = document.getElementById("search-btn");
    els.clearSearchBtn = document.getElementById("clear-search-btn");
    els.queryResultText = document.getElementById("query-result-text");
    els.tableBody = document.getElementById("task-table-body");
    els.todayTaskSummary = document.getElementById("today-task-summary");
    els.todayPinnedList = document.getElementById("today-pinned-list") || document.querySelector(".today-pinned-list");
    els.todayTaskList = document.getElementById("today-task-list");
    els.todayCategoryFilter = document.getElementById("today-category-filter");
    els.upcomingBoard = document.getElementById("upcoming-board");
    els.upcomingSummary = document.getElementById("upcoming-summary");
    els.upcomingTaskList = document.getElementById("upcoming-task-list");
    els.exportWordBtn = document.getElementById("export-word-btn");
    els.exportExcelBtn = document.getElementById("export-excel-btn");
    els.toast = document.getElementById("toast");
    els.notificationToggle = document.getElementById("notification-toggle");
    els.requestNotificationBtn = document.getElementById("request-notification-btn");
    els.exportDate = document.getElementById("export-date");
    els.exportStatus = document.getElementById("export-status");
    els.taskListFilter = document.getElementById("task-list-filter");
    els.todayAutoDate = document.getElementById("today-auto-date");
    els.todayExpectedCheckin = document.getElementById("today-expected-checkin");
    els.todayExpectedCheckout = document.getElementById("today-expected-checkout");
    els.todayOccupancyRate = document.getElementById("today-occupancy-rate");
    els.todayOverviewSaveBtn = document.getElementById("today-overview-save-btn");
    els.panelToggleButtons = Array.prototype.slice.call(document.querySelectorAll(".panel-toggle-btn"));
  }

  function bindEvents() {
    els.taskForm.addEventListener("submit", handleAddTask);
    els.taskCategory.addEventListener("change", updateSubcategoryOptions);
    els.taskSubcategory.addEventListener("change", updateFormLockState);
    els.allDayBtn.addEventListener("click", toggleAllDayMode);
    els.clearFormBtn.addEventListener("click", handleClearForm);
    els.cancelEditBtn.addEventListener("click", cancelEditing);
    els.searchBtn.addEventListener("click", applyDateQuery);
    els.clearSearchBtn.addEventListener("click", clearDateQuery);
    if (els.queryStatus) {
      els.queryStatus.addEventListener("change", applyDateQuery);
    }
    if (els.queryCategory) {
      els.queryCategory.addEventListener("change", applyDateQuery);
    }
    if (els.queryKeyword) {
      els.queryKeyword.addEventListener("keydown", function (event) {
        if (event.key === "Enter") {
          event.preventDefault();
          applyDateQuery();
        }
      });
    }
    if (els.taskListFilter) {
      els.taskListFilter.addEventListener("change", handleTaskListFilterChange);
    }
    if (els.todayExpectedCheckin) {
      els.todayExpectedCheckin.addEventListener("input", handleTodayOverviewInput);
    }
    if (els.todayExpectedCheckout) {
      els.todayExpectedCheckout.addEventListener("input", handleTodayOverviewInput);
    }
    if (els.todayOccupancyRate) {
      els.todayOccupancyRate.addEventListener("change", handleTodayOccupancyCommit);
      els.todayOccupancyRate.addEventListener("blur", handleTodayOccupancyCommit);
    }
    if (els.todayOverviewSaveBtn) {
      els.todayOverviewSaveBtn.addEventListener("click", handleTodayOverviewForceSave);
    }
    els.tableBody.addEventListener("click", handleTableAction);
    if (els.todayTaskList) {
      els.todayTaskList.addEventListener("click", handleTodayListAction);
    }
    if (els.todayPinnedList) {
      els.todayPinnedList.addEventListener("click", handleTodayListAction);
    }
    els.todayCategoryFilter.addEventListener("change", handleTodayCategoryChange);
    els.exportWordBtn.addEventListener("click", handleExportWord);
    if (els.exportExcelBtn) {
      els.exportExcelBtn.addEventListener("click", handleExportExcel);
    }
    els.requestNotificationBtn.addEventListener("click", requestNotificationPermission);
    if (els.panelToggleButtons.length > 0) {
      els.panelToggleButtons.forEach(function (btn) {
        btn.addEventListener("click", handlePanelToggle);
      });
    }

    document.addEventListener("visibilitychange", function () {
      if (!document.hidden) {
        checkReminders();
      }
    });
  }

  function setDefaultDueTime() {
    setAllDayMode(false);
    if (els.taskStartDate) {
      els.taskStartDate.value = toDateKey(new Date());
    }
    if (els.taskEndDate) {
      els.taskEndDate.value = "";
    }
    els.taskStartAt.value = "";
    els.taskEndAt.value = "";
  }

  function apply24HourInputMode() {
    [els.taskStartAt, els.taskEndAt].forEach(function (input) {
      if (!input) {
        return;
      }
      if (input.type === "time") {
        input.setAttribute("step", "60");
        input.removeAttribute("inputmode");
        input.removeAttribute("maxlength");
        input.removeAttribute("pattern");
        input.removeAttribute("placeholder");
      } else if (input.type === "date") {
        input.removeAttribute("inputmode");
        input.removeAttribute("maxlength");
        input.removeAttribute("pattern");
        input.removeAttribute("placeholder");
      }
    });
  }

  function handleTodayOverviewInput() {
    state.todayOverview = normalizeTodayOverviewRecord(
      {
        checkin: els.todayExpectedCheckin ? els.todayExpectedCheckin.value : "",
        checkout: els.todayExpectedCheckout ? els.todayExpectedCheckout.value : "",
        occupancy: state.todayOverview ? state.todayOverview.occupancy : "",
        dateKey: getTodayDateKey(),
      },
      true,
    );
    saveTodayOverview();
  }

  function handleTodayOccupancyCommit() {
    state.todayOverview = normalizeTodayOverviewRecord(
      {
        checkin: state.todayOverview ? state.todayOverview.checkin : "",
        checkout: state.todayOverview ? state.todayOverview.checkout : "",
        occupancy: els.todayOccupancyRate ? els.todayOccupancyRate.value : "",
        dateKey: getTodayDateKey(),
      },
      true,
    );
    if (els.todayOccupancyRate) {
      els.todayOccupancyRate.value = state.todayOverview.occupancy;
    }
    saveTodayOverview();
  }

  async function handleTodayOverviewForceSave() {
    state.todayOverview = normalizeTodayOverviewRecord(
      {
        checkin: els.todayExpectedCheckin ? els.todayExpectedCheckin.value : "",
        checkout: els.todayExpectedCheckout ? els.todayExpectedCheckout.value : "",
        occupancy: els.todayOccupancyRate ? els.todayOccupancyRate.value : "",
        dateKey: getTodayDateKey(),
      },
      true,
    );
    renderTodayOverviewBar();
    saveTodayOverview();
    const cloudUrl = getCurrentCloudUrl();
    if (!cloudUrl || !state.cloudInitDone) {
      showToast("已儲存（本機）。");
      return;
    }
    const ok = await pushCloudBackupNow();
    showToast(ok ? "已儲存到雲端。" : "已儲存，本機成功，雲端稍後同步。");
  }

  function renderTodayOverviewBar() {
    handleDailyRollover();
    if (els.todayAutoDate) {
      els.todayAutoDate.textContent = formatDateOnly(new Date());
    }
    if (els.todayExpectedCheckin) {
      els.todayExpectedCheckin.value = state.todayOverview.checkin || "";
    }
    if (els.todayExpectedCheckout) {
      els.todayExpectedCheckout.value = state.todayOverview.checkout || "";
    }
    if (els.todayOccupancyRate) {
      els.todayOccupancyRate.value = state.todayOverview.occupancy || "";
    }
  }

  function handleDailyRollover() {
    const todayKey = getTodayDateKey();
    if (!state.currentDayKey) {
      state.currentDayKey = todayKey;
      state.todayOverview = normalizeTodayOverviewRecord(state.todayOverview, true);
      return;
    }
    if (state.currentDayKey === todayKey) {
      return;
    }
    const hadValue =
      Boolean(state.todayOverview && state.todayOverview.checkin) ||
      Boolean(state.todayOverview && state.todayOverview.checkout) ||
      Boolean(state.todayOverview && state.todayOverview.occupancy);
    state.currentDayKey = todayKey;
    state.todayOverview = createEmptyTodayOverview(todayKey);
    saveTodayOverview();
    if (hadValue) {
      showToast("已換日，預進/預退/住房率已清空。");
    }
  }

  function handleTaskListFilterChange() {
    state.taskListFilter = normalizeTaskListFilter(els.taskListFilter ? els.taskListFilter.value : "all");
    if (els.taskListFilter) {
      els.taskListFilter.value = state.taskListFilter;
    }
    renderTaskTable();
  }

  function syncCollapsiblePanels() {
    if (!els.panelToggleButtons || els.panelToggleButtons.length === 0) {
      return;
    }
    els.panelToggleButtons.forEach(function (btn) {
      const targetId = String(btn.dataset.target || "").trim();
      if (!targetId) {
        return;
      }
      const panel = document.getElementById(targetId);
      if (!panel) {
        return;
      }
      updatePanelToggleButton(btn, panel);
    });
  }

  function handlePanelToggle(event) {
    const btn = event.currentTarget;
    const targetId = String(btn.dataset.target || "").trim();
    if (!targetId) {
      return;
    }
    const panel = document.getElementById(targetId);
    if (!panel) {
      return;
    }
    panel.classList.toggle("is-collapsed");
    updatePanelToggleButton(btn, panel);
  }

  function setPanelCollapsed(panelId, collapsed) {
    const targetId = String(panelId || "").trim();
    if (!targetId) {
      return;
    }
    const panel = document.getElementById(targetId);
    if (!panel) {
      return;
    }
    if (collapsed) {
      panel.classList.add("is-collapsed");
    } else {
      panel.classList.remove("is-collapsed");
    }
    const btn = document.querySelector('.panel-toggle-btn[data-target="' + targetId + '"]');
    if (btn) {
      updatePanelToggleButton(btn, panel);
    }
  }

  function updatePanelToggleButton(btn, panel) {
    const isCollapsed = panel.classList.contains("is-collapsed");
    btn.textContent = isCollapsed ? "展開" : "收合";
    btn.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
  }

  function toggleAllDayMode() {
    setAllDayMode(!state.formAllDay);
  }

  function setAllDayMode(enabled) {
    const next = Boolean(enabled);
    const currentStartDate = String(els.taskStartDate ? els.taskStartDate.value || "" : "").trim();
    const currentEndDate = String(els.taskEndDate ? els.taskEndDate.value || "" : "").trim();
    const currentStart = String(els.taskStartAt.value || "").trim();
    const currentEnd = String(els.taskEndAt.value || "").trim();
    state.formAllDay = next;

    els.allDayBtn.setAttribute("aria-pressed", next ? "true" : "false");
    els.allDayBtn.classList.toggle("active", next);
    els.allDayBtn.textContent = next ? "全日中" : "全日";
    els.taskStartAt.disabled = next;
    els.taskEndAt.disabled = next;
    if (next) {
      els.taskStartAt.value = "";
      els.taskEndAt.value = "";
    } else {
      els.taskStartAt.value = normalizeTimeInputFromAny(currentStart, "start");
      els.taskEndAt.value = normalizeTimeInputFromAny(currentEnd, "end");
    }
    if (els.taskStartDate && !els.taskStartDate.value) {
      els.taskStartDate.value =
        normalizeDateInputFromAny(currentStartDate) ||
        normalizeDateInputFromAny(currentEndDate) ||
        toDateKey(new Date());
    }
    if (els.taskEndDate && !els.taskEndDate.value) {
      els.taskEndDate.value = normalizeDateInputFromAny(currentEndDate);
    }
    apply24HourInputMode();
  }

  function normalizeDateInputFromAny(value) {
    const currentValue = String(value || "").trim();
    if (!currentValue) {
      return "";
    }
    const normalizedDate = currentValue.replace(/\//g, "-");
    if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(normalizedDate)) {
      return normalizedDate.slice(0, 10);
    }
    const parsedMs = new Date(normalizedDate).getTime();
    if (!Number.isNaN(parsedMs)) {
      return toDateKey(new Date(parsedMs));
    }
    return "";
  }

  function normalizeTimeInputFromAny(value, part) {
    const currentValue = String(value || "").trim();
    if (!currentValue) {
      return "";
    }
    const direct = currentValue.match(/^(\d{1,2}):(\d{2})$/);
    if (direct) {
      const hh = Number(direct[1]);
      const mm = Number(direct[2]);
      if (hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59) {
        return String(hh).padStart(2, "0") + ":" + String(mm).padStart(2, "0");
      }
    }

    const parsedMs = parseTaskTimeInput(currentValue, false, part);
    if (!Number.isNaN(parsedMs)) {
      return formatTime(new Date(parsedMs), false);
    }
    if (normalizeDateInputFromAny(currentValue)) {
      return part === "end" ? "18:00" : "09:00";
    }
    return "";
  }

  function updateSubcategoryOptions() {
    const category = String(els.taskCategory.value || "").trim();
    const current = String(els.taskSubcategory.value || "").trim();

    if (!isValidCategory(category)) {
      els.taskSubcategory.innerHTML = '<option value="">請先選主分類</option>';
      els.taskSubcategory.disabled = true;
      els.taskSubcategory.required = false;
      updateFormLockState();
      return;
    }

    const options = getSubcategoryOptions(category);
    if (options.length === 0) {
      els.taskSubcategory.innerHTML = '<option value="">此主分類無子分類</option>';
      els.taskSubcategory.disabled = true;
      els.taskSubcategory.required = false;
      updateFormLockState();
      return;
    }

    els.taskSubcategory.disabled = false;
    els.taskSubcategory.required = true;
    els.taskSubcategory.innerHTML =
      '<option value="" disabled selected>請選擇子分類</option>' +
      options
        .map(function (item) {
          return '<option value="' + item + '">' + item + "</option>";
        })
        .join("");

    els.taskSubcategory.value = options.indexOf(current) !== -1 ? current : "";
    updateFormLockState();
  }

  function setupQueryCategoryOptions() {
    if (!els.queryCategory) {
      return;
    }
    const current = normalizeQueryCategory(els.queryCategory.value || state.queryCategory);
    const optionsHtml = ['<option value="all">全部主分類</option>']
      .concat(
        CATEGORIES.map(function (category) {
          return '<option value="' + category + '">' + category + "</option>";
        }),
      )
      .join("");
    els.queryCategory.innerHTML = optionsHtml;
    els.queryCategory.value = current;
    state.queryCategory = current;
  }

  function isCategorySelectionReady() {
    const category = String(els.taskCategory.value || "").trim();
    if (!isValidCategory(category)) {
      return false;
    }
    if (!hasSubcategoryOptions(category)) {
      return true;
    }
    const subcategory = String(els.taskSubcategory.value || "").trim();
    return Boolean(normalizeSubcategory(category, subcategory));
  }

  function updateFormLockState() {
    const locked = !isCategorySelectionReady();
    const category = String(els.taskCategory.value || "").trim();
    const subcategory = String(els.taskSubcategory.value || "").trim();

    [els.taskTitle, els.taskOwner, els.taskStartDate, els.taskEndDate, els.taskStartAt, els.taskEndAt, els.taskPinned, els.taskDescription].forEach(function (el) {
      el.disabled = locked;
    });
    els.allDayBtn.disabled = locked;
    els.addTaskBtn.disabled = locked;

    if (!isValidCategory(category)) {
      els.categoryTip.textContent = "請先選主分類，再填寫其餘欄位。";
      return;
    }

    if (hasSubcategoryOptions(category) && !normalizeSubcategory(category, subcategory)) {
      els.categoryTip.textContent = "此主分類需要子分類，請先選擇子分類。";
      return;
    }

    if (hasSubcategoryOptions(category)) {
      els.categoryTip.textContent = "目前分類：" + category + " / " + subcategory;
      return;
    }

    els.categoryTip.textContent = "目前分類：" + category;
  }

  function loadTasks() {
    try {
      const raw = localStorage.getItem(getScopedStorageKey(STORAGE_KEY));
      if (!raw) {
        return [];
      }
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        return [];
      }
      return parsed.map(normalizeTask).filter(Boolean).sort(sortByDueTime);
    } catch (error) {
      console.error("loadTasks error", error);
      return [];
    }
  }

  function saveTasks() {
    localStorage.setItem(getScopedStorageKey(STORAGE_KEY), JSON.stringify(state.tasks));
    scheduleCloudPush();
  }

  function loadDeletedTaskIds() {
    try {
      const raw = localStorage.getItem(getScopedStorageKey(DELETED_TASK_IDS_KEY));
      if (!raw) {
        return {};
      }
      const parsed = JSON.parse(raw);
      return normalizeDeletedTaskIds(parsed);
    } catch (error) {
      console.error("loadDeletedTaskIds error", error);
      return {};
    }
  }

  function saveDeletedTaskIds() {
    localStorage.setItem(getScopedStorageKey(DELETED_TASK_IDS_KEY), JSON.stringify(state.deletedTaskIds || {}));
    scheduleCloudPush();
  }

  function normalizeTask(input) {
    if (!input || typeof input !== "object") {
      return null;
    }
    const rawStartAt = input.startAt || input.dueAt || null;
    const rawEndAt = input.endAt || null;
    const hasStartAt = Boolean(rawStartAt);
    const hasEndAt = Boolean(rawEndAt);
    const startMs = hasStartAt ? new Date(rawStartAt).getTime() : Number.NaN;
    const endMs = hasEndAt ? new Date(rawEndAt).getTime() : Number.NaN;
    if (hasStartAt && Number.isNaN(startMs)) {
      return null;
    }
    if (hasEndAt && Number.isNaN(endMs)) {
      return null;
    }
    if (hasStartAt && hasEndAt && endMs < startMs) {
      return null;
    }
    const status = input.status === "done" ? "done" : "pending";
    const allDay = Boolean(input.allDay);
    const remindedAtMs = input.remindedAt ? new Date(input.remindedAt).getTime() : Number.NaN;
    const createdAtMs = input.createdAt ? new Date(input.createdAt).getTime() : Number.NaN;
    const updatedAtMs = input.updatedAt ? new Date(input.updatedAt).getTime() : Number.NaN;
    const createdAtIso = Number.isNaN(createdAtMs) ? new Date().toISOString() : new Date(createdAtMs).toISOString();
    const category = normalizeCategory(input.category);
    return {
      id: typeof input.id === "string" ? input.id : buildId(),
      category: category,
      subcategory: normalizeSubcategory(category, input.subcategory),
      title: String(input.title || "").trim(),
      owner: String(input.owner || "").trim(),
      completedBy: String(input.completedBy || "").trim(),
      description: String(input.description || "").trim(),
      startAt: hasStartAt ? new Date(startMs).toISOString() : null,
      endAt: hasEndAt ? new Date(endMs).toISOString() : null,
      dueAt: hasStartAt ? new Date(startMs).toISOString() : null,
      allDay: allDay,
      status: status,
      pinned: Boolean(input.pinned),
      remindedAt: Number.isNaN(remindedAtMs) ? null : new Date(remindedAtMs).toISOString(),
      createdAt: createdAtIso,
      updatedAt: Number.isNaN(updatedAtMs) ? createdAtIso : new Date(updatedAtMs).toISOString(),
    };
  }

  function handleAddTask(event) {
    event.preventDefault();

    if (!isCategorySelectionReady()) {
      showToast("請先完成分類選擇。");
      return;
    }

    const category = String(els.taskCategory.value || "").trim();
    const subcategory = String(els.taskSubcategory.value || "").trim();
    const title = els.taskTitle.value.trim();
    const owner = els.taskOwner.value.trim();
    const description = els.taskDescription.value.trim();
    const startDateInput = String(els.taskStartDate ? els.taskStartDate.value || "" : "").trim();
    const endDateInput = String(els.taskEndDate ? els.taskEndDate.value || "" : "").trim();
    const startAtInput = String(els.taskStartAt.value || "").trim();
    const endAtInput = String(els.taskEndAt.value || "").trim();
    const pinned = Boolean(els.taskPinned.checked);
    const allDay = Boolean(state.formAllDay);
    const normalizedSubcategory = normalizeSubcategory(category, subcategory);
    const range = parseTaskTimeRange(startDateInput, startAtInput, endDateInput, endAtInput, allDay);
    if (!range.ok) {
      showToast(range.message);
      return;
    }

    if (!isValidCategory(category) || !title || !owner) {
      showToast("請填寫完整資料後再新增。");
      return;
    }
    if (hasSubcategoryOptions(category) && !normalizedSubcategory) {
      showToast("請先選擇子分類。");
      return;
    }

    if (state.editingTaskId) {
      const task = state.tasks.find(function (item) {
        return item.id === state.editingTaskId;
      });
      if (!task) {
        showToast("找不到要修改的待辦。");
        resetTaskForm();
        return;
      }
      task.category = category;
      task.subcategory = normalizedSubcategory;
      task.title = title;
      task.owner = owner;
      task.description = description;
      task.startAt = range.startAtIso;
      task.endAt = range.endAtIso;
      task.dueAt = range.startAtIso;
      task.allDay = allDay;
      task.pinned = pinned;
      if (task.status === "pending") {
        task.remindedAt = null;
      }
      touchTask(task);
      state.tasks.sort(sortByDueTime);
      saveTasks();
      renderAll();
      resetTaskForm();
      showToast("待辦已修改。");
      return;
    }

    const nowIso = new Date().toISOString();
    state.tasks.push({
      id: buildId(),
      category: category,
      subcategory: normalizedSubcategory,
      title: title,
      owner: owner,
      completedBy: "",
      description: description,
      startAt: range.startAtIso,
      endAt: range.endAtIso,
      dueAt: range.startAtIso,
      allDay: allDay,
      status: "pending",
      pinned: pinned,
      remindedAt: null,
      createdAt: nowIso,
      updatedAt: nowIso,
    });

    state.tasks.sort(sortByDueTime);
    saveTasks();
    renderAll();
    resetTaskForm();
    showToast("待辦已新增。");
  }

  function applyDateQuery() {
    state.queryDate = els.queryDate.value;
    state.queryStatus = normalizeQueryStatus(els.queryStatus ? els.queryStatus.value : "all");
    state.queryCategory = normalizeQueryCategory(els.queryCategory ? els.queryCategory.value : "all");
    state.queryKeyword = normalizeQueryKeyword(els.queryKeyword ? els.queryKeyword.value : "");
    if (els.queryStatus) {
      els.queryStatus.value = state.queryStatus;
    }
    if (els.queryCategory) {
      els.queryCategory.value = state.queryCategory;
    }
    if (els.queryKeyword) {
      els.queryKeyword.value = state.queryKeyword;
    }
    setPanelCollapsed("task-list-panel", false);
    renderAll();
  }

  function clearDateQuery() {
    state.queryDate = "";
    state.queryStatus = "all";
    state.queryCategory = "all";
    state.queryKeyword = "";
    els.queryDate.value = "";
    if (els.queryStatus) {
      els.queryStatus.value = "all";
    }
    if (els.queryCategory) {
      els.queryCategory.value = "all";
    }
    if (els.queryKeyword) {
      els.queryKeyword.value = "";
    }
    renderAll();
  }

  function handleTodayCategoryChange() {
    state.todayCategory = String(els.todayCategoryFilter.value || "all");
    renderTodayTasks();
  }

  function handleTodayListAction(event) {
    const target = event.target.closest("button[data-action]");
    if (!target) {
      return;
    }
    const action = String(target.dataset.action || "").trim();
    const taskId = String(target.dataset.id || "").trim();
    runTaskAction(action, taskId);
  }

  function handleTableAction(event) {
    const target = event.target.closest("button[data-action]");
    if (!target) {
      return;
    }
    const action = String(target.dataset.action || "").trim();
    const taskId = String(target.dataset.id || "").trim();
    runTaskAction(action, taskId);
  }

  function runTaskAction(rawAction, taskId) {
    const action = rawAction === "cancel" ? "delete" : rawAction;
    if (!action || !taskId) {
      return;
    }
    const task = state.tasks.find(function (item) {
      return item.id === taskId;
    });
    if (!task) {
      return;
    }

    if (action === "toggle") {
      if (task.status === "pending") {
        const input = window.prompt("請輸入完成人：", task.completedBy || "");
        if (input === null) {
          return;
        }
        const completedBy = input.trim();
        if (!completedBy) {
          showToast("請填寫完成人後才能完成。");
          return;
        }
        task.status = "done";
        task.completedBy = completedBy;
        task.pinned = false;
      } else {
        task.status = "pending";
        task.remindedAt = null;
        task.completedBy = "";
      }
      touchTask(task);
      saveTasks();
      renderAll();
      showToast(task.status === "done" ? "已標記完成。" : "已恢復為待辦。");
      return;
    }

    if (action === "edit") {
      startEditing(task);
      return;
    }

    if (action === "copy") {
      copyTaskText(task);
      return;
    }

    if (action === "pin") {
      task.pinned = !task.pinned;
      touchTask(task);
      saveTasks();
      renderAll();
      showToast(task.pinned ? "已設為置頂。" : "已取消置頂。");
      return;
    }

    if (action === "delete") {
      const ok = window.confirm("確定要取消並刪除此待辦嗎？");
      if (!ok) {
        return;
      }
      state.tasks = state.tasks.filter(function (item) {
        return item.id !== taskId;
      });
      if (state.editingTaskId === taskId) {
        resetTaskForm();
      }
      state.deletedTaskIds[taskId] = new Date().toISOString();
      saveTasks();
      saveDeletedTaskIds();
      renderAll();
      showToast("待辦已刪除。");
    }
  }

  function startEditing(task) {
    state.editingTaskId = task.id;
    els.taskCategory.value = task.category;
    updateSubcategoryOptions();
    if (hasSubcategoryOptions(task.category)) {
      els.taskSubcategory.value = task.subcategory || "";
    }
    els.taskTitle.value = task.title;
    els.taskOwner.value = task.owner;
    setAllDayMode(Boolean(task.allDay));
    const startAt = getTaskStartAt(task);
    const endAt = getTaskEndAt(task);
    if (els.taskStartDate) {
      const baseDate = startAt || endAt;
      els.taskStartDate.value = baseDate ? toDateKey(new Date(baseDate)) : toDateKey(new Date());
    }
    if (els.taskEndDate) {
      els.taskEndDate.value = endAt ? toDateKey(new Date(endAt)) : "";
    }
    if (startAt) {
      els.taskStartAt.value = task.allDay ? "" : formatTime(new Date(startAt), false);
    } else {
      els.taskStartAt.value = "";
    }
    if (endAt) {
      els.taskEndAt.value = task.allDay ? "" : formatTime(new Date(endAt), false);
    } else {
      els.taskEndAt.value = "";
    }
    els.taskPinned.checked = Boolean(task.pinned);
    els.taskDescription.value = task.description || "";
    els.addTaskBtn.textContent = "儲存修改";
    els.cancelEditBtn.classList.remove("hidden");
    updateFormLockState();
    scrollToTaskFormPanel();
    if (els.taskTitle && typeof els.taskTitle.focus === "function") {
      try {
        els.taskTitle.focus();
      } catch (error) {
        // ignore unsupported focus options in older browsers
      }
      if (typeof els.taskTitle.select === "function") {
        els.taskTitle.select();
      }
    }
    showToast("已載入待辦，可開始修改。");
  }

  function scrollToTaskFormPanel() {
    const panel = els.taskForm && typeof els.taskForm.closest === "function" ? els.taskForm.closest(".panel") : null;
    if (!panel) {
      return;
    }
    const prefersReducedMotion =
      typeof window.matchMedia === "function" && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
    const behavior = prefersReducedMotion ? "auto" : "smooth";
    const panelTop = panel.getBoundingClientRect().top + (window.pageYOffset || document.documentElement.scrollTop || 0);
    const targetTop = Math.max(0, Math.floor(panelTop - 18));

    try {
      window.scrollTo({ top: targetTop, behavior: behavior });
    } catch (error) {
      window.scrollTo(0, targetTop);
    }

    // Some desktop browsers ignore smooth scrolling in busy layouts; enforce once more.
    setTimeout(function () {
      const rect = panel.getBoundingClientRect();
      const outOfView = rect.top < 0 || rect.top > window.innerHeight * 0.35;
      if (!outOfView) {
        return;
      }
      const top = panel.getBoundingClientRect().top + (window.pageYOffset || document.documentElement.scrollTop || 0);
      const retryTop = Math.max(0, Math.floor(top - 18));
      try {
        window.scrollTo({ top: retryTop, behavior: "auto" });
      } catch (error) {
        window.scrollTo(0, retryTop);
      }
    }, 180);
  }

  function cancelEditing() {
    if (!state.editingTaskId) {
      return;
    }
    resetTaskForm();
    showToast("已取消修改。");
  }

  function handleClearForm() {
    resetTaskForm();
    showToast("表單已清空。");
  }

  function resetTaskForm() {
    state.editingTaskId = null;
    els.taskForm.reset();
    updateSubcategoryOptions();
    setDefaultDueTime();
    els.addTaskBtn.textContent = "新增待辦";
    els.cancelEditBtn.classList.add("hidden");
    updateFormLockState();
  }

  function renderAll() {
    renderQueryText();
    renderTaskTable();
    renderTodayTasks();
    renderUpcomingBoard();
  }

  function renderQueryText() {
    const list = getFilteredTasks();
    const hasKeyword = Boolean(normalizeQueryKeyword(state.queryKeyword));
    if (!state.queryDate && state.queryStatus === "all" && state.queryCategory === "all" && !hasKeyword) {
      els.queryResultText.textContent = "目前顯示全部待辦，共 " + list.length + " 筆。";
      return;
    }
    els.queryResultText.textContent = "目前顯示「" + buildQueryConditionText() + "」，共 " + list.length + " 筆。";
  }

  function renderTodayTasks() {
    if (!els.todayPinnedList || !els.todayTaskSummary || !els.todayTaskList) {
      return;
    }
    const todayKey = toDateKey(new Date());
    const selectedCategory = state.todayCategory || "all";
    const todayAllList = state.tasks
      .filter(function (task) {
        if (selectedCategory !== "all" && task.category !== selectedCategory) {
          return false;
        }
        const startAt = getTaskStartAt(task);
        return startAt && toDateKey(new Date(startAt)) === todayKey;
      })
      .sort(sortByDueTime);

    const pinnedList = state.tasks
      .filter(function (task) {
        return Boolean(task.pinned) && task.status === "pending";
      })
      .sort(sortByDueTime);

    const normalList = todayAllList.filter(function (task) {
      return task.status === "pending" && !task.pinned;
    });

    const categoryLabel = selectedCategory === "all" ? "全部主分類" : selectedCategory;
    els.todayPinnedList.innerHTML =
      pinnedList.length === 0
        ? '<li class="today-pinned-empty">目前沒有置頂事項。</li>'
        : pinnedList
            .map(function (task) {
              return renderTodayTaskItem(task, true);
            })
            .join("");

    if (todayAllList.length === 0) {
      els.todayTaskSummary.textContent = "今日共 0 筆事項（" + categoryLabel + "）。";
      els.todayTaskList.innerHTML = '<li class="today-empty">今日沒有排程事項。</li>';
      return;
    }

    const pendingCount = todayAllList.filter(function (task) {
      return task.status === "pending";
    }).length;
    const doneCount = todayAllList.length - pendingCount;
    els.todayTaskSummary.textContent =
      "今日共 " +
      todayAllList.length +
      " 筆，待處理 " +
      pendingCount +
      " 筆，已完成 " +
      doneCount +
      " 筆（" +
      categoryLabel +
      "）。";

    els.todayTaskList.innerHTML =
      normalList.length === 0
        ? '<li class="today-empty">今日一般事項為 0 筆。</li>'
        : normalList
            .map(function (task) {
              return renderTodayTaskItem(task, false);
            })
            .join("");
  }

  function renderTodayTaskItem(task, inPinnedZone) {
    const isDone = task.status === "done";
    const statusClass = isDone ? "status-done" : "status-pending";
    const statusText = isDone ? "已完成" : "待處理";
    const pinTag = task.pinned ? '<span class="today-pill">置頂</span>' : "";
    const pinBtnText = task.pinned ? "取消置頂" : "置頂";
    const doneBtnText = isDone ? "取消完成" : "完成";
    const countdownText = formatCountdown(getTaskStartAt(task), task.allDay, isDone);
    const completionText = isDone && task.completedBy ? " | 完成人：" + escapeHtml(task.completedBy) : "";
    const metaLineText =
      escapeHtml(task.category) +
      (task.subcategory ? " / " + escapeHtml(task.subcategory) : "") +
      " | 填寫人：" +
      escapeHtml(task.owner || "未填寫") +
      completionText +
      " | 倒數：" +
      escapeHtml(countdownText);
    const descriptionHtml = task.description
      ? '<p class="today-desc"><span class="today-desc-label">內容：</span>' +
        escapeHtml(task.description).replace(/\n/g, "<br>") +
        "</p>"
      : "";
    const dueMs = getTaskStartAt(task) ? new Date(getTaskStartAt(task)).getTime() : Number.NaN;
    const isOverdue = !isDone && !Number.isNaN(dueMs) && dueMs < Date.now();

    return (
      '<li class="today-item' +
      (isOverdue ? " overdue" : "") +
      (isDone ? " done" : "") +
      (inPinnedZone ? " pinned-zone-item" : "") +
      '">' +
      '<div class="today-item-top">' +
      '<span class="status-badge ' +
      statusClass +
      '">' +
      statusText +
      "</span>" +
      '<span class="today-category-pill">' +
      escapeHtml(task.category || "未分類") +
      "</span>" +
      pinTag +
      '<span class="today-time">' +
      formatTimeRange(task) +
      "</span>" +
      "</div>" +
      '<p class="today-title">' +
      escapeHtml(task.title) +
      "</p>" +
      '<p class="today-meta today-meta-line">' +
      metaLineText +
      "</p>" +
      descriptionHtml +
      '<div class="today-actions row-actions">' +
      '<button class="action-btn action-done" type="button" data-action="toggle" data-id="' +
      escapeHtml(task.id) +
      '">' +
      doneBtnText +
      "</button>" +
      '<button class="action-btn action-edit" type="button" data-action="edit" data-id="' +
      escapeHtml(task.id) +
      '">修改</button>' +
      '<button class="action-btn action-copy" type="button" data-action="copy" data-id="' +
      escapeHtml(task.id) +
      '">一鍵複製</button>' +
      '<button class="action-btn action-cancel" type="button" data-action="cancel" data-id="' +
      escapeHtml(task.id) +
      '">取消</button>' +
      '<button class="action-btn action-pin" type="button" data-action="pin" data-id="' +
      escapeHtml(task.id) +
      '">' +
      pinBtnText +
      "</button>" +
      "</div>" +
      "</li>"
    );
  }

  function renderTaskTable() {
    const list = getTaskListFilteredTasks();
    if (list.length === 0) {
      els.tableBody.innerHTML =
        '<tr class="empty-row"><td colspan="10">目前沒有符合條件的待辦事項。</td></tr>';
      return;
    }

    els.tableBody.innerHTML = list
      .map(function (task) {
        const isDone = task.status === "done";
        const statusClass = isDone ? "status-done" : "status-pending";
        const statusText = isDone ? "已完成" : "待處理";
        const pinnedText = task.pinned ? "置頂中" : "-";
        const pinBtnText = task.pinned ? "取消置頂" : "置頂";
        const completedByText = task.status === "done" && task.completedBy ? escapeHtml(task.completedBy) : "-";
        const subcategoryHtml = task.subcategory
          ? escapeHtml(task.subcategory)
          : '<span style="color:#7a9198;">-</span>';
        const descHtml = task.description
          ? escapeHtml(task.description).replace(/\n/g, "<br>")
          : '<span style="color:#7a9198;">-</span>';
        return (
          '<tr class="' +
          (isDone ? "task-row-done" : "") +
          '">' +
          '<td><span class="status-badge ' +
          statusClass +
          '">' +
          statusText +
          "</span></td>" +
          "<td>" +
          pinnedText +
          "</td>" +
          "<td>" +
          escapeHtml(task.category) +
          "</td>" +
          "<td>" +
          subcategoryHtml +
          "</td>" +
          "<td>" +
          escapeHtml(task.title) +
          "</td>" +
          "<td>" +
          escapeHtml(task.owner) +
          "</td>" +
          "<td>" +
          completedByText +
          "</td>" +
          "<td>" +
          formatDueDisplay(task) +
          "</td>" +
          "<td>" +
          descHtml +
          "</td>" +
          '<td><div class="row-actions">' +
          '<button class="action-btn action-done" type="button" data-action="toggle" data-id="' +
          task.id +
          '">' +
          (isDone ? "恢復" : "完成") +
          "</button>" +
          '<button class="action-btn action-edit" type="button" data-action="edit" data-id="' +
          task.id +
          '">修改</button>' +
          '<button class="action-btn action-copy" type="button" data-action="copy" data-id="' +
          task.id +
          '">一鍵複製</button>' +
          '<button class="action-btn action-pin" type="button" data-action="pin" data-id="' +
          task.id +
          '">' +
          pinBtnText +
          "</button>" +
          '<button class="action-btn action-delete" type="button" data-action="delete" data-id="' +
          task.id +
          '">刪除</button>' +
          "</div></td>" +
          "</tr>"
        );
      })
      .join("");
  }

  function renderUpcomingBoard() {
    const now = Date.now();
    const within30 = [];
    const within60 = [];

    state.tasks
      .filter(function (task) {
        return task.status === "pending" && Boolean(getTaskStartAt(task)) && !task.allDay;
      })
      .sort(sortByDueTime)
      .forEach(function (task) {
        const diffMs = new Date(getTaskStartAt(task)).getTime() - now;
        if (diffMs < 0) {
          return;
        }
        if (diffMs <= 30 * 60 * 1000) {
          within30.push(task);
          return;
        }
        if (diffMs <= 60 * 60 * 1000) {
          within60.push(task);
        }
      });

    const total = within30.length + within60.length;
    if (total === 0) {
      els.upcomingSummary.textContent = "目前沒有 1 小時內待辦。";
      els.upcomingTaskList.innerHTML = '<li class="upcoming-empty">30 分鐘與 1 小時提醒區間目前無待辦。</li>';
      return;
    }

    els.upcomingSummary.textContent =
      "30 分鐘內 " + within30.length + " 筆，1 小時內 " + within60.length + " 筆。";

    const items = [];
    within30.forEach(function (task) {
      items.push(renderUpcomingItem(task, "30 分鐘內"));
    });
    within60.forEach(function (task) {
      items.push(renderUpcomingItem(task, "1 小時內"));
    });
    els.upcomingTaskList.innerHTML = items.join("");
  }

  function renderUpcomingItem(task, windowTag) {
    const detailText = task.description
      ? escapeHtml(task.description).replace(/\n/g, "<br>")
      : "（無）";
    const countdownText = formatCountdown(getTaskStartAt(task), task.allDay, false);
    const metaLineText =
      escapeHtml(task.category) +
      (task.subcategory ? " / " + escapeHtml(task.subcategory) : "") +
      " | 填寫人：" +
      escapeHtml(task.owner || "未填寫") +
      " | 倒數：" +
      escapeHtml(countdownText);

    return (
      '<li class="upcoming-item">' +
      '<div class="upcoming-item-top">' +
      '<span class="upcoming-tag">' +
      windowTag +
      "</span>" +
      '<span class="upcoming-time">' +
      formatTimeRange(task) +
      "</span>" +
      "</div>" +
      '<p class="upcoming-title">' +
      escapeHtml(task.title) +
      "</p>" +
      '<p class="upcoming-meta upcoming-meta-line">' +
      metaLineText +
      "</p>" +
      '<p class="upcoming-detail"><span class="upcoming-detail-label">時間：</span>' +
      escapeHtml(formatDueDisplay(task)) +
      "</p>" +
      '<p class="upcoming-detail"><span class="upcoming-detail-label">內容：</span>' +
      detailText +
      "</p>" +
      "</li>"
    );
  }

  function getFilteredTasks() {
    return getTasksByDateAndStatus(state.queryDate, state.queryStatus, state.queryCategory, state.queryKeyword);
  }

  function getTaskListFilteredTasks() {
    const list = getFilteredTasks();
    const quickFilter = normalizeTaskListFilter(state.taskListFilter);
    if (quickFilter === "pending") {
      return list.filter(function (task) {
        return task.status === "pending";
      });
    }
    if (quickFilter === "done") {
      return list.filter(function (task) {
        return task.status === "done";
      });
    }
    if (quickFilter === "pinned") {
      return list.filter(function (task) {
        return Boolean(task.pinned);
      });
    }
    return list;
  }

  function startReminderLoop() {
    if (state.reminderTimer) {
      clearInterval(state.reminderTimer);
    }
    checkReminders();
    state.reminderTimer = setInterval(checkReminders, REMINDER_CHECK_MS);
  }

  function startCountdownLoop() {
    if (state.countdownTimer) {
      clearInterval(state.countdownTimer);
    }
    state.countdownTimer = setInterval(function () {
      if (document.hidden) {
        return;
      }
      renderTodayTasks();
      renderUpcomingBoard();
    }, COUNTDOWN_REFRESH_MS);
  }

  function checkReminders() {
    renderTodayTasks();
    renderUpcomingBoard();

    const pending = state.tasks
      .filter(function (task) {
        return task.status === "pending" && Boolean(getTaskStartAt(task)) && !task.allDay;
      })
      .sort(sortByDueTime);

    if (pending.length === 0) {
      return;
    }

    const now = Date.now();
    const dueTask = pending.find(function (task) {
      const dueMs = new Date(getTaskStartAt(task)).getTime();
      return !task.remindedAt && dueMs <= now;
    });

    if (!dueTask) {
      return;
    }

    sendReminder(dueTask);
    dueTask.remindedAt = new Date().toISOString();
    saveTasks();
  }

  function sendReminder(task) {
    const text = "待辦「" + task.title + "」已到時間。";
    showToast(text);

    if (!els.notificationToggle.checked) {
      return;
    }
    if (!("Notification" in window)) {
      return;
    }
    if (Notification.permission === "granted") {
      const categoryText = task.subcategory ? task.category + "/" + task.subcategory : task.category;
      new Notification("工作交接提醒", {
        body:
          task.title +
          "（分類：" +
          categoryText +
          "，填寫人：" +
          (task.owner || "未填寫") +
          "）",
      });
    }
  }

  async function requestNotificationPermission() {
    if (!("Notification" in window)) {
      showToast("目前瀏覽器不支援通知功能。");
      return;
    }
    try {
      const result = await Notification.requestPermission();
      if (result === "granted") {
        showToast("已授權瀏覽器通知。");
      } else if (result === "denied") {
        showToast("通知已被封鎖，可到瀏覽器設定開啟。");
      } else {
        showToast("通知授權尚未開啟。");
      }
    } catch (error) {
      console.error("requestPermission error", error);
      showToast("通知授權失敗。");
    }
  }

  function collectExportContext() {
    const exportDateInput = els.exportDate ? String(els.exportDate.value || "").trim() : "";
    const effectiveDate = normalizeDateInputFromAny(exportDateInput || state.queryDate || "");
    const exportStatus = normalizeQueryStatus(els.exportStatus ? els.exportStatus.value : "all");
    let list = getTasksByDateAndStatus(effectiveDate, exportStatus);
    const conditionText = buildConditionText(effectiveDate, exportStatus);
    const hasDateFilter = Boolean(effectiveDate);
    const includeDatePrefix = !hasDateFilter;

    if (!hasDateFilter) {
      const todayKey = toDateKey(new Date());
      list = list.filter(function (task) {
        return isTaskNotPastByDateKey(task, todayKey);
      });
    }
    return {
      effectiveDate: effectiveDate,
      exportStatus: exportStatus,
      list: list.slice().sort(sortForExportCategoryThenTime),
      conditionText: conditionText,
      includeDatePrefix: includeDatePrefix,
    };
  }

  async function handleExportWord() {
    const context = collectExportContext();

    if (window.docx && window.saveAs) {
      try {
        await exportDocx(
          context.list,
          context.conditionText,
          context.effectiveDate,
          context.exportStatus,
          context.includeDatePrefix,
        );
        showToast("Word file exported.");
        return;
      } catch (error) {
        console.error("exportDocx error", error);
      }
    }

    exportLegacyDoc(
      context.list,
      context.conditionText,
      context.effectiveDate,
      context.exportStatus,
      context.includeDatePrefix,
    );
    showToast("Exported via compatibility Word mode.");
  }

  function handleExportExcel() {
    const context = collectExportContext();
    try {
      exportExcelFile(context.list, context.effectiveDate, context.includeDatePrefix);
      showToast("Excel file exported.");
    } catch (error) {
      console.error("exportExcel error", error);
      showToast("Excel 匯出失敗，請稍後重試。");
    }
  }

  function exportExcelFile(tasks, exportDate, includeDatePrefix) {
    const rows = buildExcelRows(tasks);
    const headers = ["事項名稱", "主分類", "填寫人", "完成人", "交接說明"];
    if (window.XLSX && window.XLSX.utils) {
      const built = buildExcelAoaAndMerges(tasks, exportDate, includeDatePrefix, headers);
      const aoa = built.aoa;
      const merges = built.merges;
      const worksheet = window.XLSX.utils.aoa_to_sheet(aoa);
      worksheet["!cols"] = [{ wch: 16 }, { wch: 9 }, { wch: 9 }, { wch: 9 }, { wch: 22 }];
      worksheet["!rows"] = buildExcelRowHeights(aoa);
      worksheet["!merges"] = merges;
      worksheet["!margins"] = {
        left: 0.2,
        right: 0.2,
        top: 0.25,
        bottom: 0.25,
        header: 0.15,
        footer: 0.15,
      };
      worksheet["!pageSetup"] = {
        paperSize: 9,
        orientation: "portrait",
        fitToWidth: 1,
        fitToHeight: 1,
      };
      const workbook = window.XLSX.utils.book_new();
      const sheetName = sanitizeExcelSheetName(includeDatePrefix ? "未來待辦事項" : "每日報告");
      window.XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
      window.XLSX.writeFile(workbook, buildExportFileName("xlsx", exportDate), { compression: true });
      return;
    }
    exportExcelCsvFallback(rows, headers, exportDate);
  }

  function buildExcelRows(tasks) {
    const list = Array.isArray(tasks) ? tasks : [];
    return list.map(function (task) {
      const completedBy = task && task.status === "done" && task.completedBy ? String(task.completedBy).trim() : "-";
      const description = formatExcelDescription(task && task.description ? task.description : "");
      const title = task && task.title ? String(task.title).trim() : "-";
      const timeText = formatExportTaskTime(task);
      const timePrefix = timeText && timeText !== "-" ? timeText + " " : "";
      return {
        事項名稱: timePrefix + title,
        主分類: task && task.category ? String(task.category).trim() : "-",
        填寫人: task && task.owner ? String(task.owner).trim() : "-",
        完成人: completedBy || "-",
        交接說明: description || "-",
      };
    });
  }

  function buildExcelAoaAndMerges(tasks, exportDate, includeDatePrefix, headers) {
    const list = Array.isArray(tasks) ? tasks.slice().sort(sortForExportCategoryThenTime) : [];
    const cols = Array.isArray(headers) && headers.length > 0 ? headers.length : 5;
    const aoa = [];
    const merges = [];

    function pushMergedRow(text) {
      const rowIndex = aoa.length;
      const row = [String(text || "")];
      while (row.length < cols) {
        row.push("");
      }
      aoa.push(row);
      merges.push({
        s: { r: rowIndex, c: 0 },
        e: { r: rowIndex, c: cols - 1 },
      });
    }

    function pushHeaderRow() {
      aoa.push(headers.slice());
    }

    function pushTaskRows(taskList) {
      const rows = buildExcelRows(taskList);
      if (rows.length === 0) {
        const emptyRow = [];
        for (let i = 0; i < cols; i += 1) {
          emptyRow.push("-");
        }
        aoa.push(emptyRow);
        return;
      }
      rows.forEach(function (row) {
        aoa.push(
          headers.map(function (key) {
            return row[key];
          }),
        );
      });
    }

    if (!includeDatePrefix) {
      pushMergedRow(buildArrDepOccLine(exportDate));
    }
    pushMergedRow(includeDatePrefix ? "Future To-Do / 未來待辦事項" : "Daily Briefing / 每日報告");
    aoa.push([]);

    if (!includeDatePrefix) {
      pushHeaderRow();
      pushTaskRows(list);
      return { aoa: aoa, merges: merges };
    }

    const grouped = groupExportTasksByDate(list);
    if (grouped.length === 0) {
      pushHeaderRow();
      pushTaskRows([]);
      return { aoa: aoa, merges: merges };
    }

    grouped.forEach(function (group, index) {
      pushMergedRow("日期：" + group.label);
      pushHeaderRow();
      pushTaskRows(group.tasks);
      if (index < grouped.length - 1) {
        aoa.push([]);
      }
    });
    return { aoa: aoa, merges: merges };
  }

  function sanitizeExcelSheetName(name) {
    const raw = String(name || "").trim() || "Sheet1";
    const sanitized = raw.replace(/[\\\/\?\*\[\]:]/g, "_");
    return sanitized.slice(0, 31) || "Sheet1";
  }

  function buildExcelRowHeights(aoa) {
    const rows = [];
    const list = Array.isArray(aoa) ? aoa : [];
    for (let i = 0; i < list.length; i += 1) {
      const row = Array.isArray(list[i]) ? list[i] : [];
      const joined = row
        .map(function (cell) {
          return String(cell == null ? "" : cell);
        })
        .join(" ");
      if (!joined.trim()) {
        rows.push({ hpt: 6 });
        continue;
      }
      if (row.some(function (cell) { return String(cell == null ? "" : cell).indexOf("\n") >= 0; })) {
        rows.push({ hpt: 28 });
        continue;
      }
      rows.push({ hpt: 18 });
    }
    return rows;
  }

  function formatExcelDescription(raw) {
    const text = String(raw || "")
      .replace(/\r?\n/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    if (!text) {
      return "-";
    }
    const perLine = 22;
    if (text.length <= perLine) {
      return text;
    }
    if (text.length <= perLine * 2) {
      return text.slice(0, perLine) + "\n" + text.slice(perLine);
    }
    return text.slice(0, perLine) + "\n" + text.slice(perLine, perLine * 2 - 1) + "…";
  }

  function exportExcelCsvFallback(rows, headers, exportDate) {
    const lines = [headers.join(",")].concat(
      rows.map(function (row) {
        return headers
          .map(function (key) {
            return csvEscape(row[key]);
          })
          .join(",");
      }),
    );
    const blob = new Blob(["\ufeff" + lines.join("\r\n")], { type: "text/csv;charset=utf-8" });
    const filename = buildExportFileName("csv", exportDate);
    if (window.saveAs) {
      window.saveAs(blob, filename);
      return;
    }
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
  }

  function csvEscape(value) {
    const text = String(value == null ? "" : value);
    if (!/[",\r\n]/.test(text)) {
      return text;
    }
    return '"' + text.replace(/"/g, '""') + '"';
  }

  function applyImportedBackup(imported, silent) {
    if (!imported || typeof imported !== "object") {
      return false;
    }
    const normalizedImportedDeleted = normalizeDeletedTaskIds(imported.deletedTaskIds);
    const tasksChanged = JSON.stringify(state.tasks) !== JSON.stringify(imported.tasks);
    const overviewChanged = JSON.stringify(state.todayOverview) !== JSON.stringify(imported.todayOverview);
    const deletedChanged = JSON.stringify(normalizeDeletedTaskIds(state.deletedTaskIds)) !== JSON.stringify(normalizedImportedDeleted);
    if (!tasksChanged && !overviewChanged && !deletedChanged) {
      return false;
    }

    state.tasks = imported.tasks;
    state.todayOverview = imported.todayOverview;
    state.deletedTaskIds = normalizedImportedDeleted;
    if (state.editingTaskId && !silent) {
      resetTaskForm();
    }
    saveTasks();
    saveDeletedTaskIds();
    saveTodayOverview();
    renderTodayOverviewBar();
    renderAll();
    if (!silent) {
      showToast("已讀取資料，共 " + state.tasks.length + " 筆。");
    }
    return true;
  }

  function parseBackupPayload(payload) {
    let rawTasks = [];
    let rawTodayOverview = {};
    let rawDeletedTaskIds = {};
    let rawServerId = "";

    if (Array.isArray(payload)) {
      rawTasks = payload;
    } else if (payload && typeof payload === "object") {
      rawTasks = Array.isArray(payload.tasks) ? payload.tasks : [];
      rawTodayOverview =
        payload.todayOverview && typeof payload.todayOverview === "object" ? payload.todayOverview : {};
      rawDeletedTaskIds =
        payload.deletedTaskIds && typeof payload.deletedTaskIds === "object" ? payload.deletedTaskIds : {};
      rawServerId = typeof payload.serverId === "string" ? payload.serverId.trim() : "";
    } else {
      return null;
    }

    const filtered = filterTasksByDeletedMap(rawTasks.map(normalizeTask).filter(Boolean), rawDeletedTaskIds);
    return {
      serverId: rawServerId || DEFAULT_SERVER_ID,
      tasks: filtered.tasks,
      todayOverview: normalizeTodayOverviewRecord(rawTodayOverview, true),
      deletedTaskIds: filtered.deletedTaskIds,
    };
  }

  async function exportDocx(tasks, conditionText, exportDate, exportStatus, includeDatePrefix) {
    const docxLib = window.docx;
    const Document = docxLib.Document;
    const Packer = docxLib.Packer;
    const Paragraph = docxLib.Paragraph;
    const TextRun = docxLib.TextRun;
    const Table = docxLib.Table;
    const TableCell = docxLib.TableCell;
    const TableRow = docxLib.TableRow;
    const WidthType = docxLib.WidthType;
    const AlignmentType = docxLib.AlignmentType;
    const BorderStyle = docxLib.BorderStyle;

    const dailyList = (Array.isArray(tasks) ? tasks.slice() : []).sort(sortForExportCategoryThenTime);
    const showFutureMode = Boolean(includeDatePrefix);
    const sectionTitle = showFutureMode ? "Future To-Do\n未來待辦事項" : "Daily Briefing\n每日報告";
    const tableTools = {
      Paragraph: Paragraph,
      TableCell: TableCell,
      TableRow: TableRow,
      WidthType: WidthType,
      AlignmentType: AlignmentType,
      BorderStyle: BorderStyle,
      TextRun: TextRun,
      includeDatePrefix: showFutureMode,
    };
    function createExportTable(rows) {
      return new Table({
        width: {
          size: 100,
          type: WidthType.PERCENTAGE,
        },
        alignment: AlignmentType.CENTER,
        columnWidths: [1900, 9000],
        borders: createExportBorders(BorderStyle),
        rows: rows,
      });
    }
    const children = [];
    if (!showFutureMode) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [createExportTextRun(TextRun, buildArrDepOccLine(exportDate), { bold: true })],
        }),
      );
      children.push(new Paragraph({ text: "" }));
    }

    if (showFutureMode) {
      const grouped = groupExportTasksByDate(dailyList);
      if (grouped.length === 0) {
        children.push(createExportTable([createExportSectionRow(sectionTitle, [], tableTools)]));
      } else {
        grouped.forEach(function (group, index) {
          children.push(
            new Paragraph({
              alignment: AlignmentType.LEFT,
              children: [createExportTextRun(TextRun, group.label, { bold: true })],
            }),
          );
          children.push(createExportTable([createExportSectionRow(sectionTitle, group.tasks, tableTools)]));
          if (index < grouped.length - 1) {
            children.push(new Paragraph({ text: "" }));
          }
        });
      }
    } else {
      children.push(createExportTable([createExportSectionRow(sectionTitle, dailyList, tableTools)]));
    }

    const document = new Document({
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 720,
                right: 720,
                bottom: 720,
                left: 720,
              },
            },
          },
          children: children,
        },
      ],
    });

    const blob = await Packer.toBlob(document);
    window.saveAs(blob, buildExportFileName("docx", exportDate));
  }

  function createExportSectionRow(title, tasks, tools) {
    const titleLines = String(title || "")
      .split("\n")
      .map(function (line) {
        return line.trim();
      })
      .filter(Boolean);

    return new tools.TableRow({
      children: [
        new tools.TableCell({
          width: {
            size: 1900,
            type: tools.WidthType.DXA,
          },
          borders: createExportBorders(tools.BorderStyle),
          children: createExportParagraphs(titleLines, tools.Paragraph, {
            alignment: tools.AlignmentType.CENTER,
            TextRun: tools.TextRun,
            bold: true,
          }),
        }),
        new tools.TableCell({
          width: {
            size: 9000,
            type: tools.WidthType.DXA,
          },
          borders: createExportBorders(tools.BorderStyle),
          children: createExportTaskParagraphs(tasks, tools),
        }),
      ],
    });
  }

  function createExportMergedRow(lines, tools) {
    return new tools.TableRow({
      children: [
        new tools.TableCell({
          columnSpan: 2,
          width: {
            size: 10900,
            type: tools.WidthType.DXA,
          },
          borders: createExportBorders(tools.BorderStyle),
          children: createExportParagraphs(lines, tools.Paragraph, {
            alignment: tools.center ? tools.AlignmentType.CENTER : tools.AlignmentType.LEFT,
            TextRun: tools.TextRun,
          }),
        }),
      ],
    });
  }

  function createExportParagraphs(lines, Paragraph, options) {
    const list = Array.isArray(lines) ? lines : [];
    const align = options && options.alignment ? options.alignment : undefined;
    const TextRun = options && options.TextRun ? options.TextRun : null;
    const bold = Boolean(options && options.bold);
    if (list.length === 0) {
      if (TextRun) {
        return [
          new Paragraph({
            alignment: align,
            children: [createExportTextRun(TextRun, "（無資料）", { bold: bold })],
          }),
        ];
      }
      return [new Paragraph({ text: "（無資料）", alignment: align })];
    }
    return list.map(function (line) {
      const normalized = normalizeExportLineItem(line);
      if (TextRun) {
        return new Paragraph({
          alignment: align,
          children: [
            createExportTextRun(TextRun, normalized.text, {
              color: normalized.color,
              bold: normalized.bold || bold,
            }),
          ],
        });
      }
      return new Paragraph({
        text: normalized.text,
        alignment: align,
      });
    });
  }

  function createExportBorders(BorderStyle) {
    return {
      top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    };
  }

  function normalizeExportLineItem(line) {
    if (line && typeof line === "object" && !Array.isArray(line)) {
      return {
        text: String(line.text || "").trim(),
        color: String(line.color || "").trim(),
        bold: Boolean(line.bold),
      };
    }
    return {
      text: String(line || "").trim(),
      color: "",
      bold: false,
    };
  }

  function createExportTextRun(TextRun, text, options) {
    const style = options && typeof options === "object" ? options : {};
    const color = String(style.color || "").trim().replace(/^#/, "") || EXPORT_DEFAULT_COLOR;
    return new TextRun({
      text: String(text || ""),
      bold: Boolean(style.bold),
      color: color,
      font: {
        ascii: EXPORT_FONT_LATIN,
        hAnsi: EXPORT_FONT_LATIN,
        eastAsia: EXPORT_FONT_EAST_ASIA,
        cs: EXPORT_FONT_LATIN,
      },
      size: 24,
    });
  }

  function createExportTaskParagraphs(tasks, tools) {
    const list = Array.isArray(tasks) ? tasks : [];
    if (list.length === 0) {
      return createExportParagraphs([], tools.Paragraph, {
        alignment: tools.AlignmentType.CENTER,
        TextRun: tools.TextRun,
      });
    }
    return list.map(function (task) {
      return new tools.Paragraph({
        alignment: tools.AlignmentType.CENTER,
        children: [
          createExportTextRun(tools.TextRun, formatExportTaskLine(task, { includeDatePrefix: tools.includeDatePrefix }), {
            color: getExportCategoryColor(task && task.category),
          }),
        ],
      });
    });
  }

  function getExportStatusPool(exportStatus) {
    const normalizedStatus = normalizeQueryStatus(exportStatus);
    return state.tasks
      .slice()
      .filter(function (task) {
        if (normalizedStatus === "all") {
          return true;
        }
        return task.status === normalizedStatus;
      })
      .sort(sortForTaskTable);
  }

  function buildExportSections(currentList, statusPool, exportDate) {
    const targetDate = exportDate || toDateKey(new Date());
    const endOfTargetDay = new Date(targetDate + "T23:59:59").getTime();

    const attention = statusPool.filter(function (task) {
      return task.category === CATEGORIES[9];
    });

    const daily = (Array.isArray(currentList) ? currentList.slice() : []).sort(sortForTaskTable);

    const banquets = statusPool.filter(function (task) {
      return task.category === CATEGORIES[4] || task.category === CATEGORIES[5] || task.category === CATEGORIES[7];
    });

    const future = statusPool.filter(function (task) {
      const startAt = getTaskStartAt(task);
      if (!startAt) {
        return false;
      }
      return new Date(startAt).getTime() > endOfTargetDay;
    });

    const transfer = statusPool.filter(function (task) {
      return task.category === CATEGORIES[1];
    });

    const used = new Set();
    [attention, daily, banquets, future, transfer].forEach(function (list) {
      list.forEach(function (task) {
        if (task && task.id) {
          used.add(task.id);
        }
      });
    });

    const fyi = statusPool.filter(function (task) {
      return !used.has(task.id);
    });

    return {
      attention: attention,
      daily: daily,
      banquets: banquets,
      future: future,
      transfer: transfer,
      fyi: fyi,
    };
  }

  function buildExportTaskLines(tasks, options) {
    if (!Array.isArray(tasks) || tasks.length === 0) {
      return [];
    }
    return tasks.map(function (task) {
      return formatExportTaskLine(task, options);
    });
  }

  function getExportCategoryColor(category) {
    const key = String(category || "").trim();
    if (!key) {
      return EXPORT_DEFAULT_COLOR;
    }
    return EXPORT_CATEGORY_COLORS[key] || EXPORT_DEFAULT_COLOR;
  }

  function formatExportTaskDate(task) {
    const startAt = getTaskStartAt(task);
    if (!startAt) {
      return "--/--";
    }
    const dt = new Date(startAt);
    if (Number.isNaN(dt.getTime())) {
      return "--/--";
    }
    const mm = String(dt.getMonth() + 1).padStart(2, "0");
    const dd = String(dt.getDate()).padStart(2, "0");
    return mm + "/" + dd;
  }

  function getExportTaskDateKey(task) {
    const startAt = getTaskStartAt(task);
    if (!startAt) {
      return "";
    }
    const dt = new Date(startAt);
    if (Number.isNaN(dt.getTime())) {
      return "";
    }
    return toDateKey(dt);
  }

  function groupExportTasksByDate(tasks) {
    const groupedMap = new Map();
    (Array.isArray(tasks) ? tasks : []).forEach(function (task) {
      const key = getExportTaskDateKey(task);
      if (!groupedMap.has(key)) {
        groupedMap.set(key, []);
      }
      groupedMap.get(key).push(task);
    });
    const sortedKeys = Array.from(groupedMap.keys()).sort(function (a, b) {
      if (!a && !b) {
        return 0;
      }
      if (!a) {
        return 1;
      }
      if (!b) {
        return -1;
      }
      return a.localeCompare(b);
    });
    return sortedKeys.map(function (key) {
      return {
        key: key,
        label: key ? formatExportDateLabel(key) : "未設定日期 / Undated",
        tasks: groupedMap.get(key).slice().sort(sortForExportCategoryThenTime),
      };
    });
  }

  function formatExportTaskTime(task) {
    if (!task) {
      return "-";
    }
    const startAt = getTaskStartAt(task);
    const endAt = getTaskEndAt(task);
    if (!startAt && !endAt) {
      return "-";
    }
    if (task.allDay) {
      return "全天";
    }
    const startText = startAt ? formatTime(startAt, false) : "";
    const endText = endAt ? formatTime(endAt, false) : "";
    const safeStart = startText && startText !== "--:--" ? startText : "";
    const safeEnd = endText && endText !== "--:--" ? endText : "";
    if (safeStart && safeEnd) {
      return safeStart + " 至 " + safeEnd;
    }
    return safeStart || safeEnd || "-";
  }

  function formatExportTaskLine(task, options) {
    const includeDatePrefix = Boolean(options && options.includeDatePrefix);
    const categoryText = task.subcategory ? task.category + "/" + task.subcategory : task.category;
    const ownerText = task.owner ? String(task.owner).trim() : "-";
    const doneBy = task.status === "done" && task.completedBy ? task.completedBy : "-";
    const descText = task.description ? String(task.description).replace(/\s*\n+\s*/g, " / ").trim() : "-";
    return (
      "- " +
      "時間: " +
      formatExportTaskTime(task) +
      " | " +
      (includeDatePrefix ? "日期: " + formatExportTaskDate(task) + " | " : "") +
      (task.title || "-") +
      " | " +
      categoryText +
      " | " +
      "填寫人: " +
      ownerText +
      " | Done: " +
      doneBy +
      " | " +
      descText
    );
  }

  function buildTransferFyiLines(transferTasks, fyiTasks) {
    const lines = [];
    const transferLines = buildExportTaskLines(transferTasks);
    const fyiLines = buildExportTaskLines(fyiTasks);

    if (transferLines.length > 0) {
      lines.push("[Internal Transfer]");
      transferLines.forEach(function (line) {
        lines.push(line);
      });
    }

    if (fyiLines.length > 0) {
      if (lines.length > 0) {
        lines.push("");
      }
      lines.push("[F.Y.I.]");
      fyiLines.forEach(function (line) {
        lines.push(line);
      });
    }

    return lines;
  }

  function buildArrDepOccLine(exportDate) {
    const normalizedDate = normalizeDateInputFromAny(exportDate || state.queryDate || "");
    const dateLabel = normalizedDate ? formatExportDateLabel(normalizedDate) : "未來待辦事項";
    const arr = String(state.todayOverview && state.todayOverview.checkin ? state.todayOverview.checkin : "-").trim() || "-";
    const dep = String(state.todayOverview && state.todayOverview.checkout ? state.todayOverview.checkout : "-").trim() || "-";
    const occ = normalizeOccForExport(state.todayOverview ? state.todayOverview.occupancy : "");
    return dateLabel + "     Arr: " + arr + "   Dep: " + dep + "   OCC: " + occ;
  }

  function normalizeOccForExport(value) {
    const text = String(value || "").replace(/%/g, "").trim();
    if (!text) {
      return "-";
    }
    return text + " %";
  }

  function formatExportDateLabel(dateValue) {
    const normalized = String(dateValue || "").trim();
    if (!normalized) {
      return formatDateOnly(new Date());
    }
    if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
      return normalized.replace(/-/g, "/");
    }
    return formatDateOnly(normalized);
  }

  function exportLegacyDoc(tasks, conditionText, exportDate, exportStatus, includeDatePrefix) {
    const dailyList = (Array.isArray(tasks) ? tasks.slice() : []).sort(sortForExportCategoryThenTime);
    const sectionTitleHtml = includeDatePrefix ? "Future To-Do<br>未來待辦事項" : "Daily Briefing<br>每日報告";

    function toHtmlLines(taskList) {
      if (!Array.isArray(taskList) || taskList.length === 0) {
        return "（無資料）";
      }
      return taskList
        .map(function (task) {
          return (
            "<span style='color:#" +
            getExportCategoryColor(task && task.category) +
            ";'>" +
            escapeHtml(formatExportTaskLine(task, { includeDatePrefix: includeDatePrefix })) +
            "</span>"
          );
        })
        .join("<br>");
    }

    function buildSectionTableHtml(taskList) {
      return (
        "<table>" +
        "<tr><td class='left'>" +
        sectionTitleHtml +
        "</td><td>" +
        toHtmlLines(taskList) +
        "</td></tr>" +
        "</table>"
      );
    }

    let groupedSectionHtml = "";
    if (includeDatePrefix) {
      const grouped = groupExportTasksByDate(dailyList);
      if (grouped.length === 0) {
        groupedSectionHtml = buildSectionTableHtml([]);
      } else {
        groupedSectionHtml = grouped
          .map(function (group) {
            return "<p class='group-date'>" + escapeHtml(group.label) + "</p>" + buildSectionTableHtml(group.tasks);
          })
          .join("");
      }
    } else {
      groupedSectionHtml = buildSectionTableHtml(dailyList);
    }

    const arrDepOccHtml = includeDatePrefix
      ? ""
      : "<p class='arrdepocc'>" + escapeHtml(buildArrDepOccLine(exportDate)) + "</p>";
    const html =
      "<html><head><meta charset='utf-8'><style>" +
      "@page{margin:0.5in;}" +
      "body{font-family:Calibri,'DFKai-SB','標楷體','Noto Serif TC',serif;font-size:12pt;padding:0;color:#111;}" +
      "h2{margin:0 0 8px 0;}" +
      "p{margin:4px 0;}" +
      ".group-date{margin:12px 0 6px 0;font-weight:700;}" +
      "table{width:100%;border-collapse:collapse;table-layout:fixed;margin:0 auto;}" +
      "td{border:1px solid #000;padding:8px;vertical-align:middle;line-height:1.45;text-align:center;}" +
      ".left{width:160px;text-align:center;font-weight:700;}" +
      ".arrdepocc{margin-top:12px;font-weight:700;}" +
      "</style></head><body>" +
      arrDepOccHtml +
      groupedSectionHtml +
      "</body></html>";

    const blob = new Blob(["\ufeff", html], {
      type: "application/msword;charset=utf-8",
    });

    if (window.saveAs) {
      window.saveAs(blob, buildExportFileName("doc", exportDate));
      return;
    }

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = buildExportFileName("doc", exportDate);
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(link.href);
  }

  function buildExportFileName(ext, exportDate) {
    const preferredDate = String(exportDate || state.queryDate || "").trim();
    const normalized = normalizeDateInputFromAny(preferredDate);
    if (normalized && /^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
      const mmdd = normalized.slice(5, 7) + normalized.slice(8, 10);
      return mmdd + "." + ext;
    }
    return "未來待辦事項." + ext;
  }

  function normalizeQueryStatus(raw) {
    const value = String(raw || "").trim();
    if (value === "pending" || value === "done") {
      return value;
    }
    return "all";
  }

  function normalizeQueryCategory(raw) {
    const value = String(raw || "").trim();
    if (!value || value === "all") {
      return "all";
    }
    return isValidCategory(value) ? value : "all";
  }

  function normalizeQueryKeyword(raw) {
    return String(raw || "").trim().replace(/\s+/g, " ");
  }

  function normalizeTaskListFilter(raw) {
    const value = String(raw || "").trim();
    if (value === "pending" || value === "done" || value === "pinned") {
      return value;
    }
    return "all";
  }

  function getQueryStatusLabel(status) {
    if (status === "pending") {
      return "未完成";
    }
    if (status === "done") {
      return "已完成";
    }
    return "全部狀態";
  }

  function getQueryCategoryLabel(category) {
    const normalized = normalizeQueryCategory(category);
    return normalized === "all" ? "全部主分類" : normalized;
  }

  function buildQueryConditionText() {
    return buildConditionText(state.queryDate, state.queryStatus, state.queryCategory, state.queryKeyword);
  }

  function buildConditionText(dateValue, statusValue, categoryValue, keywordValue) {
    const dateText = dateValue ? dateValue.replace(/-/g, "/") : "全部日期";
    const categoryText = getQueryCategoryLabel(categoryValue);
    const keywordText = normalizeQueryKeyword(keywordValue);
    const keywordLabel = keywordText ? ("關鍵字：" + keywordText) : "關鍵字：全部";
    return dateText + " / " + getQueryStatusLabel(statusValue) + " / " + categoryText + " / " + keywordLabel;
  }

  function getTasksByDateAndStatus(dateValue, statusValue, categoryValue, keywordValue) {
    let list = state.tasks.slice().sort(sortForTaskTable);
    const normalizedCategory = normalizeQueryCategory(categoryValue);
    const normalizedKeyword = normalizeQueryKeyword(keywordValue).toLowerCase();
    if (dateValue) {
      list = list.filter(function (task) {
        return isTaskInDateRange(task, dateValue);
      });
    }
    if (statusValue !== "all") {
      list = list.filter(function (task) {
        return task.status === statusValue;
      });
    }
    if (normalizedCategory !== "all") {
      list = list.filter(function (task) {
        return task.category === normalizedCategory;
      });
    }
    if (normalizedKeyword) {
      list = list.filter(function (task) {
        const haystack = [
          task.category,
          task.subcategory,
          task.title,
          task.owner,
          task.completedBy,
          task.description,
        ]
          .map(function (value) {
            return String(value || "");
          })
          .join(" ")
          .toLowerCase();
        return haystack.indexOf(normalizedKeyword) !== -1;
      });
    }
    return list;
  }

  function copyTaskText(task) {
    const text = buildTaskCopyText(task);
    if (!text) {
      showToast("無可複製內容。");
      return;
    }
    copyTextToClipboard(text).then(function (ok) {
      showToast(ok ? "已複製事項文字。" : "複製失敗，請手動複製。");
    });
  }

  function buildTaskCopyText(task) {
    if (!task) {
      return "";
    }
    const isDone = task.status === "done";
    const lines = [];
    lines.push("主分類：" + (task.category || "-"));
    lines.push("子分類：" + (task.subcategory || "-"));
    lines.push("事項：" + (task.title || "-"));
    lines.push("填寫人：" + (task.owner || "-"));
    if (isDone) {
      lines.push("完成人：" + (task.completedBy || "-"));
    }
    lines.push("時間：" + formatTimeRange(task));
    lines.push("內容：" + (task.description || "-"));
    return lines.join("\n");
  }

  function copyTextToClipboard(text) {
    const value = String(text || "");
    if (!value) {
      return Promise.resolve(false);
    }
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard
        .writeText(value)
        .then(function () {
          return true;
        })
        .catch(function () {
          return fallbackCopyText(value);
        });
    }
    return Promise.resolve(fallbackCopyText(value));
  }

  function fallbackCopyText(text) {
    if (!document.body) {
      return false;
    }
    const textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "readonly");
    textarea.style.position = "fixed";
    textarea.style.top = "-9999px";
    textarea.style.left = "-9999px";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();
    textarea.setSelectionRange(0, textarea.value.length);
    let copied = false;
    try {
      copied = document.execCommand("copy");
    } catch (error) {
      copied = false;
    }
    textarea.remove();
    return copied;
  }

  function getTodayDateKey() {
    return toDateKey(new Date());
  }

  function createEmptyTodayOverview(dateKey) {
    const key = normalizeDateInputFromAny(dateKey || "") || getTodayDateKey();
    return {
      checkin: "",
      checkout: "",
      occupancy: "",
      dateKey: key,
    };
  }

  function normalizeTodayOverviewRecord(raw, clearWhenNotToday) {
    const source = raw && typeof raw === "object" ? raw : {};
    const todayKey = getTodayDateKey();
    const recordDateKey = normalizeDateInputFromAny(source.dateKey || source._dateKey || source.date || "");
    if (clearWhenNotToday && recordDateKey !== todayKey) {
      return createEmptyTodayOverview(todayKey);
    }
    return {
      checkin: normalizeTodayOverviewValue(source.checkin),
      checkout: normalizeTodayOverviewValue(source.checkout),
      occupancy: normalizeOccupancyRateValue(source.occupancy),
      dateKey: recordDateKey || todayKey,
    };
  }

  function loadTodayOverview() {
    const defaults = createEmptyTodayOverview(getTodayDateKey());
    try {
      const raw = localStorage.getItem(getScopedStorageKey(TODAY_OVERVIEW_KEY));
      if (!raw) {
        return defaults;
      }
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") {
        return defaults;
      }
      return normalizeTodayOverviewRecord(parsed, true);
    } catch (error) {
      console.error("loadTodayOverview error", error);
      return defaults;
    }
  }

  function saveTodayOverview() {
    state.todayOverview = normalizeTodayOverviewRecord(state.todayOverview, true);
    localStorage.setItem(getScopedStorageKey(TODAY_OVERVIEW_KEY), JSON.stringify(state.todayOverview));
    scheduleCloudPush();
  }

  function normalizeTodayOverviewValue(raw) {
    return String(raw || "").trim();
  }

  function normalizeOccupancyRateValue(raw) {
    const value = String(raw || "")
      .replace(/[%％]/g, "")
      .trim();
    if (!value) {
      return "";
    }
    const number = Number(value);
    if (!Number.isFinite(number)) {
      return "";
    }
    const clamped = Math.min(100, Math.max(0, number));
    const rounded = Math.round(clamped * 100) / 100;
    return rounded.toFixed(2);
  }

  function buildId() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return window.crypto.randomUUID();
    }
    return "task_" + Date.now() + "_" + Math.random().toString(16).slice(2);
  }

  function touchTask(task) {
    if (!task || typeof task !== "object") {
      return;
    }
    task.updatedAt = new Date().toISOString();
  }

  function getTaskRevisionMs(task) {
    if (!task) {
      return 0;
    }
    const raw = task.updatedAt || task.createdAt;
    if (!raw) {
      return 0;
    }
    const ms = new Date(raw).getTime();
    return Number.isFinite(ms) ? ms : 0;
  }

  function mergeTasksById(baseTasks, incomingTasks) {
    const merged = new Map();
    (Array.isArray(baseTasks) ? baseTasks : []).forEach(function (task) {
      const normalized = normalizeTask(task);
      if (!normalized) {
        return;
      }
      merged.set(normalized.id, normalized);
    });

    (Array.isArray(incomingTasks) ? incomingTasks : []).forEach(function (task) {
      const normalized = normalizeTask(task);
      if (!normalized) {
        return;
      }
      const existing = merged.get(normalized.id);
      if (!existing) {
        merged.set(normalized.id, normalized);
        return;
      }
      if (getTaskRevisionMs(normalized) >= getTaskRevisionMs(existing)) {
        merged.set(normalized.id, normalized);
      }
    });

    return Array.from(merged.values()).sort(sortByDueTime);
  }

  function mergeTodayOverview(baseOverview, incomingOverview, options) {
    const base = normalizeTodayOverviewRecord(baseOverview, true);
    const incoming = normalizeTodayOverviewRecord(incomingOverview, true);
    const preferBase = Boolean(options && options.preferBaseOverview);
    const incomingCheckin = normalizeTodayOverviewValue(incoming.checkin);
    const incomingCheckout = normalizeTodayOverviewValue(incoming.checkout);
    const incomingOccupancy = normalizeOccupancyRateValue(incoming.occupancy);
    const baseCheckin = normalizeTodayOverviewValue(base.checkin);
    const baseCheckout = normalizeTodayOverviewValue(base.checkout);
    const baseOccupancy = normalizeOccupancyRateValue(base.occupancy);
    const mergedDateKey = getTodayDateKey();
    if (preferBase) {
      return {
        checkin: baseCheckin || incomingCheckin,
        checkout: baseCheckout || incomingCheckout,
        occupancy: baseOccupancy || incomingOccupancy,
        dateKey: mergedDateKey,
      };
    }
    return {
      checkin: incomingCheckin || baseCheckin,
      checkout: incomingCheckout || baseCheckout,
      occupancy: incomingOccupancy || baseOccupancy,
      dateKey: mergedDateKey,
    };
  }

  function getDeletedTaskRevisionMs(deletedTaskIds, taskId) {
    if (!deletedTaskIds || typeof deletedTaskIds !== "object") {
      return 0;
    }
    const raw = deletedTaskIds[taskId];
    if (!raw) {
      return 0;
    }
    const ms = new Date(raw).getTime();
    return Number.isFinite(ms) ? ms : 0;
  }

  function normalizeDeletedTaskIds(input) {
    const source = input && typeof input === "object" ? input : {};
    const result = {};
    Object.keys(source).forEach(function (taskId) {
      const normalizedId = String(taskId || "").trim();
      if (!normalizedId) {
        return;
      }
      const ms = new Date(source[taskId]).getTime();
      if (!Number.isFinite(ms)) {
        return;
      }
      result[normalizedId] = new Date(ms).toISOString();
    });
    return result;
  }

  function mergeDeletedTaskIds(baseDeletedTaskIds, incomingDeletedTaskIds) {
    const merged = normalizeDeletedTaskIds(baseDeletedTaskIds);
    const incoming = normalizeDeletedTaskIds(incomingDeletedTaskIds);
    Object.keys(incoming).forEach(function (taskId) {
      const incomingMs = getDeletedTaskRevisionMs(incoming, taskId);
      const currentMs = getDeletedTaskRevisionMs(merged, taskId);
      if (incomingMs >= currentMs) {
        merged[taskId] = incoming[taskId];
      }
    });
    return merged;
  }

  function filterTasksByDeletedMap(tasks, deletedTaskIds) {
    const deleted = normalizeDeletedTaskIds(deletedTaskIds);
    const result = [];
    (Array.isArray(tasks) ? tasks : []).forEach(function (task) {
      const normalized = normalizeTask(task);
      if (!normalized) {
        return;
      }
      const deletedMs = getDeletedTaskRevisionMs(deleted, normalized.id);
      if (deletedMs === 0) {
        result.push(normalized);
        return;
      }
      if (getTaskRevisionMs(normalized) > deletedMs) {
        result.push(normalized);
        delete deleted[normalized.id];
      }
    });
    return {
      tasks: result.sort(sortByDueTime),
      deletedTaskIds: deleted,
    };
  }

  function mergeBackupState(baseState, incomingState, options) {
    const base = baseState && typeof baseState === "object" ? baseState : {};
    const incoming = incomingState && typeof incomingState === "object" ? incomingState : {};
    const mergedDeletedTaskIds = mergeDeletedTaskIds(base.deletedTaskIds, incoming.deletedTaskIds);
    const filtered = filterTasksByDeletedMap(mergeTasksById(base.tasks, incoming.tasks), mergedDeletedTaskIds);
    return {
      tasks: filtered.tasks,
      todayOverview: mergeTodayOverview(base.todayOverview, incoming.todayOverview, options),
      deletedTaskIds: filtered.deletedTaskIds,
    };
  }

  function getTaskStartAt(task) {
    if (!task) {
      return null;
    }
    return task.startAt || task.dueAt || null;
  }

  function getTaskEndAt(task) {
    if (!task) {
      return null;
    }
    return task.endAt || null;
  }

  function isTaskInDateRange(task, dateKey) {
    const key = String(dateKey || "").trim();
    if (!key) {
      return true;
    }
    const startAt = getTaskStartAt(task);
    const endAt = getTaskEndAt(task);
    const startKey = startAt ? toDateKey(new Date(startAt)) : "";
    const endKey = endAt ? toDateKey(new Date(endAt)) : "";
    if (!startKey && !endKey) {
      return false;
    }
    if (startKey && endKey) {
      const rangeStart = startKey <= endKey ? startKey : endKey;
      const rangeEnd = startKey <= endKey ? endKey : startKey;
      return key >= rangeStart && key <= rangeEnd;
    }
    const singleDay = startKey || endKey;
    return key === singleDay;
  }

  function isTaskNotPastByDateKey(task, baseDateKey) {
    const key = String(baseDateKey || "").trim();
    if (!key) {
      return true;
    }
    const startAt = getTaskStartAt(task);
    const endAt = getTaskEndAt(task);
    if (!startAt && !endAt) {
      return true;
    }
    const endKey = endAt ? toDateKey(new Date(endAt)) : "";
    if (endKey) {
      return endKey >= key;
    }
    const startKey = startAt ? toDateKey(new Date(startAt)) : "";
    if (!startKey) {
      return true;
    }
    return startKey >= key;
  }

  function sortByDueTime(a, b) {
    const aStartAt = getTaskStartAt(a);
    const bStartAt = getTaskStartAt(b);
    const aMs = aStartAt ? new Date(aStartAt).getTime() : Number.POSITIVE_INFINITY;
    const bMs = bStartAt ? new Date(bStartAt).getTime() : Number.POSITIVE_INFINITY;
    return aMs - bMs;
  }

  function sortForTaskTable(a, b) {
    if (Boolean(a.pinned) !== Boolean(b.pinned)) {
      return a.pinned ? -1 : 1;
    }
    return sortByDueTime(a, b);
  }

  function getCategorySortIndex(category) {
    const idx = CATEGORIES.indexOf(String(category || "").trim());
    return idx === -1 ? Number.MAX_SAFE_INTEGER : idx;
  }

  function sortForExportCategoryThenTime(a, b) {
    const categoryDiff = getCategorySortIndex(a && a.category) - getCategorySortIndex(b && b.category);
    if (categoryDiff !== 0) {
      return categoryDiff;
    }
    const timeDiff = sortByDueTime(a, b);
    if (timeDiff !== 0) {
      return timeDiff;
    }
    const aTitle = String((a && a.title) || "");
    const bTitle = String((b && b.title) || "");
    return aTitle.localeCompare(bTitle, "zh-Hant");
  }

  function normalizeCategory(raw) {
    const value = String(raw || "").trim();
    if (CATEGORIES.indexOf(value) === -1) {
      return CATEGORIES[0];
    }
    return value;
  }

  function isValidCategory(value) {
    return CATEGORIES.indexOf(value) !== -1;
  }

  function getSubcategoryOptions(category) {
    const list = SUBCATEGORY_MAP[category];
    return Array.isArray(list) ? list.slice() : [];
  }

  function hasSubcategoryOptions(category) {
    return getSubcategoryOptions(category).length > 0;
  }

  function normalizeSubcategory(category, raw) {
    let value = String(raw || "").trim();
    if (category === "廣場" && value === "工程") {
      value = "其他";
    }
    const options = getSubcategoryOptions(category);
    if (options.length === 0) {
      return "";
    }
    return options.indexOf(value) === -1 ? "" : value;
  }

  function toDateTimeLocalValue(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hour = String(date.getHours()).padStart(2, "0");
    const minute = String(date.getMinutes()).padStart(2, "0");
    return year + "-" + month + "-" + day + " " + hour + ":" + minute;
  }

  function toDateKey(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return year + "-" + month + "-" + day;
  }

  function parseTaskTimeRange(startDateInput, startInput, endDateInput, endInput, allDay) {
    const startDateText = normalizeDateInputFromAny(startDateInput);
    const rawEndDateText = normalizeDateInputFromAny(endDateInput);
    const endDateText = rawEndDateText || startDateText;
    const startText = String(startInput || "").trim();
    const endText = String(endInput || "").trim();

    if ((startText || endText || rawEndDateText || allDay) && !startDateText) {
      return {
        ok: false,
        message: "請先選擇開始日期。",
      };
    }
    if (!startDateText && !startText && !endText && !rawEndDateText) {
      return {
        ok: true,
        startAtIso: null,
        endAtIso: null,
        message: "",
      };
    }

    if (allDay) {
      const startMs = parseTaskTimeInput(startDateText, true, "start");
      const endMs = parseTaskTimeInput(endDateText || startDateText, true, "end");
      if (Number.isNaN(startMs) || Number.isNaN(endMs)) {
        return {
          ok: false,
          message: "日期格式無效，請重新選擇。",
        };
      }
      if (endMs < startMs) {
        return {
          ok: false,
          message: "結束日期不可早於開始日期。",
        };
      }
      return {
        ok: true,
        startAtIso: new Date(startMs).toISOString(),
        endAtIso: new Date(endMs).toISOString(),
        message: "",
      };
    }

    let startMs = Number.NaN;
    let endMs = Number.NaN;

    if (startText) {
      startMs = parseTaskTimeInput(startText, false, "start", startDateText);
    }
    if (endText) {
      endMs = parseTaskTimeInput(endText, false, "end", endDateText);
    }

    if (startText && Number.isNaN(startMs)) {
      return {
        ok: false,
        message: "開始時間格式無效，請重新輸入。",
      };
    }
    if (endText && Number.isNaN(endMs)) {
      return {
        ok: false,
        message: "結束時間格式無效，請重新輸入。",
      };
    }

    if (rawEndDateText && !startText) {
      startMs = parseTaskTimeInput(startDateText, true, "start");
    }
    if (rawEndDateText && !endText) {
      endMs = parseTaskTimeInput(endDateText, true, "end");
    }

    if (!Number.isNaN(startMs) && !Number.isNaN(endMs) && endMs < startMs) {
      return {
        ok: false,
        message: "結束時間不可早於開始時間。",
      };
    }

    return {
      ok: true,
      startAtIso: Number.isNaN(startMs) ? null : new Date(startMs).toISOString(),
      endAtIso: Number.isNaN(endMs) ? null : new Date(endMs).toISOString(),
      message: "",
    };
  }

  function parseTaskTimeInput(value, allDay, part, baseDateInput) {
    const text = String(value || "").trim();
    if (!text) {
      return Number.NaN;
    }
    const normalized = text
      .replace(/[\uFF0F/]/g, "-")
      .replace(/\s+/g, " ")
      .trim();
    if (allDay) {
      const dayMatch = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (!dayMatch) {
        return Number.NaN;
      }
      return buildDateTimeMs(
        Number(dayMatch[1]),
        Number(dayMatch[2]),
        Number(dayMatch[3]),
        part === "end" ? 23 : 0,
        part === "end" ? 59 : 0,
      );
    }
    const timeOnly = normalized.match(/^(\d{1,2}):(\d{2})$/);
    if (timeOnly) {
      const baseDate = normalizeDateInputFromAny(baseDateInput) || toDateKey(new Date());
      const dayMatch = baseDate.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (!dayMatch) {
        return Number.NaN;
      }
      return buildDateTimeMs(
        Number(dayMatch[1]),
        Number(dayMatch[2]),
        Number(dayMatch[3]),
        Number(timeOnly[1]),
        Number(timeOnly[2]),
      );
    }
    const dtMatch = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2})$/);
    if (dtMatch) {
      return buildDateTimeMs(
        Number(dtMatch[1]),
        Number(dtMatch[2]),
        Number(dtMatch[3]),
        Number(dtMatch[4]),
        Number(dtMatch[5]),
      );
    }

    const fallbackMs = new Date(text).getTime();
    return Number.isNaN(fallbackMs) ? Number.NaN : fallbackMs;
  }

  function buildDateTimeMs(year, month, day, hour, minute) {
    if (
      !Number.isFinite(year) ||
      !Number.isFinite(month) ||
      !Number.isFinite(day) ||
      !Number.isFinite(hour) ||
      !Number.isFinite(minute)
    ) {
      return Number.NaN;
    }
    if (month < 1 || month > 12 || day < 1 || day > 31 || hour < 0 || hour > 23 || minute < 0 || minute > 59) {
      return Number.NaN;
    }
    const date = new Date(year, month - 1, day, hour, minute, 0, 0);
    if (
      date.getFullYear() !== year ||
      date.getMonth() !== month - 1 ||
      date.getDate() !== day ||
      date.getHours() !== hour ||
      date.getMinutes() !== minute
    ) {
      return Number.NaN;
    }
    return date.getTime();
  }

  function formatDueDisplay(task) {
    if (!task) {
      return "-";
    }
    const startAt = getTaskStartAt(task);
    const endAt = getTaskEndAt(task);
    if (!startAt && !endAt) {
      return "-";
    }
    if (task.allDay) {
      if (startAt && endAt) {
        const startDate = toDateKey(new Date(startAt));
        const endDate = toDateKey(new Date(endAt));
        if (startDate === endDate) {
          return formatDateOnly(startAt) + " 全天";
        }
        return formatDateOnly(startAt) + " 至 " + formatDateOnly(endAt) + " 全天";
      }
      const allDayBase = startAt || endAt;
      return formatDateOnly(allDayBase) + " 全天";
    }
    if (startAt && endAt) {
      const sameDate = toDateKey(new Date(startAt)) === toDateKey(new Date(endAt));
      if (sameDate) {
        return formatDateOnly(startAt) + " " + formatTime(startAt, false) + " 至 " + formatTime(endAt, false);
      }
      return formatDateTime(startAt) + " 至 " + formatDateTime(endAt);
    }
    if (startAt) {
      return formatDateTime(startAt);
    }
    return formatDateTime(endAt);
  }

  function formatDateTime(input) {
    if (!input) {
      return "-";
    }
    const date = input instanceof Date ? input : new Date(input);
    if (Number.isNaN(date.getTime())) {
      return "-";
    }
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hour = String(date.getHours()).padStart(2, "0");
    const minute = String(date.getMinutes()).padStart(2, "0");
    return year + "/" + month + "/" + day + " " + hour + ":" + minute;
  }

  function formatDateOnly(input) {
    if (!input) {
      return "-";
    }
    const date = input instanceof Date ? input : new Date(input);
    if (Number.isNaN(date.getTime())) {
      return "-";
    }
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return year + "/" + month + "/" + day;
  }

  function formatTime(input, allDay) {
    if (allDay) {
      return "全天";
    }
    if (!input) {
      return "--:--";
    }
    const date = input instanceof Date ? input : new Date(input);
    if (Number.isNaN(date.getTime())) {
      return "--:--";
    }
    const hour = String(date.getHours()).padStart(2, "0");
    const minute = String(date.getMinutes()).padStart(2, "0");
    return hour + ":" + minute;
  }

  function formatTimeRange(task) {
    if (!task) {
      return "--:--";
    }
    if (task.allDay) {
      return "全天";
    }
    const startAt = getTaskStartAt(task);
    const endAt = getTaskEndAt(task);
    if (!startAt && !endAt) {
      return "--:--";
    }
    if (startAt && endAt) {
      return formatTime(startAt, false) + " 至 " + formatTime(endAt, false);
    }
    return formatTime(startAt || endAt, false);
  }

  function formatCountdown(input, allDay, isDone) {
    if (isDone) {
      return "已完成";
    }
    if (allDay) {
      return "全天事項";
    }
    if (!input) {
      return "無時間";
    }
    const dueMs = input instanceof Date ? input.getTime() : new Date(input).getTime();
    if (Number.isNaN(dueMs)) {
      return "無時間";
    }
    const diffSec = Math.floor((dueMs - Date.now()) / 1000);
    if (diffSec >= 0) {
      return "剩餘 " + formatDuration(diffSec);
    }
    return "逾時 " + formatDuration(Math.abs(diffSec));
  }

  function formatDuration(totalSeconds) {
    const sec = Math.max(0, Math.floor(totalSeconds));
    const days = Math.floor(sec / 86400);
    const hours = Math.floor((sec % 86400) / 3600);
    const minutes = Math.floor((sec % 3600) / 60);
    const seconds = sec % 60;
    const hh = String(hours).padStart(2, "0");
    const mm = String(minutes).padStart(2, "0");
    const ss = String(seconds).padStart(2, "0");
    if (days > 0) {
      return days + "天 " + hh + ":" + mm + ":" + ss;
    }
    return hh + ":" + mm + ":" + ss;
  }

  function escapeHtml(text) {
    return String(text)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function showToast(message) {
    els.toast.textContent = message;
    els.toast.classList.add("show");
    if (state.toastTimer) {
      clearTimeout(state.toastTimer);
    }
    state.toastTimer = setTimeout(function () {
      els.toast.classList.remove("show");
    }, TOAST_MS);
  }
})();

