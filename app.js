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
  const PUSH_SYNC_INTERVAL_MS = 30000;
  const PUSH_SERVICE_WORKER_FILE = "sw.js";
  const LEGACY_MIGRATION_DONE_KEY = "handover_legacy_kvdb_migration_done_v1";
  const BACKUP_TYPE = "handover-backup";
  const BACKUP_VERSION = "0.95";
  const UI_LANGUAGE_STORAGE_KEY = "handover_ui_language_v1";
  const DEFAULT_UI_LANGUAGE = "zh";
  const REMINDER_CHECK_MS = 30 * 1000;
  const COUNTDOWN_REFRESH_MS = 1000;
  const TOAST_MS = 3000;
  const MOBILE_MAX_WIDTH = 760;
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
  const CATEGORY_LABEL_MAP = Object.freeze({
    en: Object.freeze({
      廣場: "Plaza",
      包裹代收: "Parcel Desk",
      車輛安排: "Transport",
      大廳: "Lobby",
      會議室: "Meeting Room",
      團桌: "Banquet Table",
      客房: "Guest Room",
      預訂: "Reservation",
      餐飲部: "F&B",
      待回覆信件: "Pending Reply",
      郵寄: "Mail",
      行政: "Administration",
      公告: "Notice",
      遺留物: "Lost & Found",
    }),
  });
  const SUBCATEGORY_LABEL_MAP = Object.freeze({
    en: Object.freeze({
      保留車位: "Reserved Parking",
      其他: "Other",
      團體: "Group",
      散客: "Individual",
      禮賓車: "Limousine",
      計程車: "Taxi",
      行李寄放: "Luggage Storage",
      下行李: "Luggage Delivery",
      房務相關事項: "Housekeeping",
      送房: "In-room Delivery",
      佈置: "Decoration",
      餐廳: "Restaurant",
      車票: "Train Ticket",
      叫貨領貨: "Supply Pickup",
      人事相關: "HR Related",
      待寄: "To Ship",
      待取: "To Pick Up",
    }),
  });
  const UI_TEXT = Object.freeze({
    zh: Object.freeze({
      appTitle: "工作交接助手",
      passwordSubtitle: "請輸入伺服器名稱與密碼進入系統（test 伺服器需密碼登入）。",
      passwordServerLabel: "伺服器名稱（英數、-、_）",
      passwordServerPlaceholder: "例：test / caesarmetro",
      passwordLabel: "密碼",
      passwordPlaceholder: "請輸入密碼（若伺服器要求）",
      passwordSubmit: "進入系統",
      languageLabel: "介面",
      todaySectionTitle: "當日事項",
      todayDateLabel: "當日日期",
      todayCheckinLabel: "預進",
      todayCheckoutLabel: "預退",
      todayOccLabel: "住房率",
      save: "儲存",
      todayPinnedTitle: "置頂區",
      todayFilterLabel: "主分類",
      taskFormTitle: "新增交接待辦",
      taskCategoryLabel: "主分類",
      taskSubcategoryLabel: "子分類",
      taskTitleLabel: "事項名稱",
      taskTimeLabel: "時間（可跨日：起日/起時 至 迄日/迄時，可留空）",
      timeRangeSep: "至",
      allDay: "全日",
      taskOwnerLabel: "填寫人",
      taskAssigneeLabel: "指派給",
      taskAssigneeUnassigned: "未指定（公用）",
      taskPinLabel: "置頂顯示",
      taskPinHelp: "一直顯示在網頁上",
      taskDescLabel: "交接說明",
      addTask: "新增待辦",
      clear: "清空",
      cancelEdit: "取消修改",
      queryPanelTitle: "日期查詢",
      resultPanelTitle: "查詢結果",
      taskListFilterLabel: "篩選",
      exportDateLabel: "匯出日期",
      exportStatusLabel: "匯出狀態",
      exportWord: "一鍵生成 Word",
      exportExcel: "一鍵生成 Excel",
      upcomingTitle: "下一件待辦提醒",
      close: "關閉",
      notificationToggleLabel: "通知開啟",
      requestNotification: "授權通知",
      mobileUpcomingLabel: "下一件待辦",
      collapse: "收合",
      expand: "展開",
      allCategories: "全部主分類",
      allStatus: "全部狀態",
      pending: "未完成",
      done: "已完成",
      pinned: "置頂中",
      all: "全部",
      keywordPlaceholder: "關鍵字查詢（事項/說明/填寫人）",
      taskTitlePlaceholder: "例：房號、客人姓名、團體名稱等",
      taskOwnerPlaceholder: "例：王小明",
      taskDescriptionPlaceholder: "補充細節、檔案位置、提醒事項",
      search: "查詢",
      clearSearch: "清除",
      todayNoPinned: "目前沒有置頂事項。",
      todayNoSchedule: "今日沒有排程事項。",
      todayNoNormal: "今日一般事項為 0 筆。",
      tableHeaders: ["狀態", "置頂", "主分類", "子分類", "事項", "填寫人", "完成人", "時間", "交接說明", "操作"],
      statusPending: "待處理",
      statusDone: "已完成",
      actionComplete: "完成",
      actionUndoComplete: "取消完成",
      actionEdit: "修改",
      actionCopy: "一鍵複製",
      actionCancel: "取消",
      actionPin: "置頂",
      actionUnpin: "取消置頂",
      actionDelete: "刪除",
      actionTranslateTask: "翻譯此筆",
      actionTranslating: "翻譯中...",
      actionShowOriginal: "顯示原文",
      actionShowTranslated: "顯示翻譯",
      noTranslatedContent: "此內容目前無翻譯可切換。",
      owner: "填寫人",
      assignee: "指派",
      completedBy: "完成人",
      countdown: "倒數",
      content: "內容",
      time: "時間",
      uncategorized: "未分類",
      notFilled: "未填寫",
      noData: "（無）",
      categoryTipSelectCategory: "請先選主分類，再填寫其餘欄位。",
      categoryTipNeedSub: "此主分類需要子分類，請先選擇子分類。",
      categoryTipCurrent: "目前分類：{category}",
      categoryTipCurrentSub: "目前分類：{category} / {subcategory}",
      subcategoryChooseCategory: "請先選主分類",
      subcategoryNone: "此主分類無子分類",
      subcategoryChoose: "請選擇子分類",
      queryAllText: "目前顯示全部待辦，共 {count} 筆。",
      queryConditionText: "目前顯示「{condition}」，共 {count} 筆。",
      todaySummaryZero: "今日共 0 筆事項（{category}）。",
      todaySummary: "今日共 {total} 筆，待處理 {pending} 筆，已完成 {done} 筆（{category}）。",
      taskTableEmpty: "目前沒有符合條件的待辦事項。",
      upcomingSummaryEmpty: "目前沒有 1 小時內待辦。",
      upcomingListEmpty: "30 分鐘與 1 小時提醒區間目前無待辦。",
      timeWindowDue: "已到時間",
      timeWindow30: "30 分鐘內",
      timeWindow60: "1 小時內",
    }),
    en: Object.freeze({
      appTitle: "Handover Assistant",
      passwordSubtitle: "Enter server name and password (test server requires password login).",
      passwordServerLabel: "Server (letters, numbers, -, _)",
      passwordServerPlaceholder: "e.g. test / caesarmetro",
      passwordLabel: "Password",
      passwordPlaceholder: "Enter password if required",
      passwordSubmit: "Sign In",
      languageLabel: "Language",
      todaySectionTitle: "Today Board",
      todayDateLabel: "Date",
      todayCheckinLabel: "Arr",
      todayCheckoutLabel: "Dep",
      todayOccLabel: "Occ",
      save: "Save",
      todayPinnedTitle: "Pinned",
      todayFilterLabel: "Category",
      taskFormTitle: "Add Handover Task",
      taskCategoryLabel: "Category",
      taskSubcategoryLabel: "Subcategory",
      taskTitleLabel: "Task Title",
      taskTimeLabel: "Time (cross-day allowed: start date/time to end date/time, optional)",
      timeRangeSep: "to",
      allDay: "All Day",
      taskOwnerLabel: "Owner",
      taskAssigneeLabel: "Assign To",
      taskAssigneeUnassigned: "Unassigned (Shared)",
      taskPinLabel: "Pin Display",
      taskPinHelp: "Keep visible on page",
      taskDescLabel: "Handover Notes",
      addTask: "Add Task",
      clear: "Clear",
      cancelEdit: "Cancel Edit",
      queryPanelTitle: "Date Query",
      resultPanelTitle: "Query Result",
      taskListFilterLabel: "Filter",
      exportDateLabel: "Export Date",
      exportStatusLabel: "Export Status",
      exportWord: "Export Word",
      exportExcel: "Export Excel",
      upcomingTitle: "Upcoming Reminder",
      close: "Close",
      notificationToggleLabel: "Notifications On",
      requestNotification: "Allow Notifications",
      mobileUpcomingLabel: "Upcoming",
      collapse: "Collapse",
      expand: "Expand",
      allCategories: "All Categories",
      allStatus: "All Status",
      pending: "Pending",
      done: "Done",
      pinned: "Pinned",
      all: "All",
      keywordPlaceholder: "Keyword (title/notes/owner)",
      taskTitlePlaceholder: "e.g. Room no., guest name, group name",
      taskOwnerPlaceholder: "e.g. Dennis",
      taskDescriptionPlaceholder: "Details, file location, reminders",
      search: "Search",
      clearSearch: "Clear",
      todayNoPinned: "No pinned items.",
      todayNoSchedule: "No scheduled items for today.",
      todayNoNormal: "No regular pending items today.",
      tableHeaders: ["Status", "Pinned", "Category", "Subcategory", "Task", "Owner", "Done By", "Time", "Notes", "Actions"],
      statusPending: "Pending",
      statusDone: "Done",
      actionComplete: "Done",
      actionUndoComplete: "Undo",
      actionEdit: "Edit",
      actionCopy: "Copy",
      actionCancel: "Cancel",
      actionPin: "Pin",
      actionUnpin: "Unpin",
      actionDelete: "Delete",
      actionTranslateTask: "Translate This",
      actionTranslating: "Translating...",
      actionShowOriginal: "Show Original",
      actionShowTranslated: "Show Translation",
      noTranslatedContent: "No translated content available for this field.",
      owner: "Owner",
      assignee: "Assignee",
      completedBy: "Done by",
      countdown: "Countdown",
      content: "Notes",
      time: "Time",
      uncategorized: "Uncategorized",
      notFilled: "Not set",
      noData: "(none)",
      categoryTipSelectCategory: "Select a category first, then complete the form.",
      categoryTipNeedSub: "This category requires a subcategory.",
      categoryTipCurrent: "Current: {category}",
      categoryTipCurrentSub: "Current: {category} / {subcategory}",
      subcategoryChooseCategory: "Choose category first",
      subcategoryNone: "No subcategory for this category",
      subcategoryChoose: "Choose subcategory",
      queryAllText: "Showing all tasks, total {count}.",
      queryConditionText: "Showing \"{condition}\", total {count}.",
      todaySummaryZero: "Total 0 item(s) today ({category}).",
      todaySummary: "Total {total}, pending {pending}, done {done} ({category}).",
      taskTableEmpty: "No tasks match current filters.",
      upcomingSummaryEmpty: "No tasks within 1 hour.",
      upcomingListEmpty: "No tasks in 30-min / 1-hour windows.",
      timeWindowDue: "Due now",
      timeWindow30: "Within 30 min",
      timeWindow60: "Within 1 hour",
    }),
  });
  const EXPORT_FONT_EAST_ASIA = "DFKai-SB";
  const EXPORT_FONT_LATIN = "DFKai-SB";
  const EXPORT_DEFAULT_COLOR = "1F2A2A";
  const EXCEL_DAY_BLOCK_COLORS = Object.freeze(["FFF4DD", "E8F3FF", "EAF8E6", "FCEEF6", "EFEAFF", "E8F6F5"]);
  const EXCEL_BLOCK_BORDER_COLOR = "B89B68";
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
  const TEST_EXPORT_CATEGORY_COLORS = Object.freeze({
    廣場: "222222",
    包裹代收: "2A2A2A",
    車輛安排: "323232",
    大廳: "3A3A3A",
    會議室: "424242",
    團桌: "4A4A4A",
    客房: "525252",
    預訂: "5A5A5A",
    餐飲部: "626262",
    待回覆信件: "6A6A6A",
    郵寄: "727272",
    行政: "7A7A7A",
    公告: "828282",
    遺留物: "8A8A8A",
  });
  const USER_PROFILE_MAP = Object.freeze({
    caesarmetro: Object.freeze({
      name: "caesarmetro",
      themeClass: "",
      themeVars: Object.freeze({}),
      categories: CATEGORIES.slice(),
      subcategoryMap: SUBCATEGORY_MAP,
      exportCategoryColors: EXPORT_CATEGORY_COLORS,
    }),
    test: Object.freeze({
      name: "test",
      themeClass: "theme-mono",
      themeVars: Object.freeze({
        "--bg-ivory": "#f4f4f4",
        "--bg-sand": "#e4e4e4",
        "--bg-leaf": "#2d2d2d",
        "--bg-leaf-deep": "#171717",
        "--panel-bg": "rgba(250, 250, 250, 0.96)",
        "--panel-border": "rgba(96, 96, 96, 0.34)",
        "--panel-line": "rgba(120, 120, 120, 0.86)",
        "--text-main": "#202020",
        "--text-muted": "#5a5a5a",
        "--text-soft": "#6e6e6e",
        "--gold": "#7a7a7a",
        "--gold-soft": "#b7b7b7",
        "--green-soft": "#dcdcdc",
        "--danger-soft": "#ebebeb",
        "--content-card-bg": "#f7f7f7",
        "--content-card-bg-soft": "#ececec",
        "--status-pending-bg": "linear-gradient(120deg, #e1e1e1, #cfcfcf)",
        "--status-pending-text": "#202020",
        "--status-done-bg": "linear-gradient(120deg, #dddddd, #c8c8c8)",
        "--status-done-text": "#202020",
        "--status-overdue-bg": "linear-gradient(120deg, #e8e8e8, #d4d4d4)",
        "--status-overdue-text": "#202020",
        "--table-header-bg": "rgba(150, 150, 150, 0.2)",
      }),
      categories: CATEGORIES.slice(),
      subcategoryMap: SUBCATEGORY_MAP,
      exportCategoryColors: TEST_EXPORT_CATEGORY_COLORS,
    }),
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
    mobileUpcomingOpen: false,
    currentServerId: null,
    currentServerConfig: null,
    currentProfile: null,
    authToken: "",
    authSession: null,
    assignableAccounts: [],
    serviceWorkerRegistration: null,
    pushSyncTimer: null,
    uiLanguage: DEFAULT_UI_LANGUAGE,
    appliedThemeVarKeys: [],
    cloudInitDone: false,
    cloudPushTimer: null,
    cloudPushInFlight: false,
    cloudPushQueued: false,
    cloudMutePush: false,
    cloudLastErrorAt: 0,
    deletedTaskIds: {},
    showOriginalByTaskId: {},
    translatingTaskIds: {},
    toastTimer: null,
    initialized: false,
  };

  const els = {};

  document.addEventListener("DOMContentLoaded", bootstrap);

  function bootstrap() {
    cacheElements();
    state.uiLanguage = loadUiLanguage();
    applyUiLanguageToStatic();
    initAccessGate();
  }

  function init() {
    if (state.initialized) {
      return;
    }
    ensureServerContext();
    state.currentProfile = resolveProfileForServer(state.currentServerId || DEFAULT_SERVER_ID);
    applyProfileTheme(state.currentProfile);
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
    syncMobileReminderUi();
    syncCollapsiblePanels();
    setupCategorySelectOptions();
    refreshAssigneeSelectOptions();
    syncUiLanguageSelect();
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
    const startRuntime = function () {
      startReminderLoop();
      startCountdownLoop();
      startPushSyncLoop();
      initCloudSync();
    };
    renderAll();
    startRuntime();
  }

  function initAccessGate() {
    if (!els.passwordGate || !els.passwordForm || !els.serverInput || !els.passwordInput) {
      init();
      return;
    }
    document.body.classList.add("gate-locked");
    els.passwordGate.classList.remove("hidden");
    els.passwordForm.addEventListener("submit", handlePasswordSubmit);
    els.serverInput.addEventListener("input", clearPasswordError);
    els.passwordInput.addEventListener("input", clearPasswordError);
    setTimeout(function () {
      els.serverInput.focus();
    }, 40);
  }

  function normalizeUiLanguage(value) {
    const text = String(value || "").trim().toLowerCase();
    return text === "en" ? "en" : "zh";
  }

  function isEnglishUi() {
    return normalizeUiLanguage(state.uiLanguage) === "en";
  }

  function loadUiLanguage() {
    try {
      return normalizeUiLanguage(localStorage.getItem(UI_LANGUAGE_STORAGE_KEY) || DEFAULT_UI_LANGUAGE);
    } catch (error) {
      return DEFAULT_UI_LANGUAGE;
    }
  }

  function saveUiLanguage() {
    try {
      localStorage.setItem(UI_LANGUAGE_STORAGE_KEY, normalizeUiLanguage(state.uiLanguage));
    } catch (error) {
      // ignore storage failure in private mode
    }
  }

  function getUiText(key, params) {
    const lang = normalizeUiLanguage(state.uiLanguage);
    const table = UI_TEXT[lang] || UI_TEXT.zh;
    const fallback = UI_TEXT.zh || {};
    let value = table[key];
    if (value === undefined) {
      value = fallback[key];
    }
    if (typeof value !== "string") {
      return value;
    }
    return value.replace(/\{(\w+)\}/g, function (_, token) {
      return params && params[token] !== undefined ? String(params[token]) : "";
    });
  }

  function getCategoryDisplayName(category) {
    const key = String(category || "").trim();
    if (!key) {
      return "";
    }
    const lang = normalizeUiLanguage(state.uiLanguage);
    const map = CATEGORY_LABEL_MAP[lang];
    if (map && map[key]) {
      return map[key];
    }
    return key;
  }

  function getSubcategoryDisplayName(subcategory) {
    const key = String(subcategory || "").trim();
    if (!key) {
      return "";
    }
    const lang = normalizeUiLanguage(state.uiLanguage);
    const map = SUBCATEGORY_LABEL_MAP[lang];
    if (map && map[key]) {
      return map[key];
    }
    return key;
  }

  function setElementText(id, value) {
    const el = document.getElementById(id);
    if (!el) {
      return;
    }
    el.textContent = String(value == null ? "" : value);
  }

  function setElementPlaceholder(id, value) {
    const el = document.getElementById(id);
    if (!el) {
      return;
    }
    el.setAttribute("placeholder", String(value == null ? "" : value));
  }

  function setSelectOptionText(selectEl, value, label) {
    if (!selectEl) {
      return;
    }
    const option = Array.prototype.find.call(selectEl.options || [], function (item) {
      return String(item.value) === String(value);
    });
    if (option) {
      option.textContent = String(label || "");
    }
  }

  function syncUiLanguageSelect() {
    if (!els.uiLanguageSelect) {
      return;
    }
    els.uiLanguageSelect.value = normalizeUiLanguage(state.uiLanguage);
  }

  function applyUiLanguageToStatic() {
    const lang = normalizeUiLanguage(state.uiLanguage);
    document.documentElement.lang = lang === "en" ? "en" : "zh-Hant";
    document.title = getUiText("appTitle");
    setElementText("password-title", getUiText("appTitle"));
    setElementText("password-subtitle", getUiText("passwordSubtitle"));
    setElementText("password-server-label", getUiText("passwordServerLabel"));
    setElementText("password-label", getUiText("passwordLabel"));
    setElementText("password-submit", getUiText("passwordSubmit"));
    setElementText("ui-language-label", getUiText("languageLabel"));
    setElementText("today-section-title", getUiText("todaySectionTitle"));
    setElementText("today-checkin-label", getUiText("todayCheckinLabel"));
    setElementText("today-checkout-label", getUiText("todayCheckoutLabel"));
    setElementText("today-occ-label", getUiText("todayOccLabel"));
    setElementText("today-overview-save-btn", getUiText("save"));
    setElementText("today-pinned-title", getUiText("todayPinnedTitle") || (lang === "en" ? "Pinned" : "置頂區"));
    setElementText("today-filter-label", getUiText("todayFilterLabel"));
    setElementText("task-form-title", getUiText("taskFormTitle"));
    setElementText("task-category-label", getUiText("taskCategoryLabel"));
    setElementText("task-subcategory-label", getUiText("taskSubcategoryLabel"));
    setElementText("task-title-label", getUiText("taskTitleLabel"));
    setElementText("task-time-label", getUiText("taskTimeLabel"));
    setElementText("time-range-sep", getUiText("timeRangeSep"));
    setElementText("all-day-btn", getUiText("allDay"));
    setElementText("task-owner-label", getUiText("taskOwnerLabel"));
    setElementText("task-assignee-label", getUiText("taskAssigneeLabel"));
    setElementText("task-pin-label", getUiText("taskPinLabel"));
    setElementText("task-pin-help", getUiText("taskPinHelp"));
    setElementText("task-desc-label", getUiText("taskDescLabel"));
    setElementText("add-task-btn", state.editingTaskId ? getUiText("save") : getUiText("addTask"));
    setElementText("clear-form-btn", getUiText("clear"));
    setElementText("cancel-edit-btn", getUiText("cancelEdit"));
    setElementText("query-panel-title", getUiText("queryPanelTitle"));
    setElementText("result-panel-title", getUiText("resultPanelTitle"));
    setElementText("task-list-filter-label", getUiText("taskListFilterLabel"));
    setElementText("export-date-label", getUiText("exportDateLabel"));
    setElementText("export-status-label", getUiText("exportStatusLabel"));
    setElementText("export-word-btn", getUiText("exportWord"));
    setElementText("export-excel-btn", getUiText("exportExcel"));
    setElementText("upcoming-title", getUiText("upcomingTitle"));
    setElementText("mobile-upcoming-close", getUiText("close"));
    setElementText("notification-toggle-label", getUiText("notificationToggleLabel"));
    setElementText("request-notification-btn", getUiText("requestNotification"));
    setElementPlaceholder("task-title", getUiText("taskTitlePlaceholder"));
    setElementPlaceholder("task-owner", getUiText("taskOwnerPlaceholder"));
    setElementPlaceholder("task-description", getUiText("taskDescriptionPlaceholder"));
    setElementPlaceholder("query-keyword", getUiText("keywordPlaceholder"));
    setElementPlaceholder("server-input", getUiText("passwordServerPlaceholder"));
    setElementPlaceholder("password-input", getUiText("passwordPlaceholder"));
    setElementPlaceholder("today-occupancy-rate", lang === "en" ? "e.g. 99.47" : "例: 99.47");
    refreshAssigneeSelectOptions();
    setElementText("mobile-upcoming-toggle-text", getUiText("mobileUpcomingLabel"));
    if (els.mobileAddBtn) {
      els.mobileAddBtn.setAttribute(
        "aria-label",
        lang === "en" ? "Jump to add handover task" : "前往新增交接待辦",
      );
    }
    if (els.mobileTopBtn) {
      els.mobileTopBtn.setAttribute("aria-label", lang === "en" ? "Back to top" : "回到頂部");
    }
    if (els.tableHeaderCells && els.tableHeaderCells.length > 0) {
      const labels = getUiText("tableHeaders");
      if (Array.isArray(labels)) {
        els.tableHeaderCells.forEach(function (cell, index) {
          if (labels[index]) {
            cell.textContent = labels[index];
          }
        });
      }
    }
    if (els.queryStatus) {
      setSelectOptionText(els.queryStatus, "all", getUiText("allStatus"));
      setSelectOptionText(els.queryStatus, "pending", getUiText("pending"));
      setSelectOptionText(els.queryStatus, "done", getUiText("done"));
      els.queryStatus.setAttribute("aria-label", lang === "en" ? "Status filter" : "狀態篩選");
    }
    if (els.exportStatus) {
      setSelectOptionText(els.exportStatus, "all", getUiText("allStatus"));
      setSelectOptionText(els.exportStatus, "pending", getUiText("pending"));
      setSelectOptionText(els.exportStatus, "done", getUiText("done"));
    }
    if (els.taskListFilter) {
      setSelectOptionText(els.taskListFilter, "all", getUiText("all"));
      setSelectOptionText(els.taskListFilter, "pending", getUiText("pending"));
      setSelectOptionText(els.taskListFilter, "done", getUiText("done"));
      setSelectOptionText(els.taskListFilter, "pinned", getUiText("pinned"));
      els.taskListFilter.setAttribute("aria-label", lang === "en" ? "Result filter" : "查詢結果篩選");
    }
    if (els.queryCategory) {
      els.queryCategory.setAttribute("aria-label", lang === "en" ? "Category filter" : "主分類查詢");
    }
    if (els.searchBtn) {
      els.searchBtn.textContent = getUiText("search");
    }
    if (els.clearSearchBtn) {
      els.clearSearchBtn.textContent = getUiText("clearSearch");
    }
    if (els.todayDateLabel) {
      els.todayDateLabel.innerHTML = getUiText("todayDateLabel") + "：<strong id=\"today-auto-date\">-</strong>";
      els.todayAutoDate = document.getElementById("today-auto-date");
    }
    if (els.panelToggleButtons && els.panelToggleButtons.length > 0) {
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
    syncUiLanguageSelect();
  }

  function normalizeServerInput(value) {
    return String(value || "").trim().toLowerCase();
  }

  function normalizeAccountName(value) {
    return String(value || "")
      .trim()
      .toLowerCase()
      .replace(/[^a-z0-9_-]/g, "");
  }

  function cloneCategoryList(source) {
    const defaults = CATEGORIES.slice();
    if (!Array.isArray(source) || source.length === 0) {
      return defaults;
    }
    const seen = new Set();
    const list = [];
    source.forEach(function (item) {
      const value = String(item || "").trim();
      if (!value || seen.has(value)) {
        return;
      }
      seen.add(value);
      list.push(value);
    });
    return list.length > 0 ? list : defaults;
  }

  function cloneSubcategoryMap(source, categories) {
    const base = source && typeof source === "object" ? source : {};
    const map = {};
    (Array.isArray(categories) ? categories : []).forEach(function (category) {
      const raw = Array.isArray(base[category]) ? base[category] : Array.isArray(SUBCATEGORY_MAP[category]) ? SUBCATEGORY_MAP[category] : [];
      const list = [];
      raw.forEach(function (item) {
        const value = String(item || "").trim();
        if (!value || list.indexOf(value) !== -1) {
          return;
        }
        list.push(value);
      });
      map[category] = list;
    });
    return map;
  }

  function cloneExportColorMap(source) {
    const map = {};
    Object.keys(EXPORT_CATEGORY_COLORS).forEach(function (key) {
      const value = String(EXPORT_CATEGORY_COLORS[key] || "").trim();
      map[key] = value || EXPORT_DEFAULT_COLOR;
    });
    if (source && typeof source === "object") {
      Object.keys(source).forEach(function (key) {
        const value = String(source[key] || "").trim();
        if (!value) {
          return;
        }
        map[key] = value;
      });
    }
    return map;
  }

  function normalizeThemeClass(value) {
    const text = String(value || "").trim();
    if (!text) {
      return "";
    }
    return /^theme-[a-z0-9_-]+$/i.test(text) ? text : "";
  }

  function sanitizeThemeVars(source) {
    if (!source || typeof source !== "object") {
      return {};
    }
    const vars = {};
    Object.keys(source).forEach(function (key) {
      const k = String(key || "").trim();
      const value = String(source[key] || "").trim();
      if (!/^--[a-z0-9_-]+$/i.test(k) || !value) {
        return;
      }
      vars[k] = value;
    });
    return vars;
  }

  function resolveProfileForServer(serverId) {
    const key = normalizeServerInput(serverId || DEFAULT_SERVER_ID);
    const raw = USER_PROFILE_MAP[key] || USER_PROFILE_MAP[DEFAULT_SERVER_ID] || {};
    const categories = cloneCategoryList(raw.categories);
    return {
      name: String(raw.name || key || DEFAULT_SERVER_ID),
      themeClass: normalizeThemeClass(raw.themeClass),
      themeVars: sanitizeThemeVars(raw.themeVars),
      categories: categories,
      subcategoryMap: cloneSubcategoryMap(raw.subcategoryMap, categories),
      exportCategoryColors: cloneExportColorMap(raw.exportCategoryColors),
    };
  }

  function getActiveProfile() {
    if (!state.currentProfile || typeof state.currentProfile !== "object") {
      state.currentProfile = resolveProfileForServer(state.currentServerId || DEFAULT_SERVER_ID);
    }
    return state.currentProfile;
  }

  function getActiveCategories() {
    return getActiveProfile().categories.slice();
  }

  function getActiveSubcategoryMap() {
    const map = getActiveProfile().subcategoryMap;
    return map && typeof map === "object" ? map : {};
  }

  function getActiveExportColorMap() {
    const map = getActiveProfile().exportCategoryColors;
    return map && typeof map === "object" ? map : {};
  }

  function applyProfileTheme(profile) {
    const root = document.documentElement;
    const body = document.body;
    const safeProfile = profile && typeof profile === "object" ? profile : {};
    const prevKeys = Array.isArray(state.appliedThemeVarKeys) ? state.appliedThemeVarKeys : [];
    prevKeys.forEach(function (key) {
      root.style.removeProperty(key);
    });
    const vars = safeProfile.themeVars && typeof safeProfile.themeVars === "object" ? safeProfile.themeVars : {};
    const nextKeys = [];
    Object.keys(vars).forEach(function (key) {
      root.style.setProperty(key, vars[key]);
      nextKeys.push(key);
    });
    state.appliedThemeVarKeys = nextKeys;
    body.classList.remove("theme-mono");
    if (safeProfile.themeClass) {
      body.classList.add(safeProfile.themeClass);
    }
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

  function buildCloudStateUrl(baseUrl, serverId, apiVersion) {
    const base = normalizeCloudApiBase(baseUrl);
    if (!base || !serverId) {
      return "";
    }
    const version = String(apiVersion || "v1").toLowerCase() === "v2" ? "v2" : "v1";
    return base + "/" + version + "/state/" + encodeURIComponent(serverId);
  }

  function buildCloudAuthLoginUrl(baseUrl) {
    const base = normalizeCloudApiBase(baseUrl);
    if (!base) {
      return "";
    }
    return base + "/v2/auth/login";
  }

  function shouldUseV2StateApi(serverId) {
    const normalizedServerId = normalizeServerInput(serverId || "");
    if (!normalizedServerId) {
      return false;
    }
    if (!state.authToken || !state.authSession || typeof state.authSession !== "object") {
      return false;
    }
    return normalizeServerInput(state.authSession.serverId || "") === normalizedServerId;
  }

  function getCurrentCloudStateApiVersion() {
    const currentServerId = state.currentServerId || DEFAULT_SERVER_ID;
    return shouldUseV2StateApi(currentServerId) ? "v2" : "v1";
  }

  function getCurrentCloudUrl() {
    const server = getCurrentServerConfig();
    if (!server) {
      return "";
    }
    const cloudBase = getCloudApiBase(server);
    return buildCloudStateUrl(cloudBase, server.serverId, getCurrentCloudStateApiVersion());
  }

  function getCurrentCloudAuthHeaders() {
    if (getCurrentCloudStateApiVersion() !== "v2") {
      return {};
    }
    if (!state.authToken) {
      return {};
    }
    return {
      Authorization: "Bearer " + state.authToken,
    };
  }

  function getCurrentAuthUsername() {
    return normalizeAccountName(state.authSession && state.authSession.username ? state.authSession.username : "");
  }

  function getCurrentAuthRole() {
    return String(state.authSession && state.authSession.role ? state.authSession.role : "")
      .trim()
      .toLowerCase();
  }

  function canAssignToOthers() {
    const role = getCurrentAuthRole();
    return role === "admin" || role === "manager";
  }

  function findAccountByUsername(username) {
    const normalized = normalizeAccountName(username);
    if (!normalized) {
      return null;
    }
    return state.assignableAccounts.find(function (account) {
      return normalizeAccountName(account.username) === normalized;
    }) || null;
  }

  function refreshAssigneeSelectOptions() {
    if (!els.taskAssignee) {
      return;
    }
    const currentValue = normalizeAccountName(els.taskAssignee.value || "");
    const options = [];
    options.push('<option value="">' + escapeHtml(getUiText("taskAssigneeUnassigned")) + "</option>");

    const seen = {};
    const list = Array.isArray(state.assignableAccounts) ? state.assignableAccounts : [];
    list.forEach(function (account) {
      const username = normalizeAccountName(account && account.username ? account.username : "");
      if (!username || seen[username]) {
        return;
      }
      seen[username] = true;
      const role = String(account && account.role ? account.role : "").trim();
      const label = role ? username + " (" + role + ")" : username;
      options.push('<option value="' + escapeHtml(username) + '">' + escapeHtml(label) + "</option>");
    });

    els.taskAssignee.innerHTML = options.join("");

    const currentUser = getCurrentAuthUsername();
    if (!canAssignToOthers() && currentUser) {
      if (!findAccountByUsername(currentUser)) {
        const role = getCurrentAuthRole();
        const roleText = role ? " (" + role + ")" : "";
        const fallbackOption = document.createElement("option");
        fallbackOption.value = currentUser;
        fallbackOption.textContent = currentUser + roleText;
        els.taskAssignee.appendChild(fallbackOption);
      }
      els.taskAssignee.value = currentUser;
      els.taskAssignee.disabled = true;
      return;
    }

    els.taskAssignee.disabled = false;
    if (currentValue && findAccountByUsername(currentValue)) {
      els.taskAssignee.value = currentValue;
      return;
    }
    els.taskAssignee.value = "";
  }

  async function loadAssignableAccounts(server) {
    const currentServer = server || getCurrentServerConfig();
    if (!currentServer) {
      state.assignableAccounts = [];
      refreshAssigneeSelectOptions();
      return;
    }

    const serverId = normalizeServerInput(currentServer.serverId || "");
    if (serverId !== "test" || !state.authToken) {
      state.assignableAccounts = [];
      refreshAssigneeSelectOptions();
      return;
    }

    const cloudBase = getCloudApiBase(currentServer);
    if (!cloudBase) {
      state.assignableAccounts = [];
      refreshAssigneeSelectOptions();
      return;
    }

    const url = cloudBase + "/v2/accounts/" + encodeURIComponent(serverId);
    try {
      const response = await fetchWithTimeout(
        url,
        {
          method: "GET",
          cache: "no-store",
          headers: getCurrentCloudAuthHeaders(),
        },
        CLOUD_REQUEST_TIMEOUT_MS,
      );
      if (!response.ok) {
        throw new Error("account list failed: " + response.status);
      }
      const parsed = await response.json();
      const list = parsed && Array.isArray(parsed.accounts) ? parsed.accounts : [];
      state.assignableAccounts = list
        .map(function (item) {
          return {
            username: normalizeAccountName(item && item.username ? item.username : ""),
            role: String(item && item.role ? item.role : "").trim().toLowerCase(),
          };
        })
        .filter(function (item) {
          return Boolean(item.username);
        });
      refreshAssigneeSelectOptions();
    } catch (error) {
      console.error("loadAssignableAccounts error", error);
      state.assignableAccounts = [];
      refreshAssigneeSelectOptions();
    }
  }

  function isPushSupported() {
    return (
      typeof window !== "undefined" &&
      typeof navigator !== "undefined" &&
      "serviceWorker" in navigator &&
      "PushManager" in window &&
      "Notification" in window
    );
  }

  function canUseCloudPush(server) {
    const currentServer = server || getCurrentServerConfig();
    if (!currentServer) {
      return false;
    }
    const serverId = normalizeServerInput(currentServer.serverId || "");
    if (serverId !== "test") {
      return false;
    }
    if (!state.authToken) {
      return false;
    }
    const cloudBase = getCloudApiBase(currentServer);
    return Boolean(cloudBase);
  }

  function getAppBasePath() {
    const path = String(window.location.pathname || "/");
    if (!path) {
      return "/";
    }
    if (path.endsWith("/")) {
      return path;
    }
    const slashIndex = path.lastIndexOf("/");
    if (slashIndex < 0) {
      return "/";
    }
    return path.slice(0, slashIndex + 1) || "/";
  }

  function buildCloudPushUrl(baseUrl, serverId, action) {
    const base = normalizeCloudApiBase(baseUrl);
    const normalizedServerId = normalizeServerInput(serverId || "");
    const normalizedAction = String(action || "").trim().toLowerCase();
    if (!base || !normalizedServerId || (normalizedAction !== "subscribe" && normalizedAction !== "unsubscribe")) {
      return "";
    }
    return base + "/v2/push/" + encodeURIComponent(normalizedServerId) + "/" + normalizedAction;
  }

  function buildCloudPushPublicKeyUrl(baseUrl) {
    const base = normalizeCloudApiBase(baseUrl);
    if (!base) {
      return "";
    }
    return base + "/v2/push/public-key";
  }

  function toAbsoluteUrl(input) {
    try {
      return new URL(String(input || ""), window.location.href).toString();
    } catch (error) {
      return "";
    }
  }

  function buildServiceWorkerScriptUrl(server) {
    const currentServer = server || getCurrentServerConfig();
    const cloudBase = normalizeCloudApiBase(getCloudApiBase(currentServer));
    const appBasePath = getAppBasePath();
    const params = new URLSearchParams();
    if (cloudBase) {
      params.set("cloudBase", cloudBase);
    }
    params.set("appBasePath", appBasePath);
    params.set("v", BACKUP_VERSION);
    const scriptPath = PUSH_SERVICE_WORKER_FILE + "?" + params.toString();
    return toAbsoluteUrl(scriptPath);
  }

  function postServiceWorkerConfig(registration, server) {
    if (!registration) {
      return;
    }
    const currentServer = server || getCurrentServerConfig();
    const payload = {
      type: "HANDOVER_SW_CONFIG",
      cloudBase: normalizeCloudApiBase(getCloudApiBase(currentServer)),
      appBasePath: getAppBasePath(),
    };
    const target = registration.active || registration.waiting || registration.installing;
    if (!target || typeof target.postMessage !== "function") {
      return;
    }
    try {
      target.postMessage(payload);
    } catch (error) {
      // ignore unavailable service worker messaging
    }
  }

  function urlBase64ToUint8Array(base64String) {
    const value = String(base64String || "").trim();
    if (!value) {
      return null;
    }
    const padding = "=".repeat((4 - (value.length % 4)) % 4);
    const base64 = (value + padding).replace(/-/g, "+").replace(/_/g, "/");
    const rawData = window.atob(base64);
    const outputArray = new Uint8Array(rawData.length);
    for (let i = 0; i < rawData.length; i += 1) {
      outputArray[i] = rawData.charCodeAt(i);
    }
    return outputArray;
  }

  async function ensureServiceWorkerRegistration(server) {
    if (!isPushSupported()) {
      return null;
    }
    const scriptUrl = buildServiceWorkerScriptUrl(server);
    if (!scriptUrl) {
      return null;
    }
    const scope = getAppBasePath();
    const registration = await navigator.serviceWorker.register(scriptUrl, { scope: scope });
    state.serviceWorkerRegistration = registration;
    postServiceWorkerConfig(registration, server);
    navigator.serviceWorker.ready
      .then(function (readyRegistration) {
        postServiceWorkerConfig(readyRegistration || registration, server);
      })
      .catch(function () {
        return null;
      });
    return registration;
  }

  async function fetchCloudPushPublicKey(server) {
    const currentServer = server || getCurrentServerConfig();
    if (!currentServer) {
      return "";
    }
    const cloudBase = getCloudApiBase(currentServer);
    const url = buildCloudPushPublicKeyUrl(cloudBase);
    if (!url) {
      return "";
    }
    try {
      const response = await fetchWithTimeout(
        url,
        {
          method: "GET",
          cache: "no-store",
        },
        CLOUD_REQUEST_TIMEOUT_MS,
      );
      if (!response.ok) {
        return "";
      }
      const parsed = await response.json();
      const key = parsed && parsed.publicKey ? String(parsed.publicKey).trim() : "";
      return key;
    } catch (error) {
      console.error("fetchCloudPushPublicKey error", error);
      return "";
    }
  }

  async function postPushSubscription(server, subscription) {
    const currentServer = server || getCurrentServerConfig();
    if (!currentServer || !subscription) {
      return false;
    }
    const url = buildCloudPushUrl(getCloudApiBase(currentServer), currentServer.serverId, "subscribe");
    if (!url) {
      return false;
    }
    const response = await fetchWithTimeout(
      url,
      {
        method: "POST",
        cache: "no-store",
        headers: Object.assign(
          {
            "Content-Type": "application/json",
          },
          getCurrentCloudAuthHeaders(),
        ),
        body: JSON.stringify({ subscription: subscription }),
      },
      CLOUD_REQUEST_TIMEOUT_MS,
    );
    return response.ok;
  }

  async function postPushUnsubscribe(server, endpoint) {
    const currentServer = server || getCurrentServerConfig();
    if (!currentServer || !endpoint) {
      return false;
    }
    const url = buildCloudPushUrl(getCloudApiBase(currentServer), currentServer.serverId, "unsubscribe");
    if (!url) {
      return false;
    }
    const response = await fetchWithTimeout(
      url,
      {
        method: "POST",
        cache: "no-store",
        headers: Object.assign(
          {
            "Content-Type": "application/json",
          },
          getCurrentCloudAuthHeaders(),
        ),
        body: JSON.stringify({ endpoint: endpoint }),
      },
      CLOUD_REQUEST_TIMEOUT_MS,
    );
    return response.ok;
  }

  async function disableCloudPushSubscription(server) {
    if (!isPushSupported()) {
      return;
    }
    try {
      const registration = state.serviceWorkerRegistration || (await navigator.serviceWorker.getRegistration(getAppBasePath()));
      if (!registration || !registration.pushManager) {
        return;
      }
      const subscription = await registration.pushManager.getSubscription();
      if (!subscription) {
        return;
      }
      const endpoint = subscription.endpoint ? String(subscription.endpoint) : "";
      if (endpoint && canUseCloudPush(server || getCurrentServerConfig())) {
        try {
          await postPushUnsubscribe(server, endpoint);
        } catch (error) {
          console.error("postPushUnsubscribe error", error);
        }
      }
      await subscription.unsubscribe();
    } catch (error) {
      console.error("disableCloudPushSubscription error", error);
    }
  }

  async function ensureCloudPushSubscription(server, options) {
    const opts = options && typeof options === "object" ? options : {};
    if (!isPushSupported()) {
      return { ok: false, reason: "UNSUPPORTED" };
    }
    if (!canUseCloudPush(server)) {
      return { ok: false, reason: "NOT_ELIGIBLE" };
    }
    if (!els.notificationToggle || !els.notificationToggle.checked) {
      return { ok: false, reason: "TOGGLE_OFF" };
    }

    const registration = await ensureServiceWorkerRegistration(server);
    if (!registration || !registration.pushManager) {
      return { ok: false, reason: "REGISTER_FAILED" };
    }

    let permission = Notification.permission;
    if (permission !== "granted" && opts.requestPermission !== false) {
      permission = await Notification.requestPermission();
    }
    if (permission !== "granted") {
      return { ok: false, reason: "PERMISSION_NOT_GRANTED" };
    }

    let subscription = await registration.pushManager.getSubscription();
    if (!subscription) {
      const publicKey = await fetchCloudPushPublicKey(server);
      const keyBytes = urlBase64ToUint8Array(publicKey);
      if (!keyBytes) {
        return { ok: false, reason: "NO_PUBLIC_KEY" };
      }
      subscription = await registration.pushManager.subscribe({
        userVisibleOnly: true,
        applicationServerKey: keyBytes,
      });
    }

    try {
      const ok = await postPushSubscription(server, subscription.toJSON ? subscription.toJSON() : subscription);
      if (!ok) {
        return { ok: false, reason: "SUBSCRIBE_API_FAILED" };
      }
    } catch (error) {
      console.error("postPushSubscription error", error);
      return { ok: false, reason: "SUBSCRIBE_API_FAILED" };
    }

    return { ok: true };
  }

  function startPushSyncLoop() {
    if (state.pushSyncTimer) {
      clearInterval(state.pushSyncTimer);
      state.pushSyncTimer = null;
    }
    if (!isPushSupported()) {
      return;
    }
    state.pushSyncTimer = setInterval(function () {
      if (!document.hidden) {
        ensureCloudPushSubscription(getCurrentServerConfig(), {
          requestPermission: false,
        }).catch(function () {
          return null;
        });
      }
    }, PUSH_SYNC_INTERVAL_MS);
  }

  function buildCloudTranslateUrl(baseUrl, serverId) {
    const base = normalizeCloudApiBase(baseUrl);
    if (!base || !serverId) {
      return "";
    }
    return base + "/v1/translate/" + encodeURIComponent(serverId);
  }

  function getCurrentCloudTranslateUrl() {
    const server = getCurrentServerConfig();
    if (!server) {
      return "";
    }
    const cloudBase = getCloudApiBase(server);
    return buildCloudTranslateUrl(cloudBase, server.serverId);
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
      disableCloudPushSubscription(getCurrentServerConfig()).catch(function () {
        return null;
      });
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
      ensureCloudPushSubscription(getCurrentServerConfig(), {
        requestPermission: false,
      }).catch(function () {
        return null;
      });
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
    const response = await fetchWithTimeout(
      url,
      { method: "GET", cache: "no-store", headers: getCurrentCloudAuthHeaders() },
      CLOUD_REQUEST_TIMEOUT_MS,
    );
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

      const existingResponse = await fetchWithTimeout(
        url,
        { method: "GET", cache: "no-store", headers: getCurrentCloudAuthHeaders() },
        CLOUD_REQUEST_TIMEOUT_MS,
      );
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
          headers: Object.assign({ "Content-Type": "application/json" }, getCurrentCloudAuthHeaders()),
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

  function normalizeTaskTranslations(raw) {
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
      return {};
    }
    const result = {};
    ["en", "zh"].forEach(function (lang) {
      const item = raw[lang];
      if (!item || typeof item !== "object" || Array.isArray(item)) {
        return;
      }
      result[lang] = {
        title: String(item.title || "").trim(),
        description: String(item.description || "").trim(),
        subcategory: String(item.subcategory || "").trim(),
        _sig: String(item._sig || "").trim(),
      };
    });
    return result;
  }

  function getTaskTranslationSignature(task) {
    if (!task || typeof task !== "object") {
      return "";
    }
    return "desc::" + String(task.description || "");
  }

  function hasCjkText(text) {
    return /[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]/.test(String(text || ""));
  }

  function hasLatinText(text) {
    return /[a-z]/i.test(String(text || ""));
  }

  function shouldTranslateTextForTarget(text, targetLang) {
    const source = String(text || "").trim();
    if (!source) {
      return false;
    }
    const target = normalizeUiLanguage(targetLang);
    if (target === "en") {
      return hasCjkText(source);
    }
    if (target === "zh") {
      return hasLatinText(source) && !hasCjkText(source);
    }
    return false;
  }

  function getTaskTranslatedField(task, field, lang) {
    if (!task || !field) {
      return "";
    }
    if (field !== "description") {
      return "";
    }
    const normalizedLang = normalizeUiLanguage(lang);
    const translations = task.translations && typeof task.translations === "object" ? task.translations : null;
    if (!translations) {
      return "";
    }
    const data = translations[normalizedLang];
    if (!data || typeof data !== "object") {
      return "";
    }
    if (String(data._sig || "") !== getTaskTranslationSignature(task)) {
      return "";
    }
    return String(data[field] || "").trim();
  }

  function getTaskDisplayField(task, field) {
    if (!task || !field) {
      return "";
    }
    const original = String(task[field] || "").trim();
    const translated = getTaskTranslatedField(task, field, state.uiLanguage);
    return translated || original;
  }

  function isTaskTranslationUsable(task, translationEntry, targetLang) {
    if (!task || typeof task !== "object") {
      return false;
    }
    if (!translationEntry || typeof translationEntry !== "object" || Array.isArray(translationEntry)) {
      return false;
    }
    const target = normalizeUiLanguage(targetLang);
    const source = String(task.description || "").trim();
    if (!source || !shouldTranslateTextForTarget(source, target)) {
      return true;
    }
    const translated = String(translationEntry.description || "").trim();
    if (!translated) {
      return false;
    }
    // If translated note is exactly same as source when translation is required,
    // treat it as stale/failed cache and retry.
    if (translated === source) {
      return false;
    }
    return true;
  }

  function hasTranslatedDescriptionForCurrentLanguage(task) {
    if (!task || typeof task !== "object") {
      return false;
    }
    const original = String(task.description || "").trim();
    if (!original) {
      return false;
    }
    const translated = getTaskTranslatedField(task, "description", state.uiLanguage);
    if (!translated) {
      return false;
    }
    return translated !== original;
  }

  function isTaskShowingOriginalDescription(task) {
    if (!task || !task.id) {
      return false;
    }
    if (!hasTranslatedDescriptionForCurrentLanguage(task)) {
      return false;
    }
    return Boolean(state.showOriginalByTaskId && state.showOriginalByTaskId[task.id]);
  }

  function getTaskDisplayDescription(task) {
    if (!task || typeof task !== "object") {
      return "";
    }
    const original = String(task.description || "").trim();
    if (!original) {
      return "";
    }
    if (isTaskShowingOriginalDescription(task)) {
      return original;
    }
    const translated = getTaskTranslatedField(task, "description", state.uiLanguage);
    return translated || original;
  }

  function toggleTaskOriginalDescription(task) {
    if (!task || !task.id) {
      return;
    }
    if (!hasTranslatedDescriptionForCurrentLanguage(task)) {
      showToast(getUiText("noTranslatedContent"));
      return;
    }
    if (!state.showOriginalByTaskId || typeof state.showOriginalByTaskId !== "object") {
      state.showOriginalByTaskId = {};
    }
    if (state.showOriginalByTaskId[task.id]) {
      delete state.showOriginalByTaskId[task.id];
    } else {
      state.showOriginalByTaskId[task.id] = true;
    }
    renderAll();
  }

  async function requestCloudTranslations(targetLang, texts) {
    const target = normalizeUiLanguage(targetLang);
    const sourceList = (Array.isArray(texts) ? texts : [])
      .map(function (item) {
        return String(item || "");
      })
      .filter(function (item) {
        return item.length > 0;
      });
    if (sourceList.length === 0) {
      return [];
    }
    const url = getCurrentCloudTranslateUrl();
    if (!url) {
      throw new Error("translate url missing");
    }
    let response = null;
    let lastStatus = 0;
    const payload = JSON.stringify({
      targetLang: target,
      texts: sourceList,
    });
    for (let attempt = 1; attempt <= 3; attempt += 1) {
      response = await fetchWithTimeout(
        url,
        {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: payload,
        },
        CLOUD_REQUEST_TIMEOUT_MS + 6000,
      );
      if (response.ok) {
        break;
      }
      lastStatus = Number(response.status || 0);
      if (!isRetryableTranslateStatus(lastStatus) || attempt >= 3) {
        break;
      }
      await waitMs(300 * attempt * attempt);
    }
    if (!response || !response.ok) {
      throw new Error("translate failed: " + String(lastStatus || (response && response.status) || 0));
    }
    const parsed = await response.json();
    if (!parsed || !Array.isArray(parsed.translations)) {
      throw new Error("translate invalid response");
    }
    const result = parsed.translations.map(function (item) {
      return String(item == null ? "" : item).trim();
    });
    if (result.length !== sourceList.length) {
      throw new Error("translate response length mismatch");
    }
    return result;
  }

  function isRetryableTranslateStatus(status) {
    const code = Number(status || 0);
    return code === 429 || code === 500 || code === 502 || code === 503 || code === 504;
  }

  function waitMs(ms) {
    const delay = Number(ms) > 0 ? Number(ms) : 0;
    return new Promise(function (resolve) {
      setTimeout(resolve, delay);
    });
  }

  function ensureTaskTranslationContainer(task) {
    if (!task || typeof task !== "object") {
      return;
    }
    if (!task.translations || typeof task.translations !== "object" || Array.isArray(task.translations)) {
      task.translations = {};
    }
  }

  function setTaskTranslationEntry(task, targetLang, descriptionValue) {
    if (!task || typeof task !== "object") {
      return;
    }
    const target = normalizeUiLanguage(targetLang);
    ensureTaskTranslationContainer(task);
    task.translations[target] = {
      title: String(task.title || "").trim(),
      description: String(descriptionValue == null ? task.description : descriptionValue).trim(),
      subcategory: String(task.subcategory || "").trim(),
      _sig: getTaskTranslationSignature(task),
    };
  }

  function hasTaskTranslationCacheForLanguage(task, targetLang) {
    if (!task || typeof task !== "object") {
      return false;
    }
    const target = normalizeUiLanguage(targetLang);
    const entry =
      task.translations && typeof task.translations === "object" && !Array.isArray(task.translations)
        ? task.translations[target]
        : null;
    if (!entry || typeof entry !== "object") {
      return false;
    }
    if (String(entry._sig || "") !== getTaskTranslationSignature(task)) {
      return false;
    }
    return isTaskTranslationUsable(task, entry, target);
  }

  function canTaskTranslateForLanguage(task, targetLang) {
    if (!task || typeof task !== "object") {
      return false;
    }
    const source = String(task.description || "").trim();
    if (!source) {
      return false;
    }
    return shouldTranslateTextForTarget(source, normalizeUiLanguage(targetLang));
  }

  function needsTaskManualTranslationForLanguage(task, targetLang) {
    return canTaskTranslateForLanguage(task, targetLang) && !hasTaskTranslationCacheForLanguage(task, targetLang);
  }

  function isTaskTranslationInProgress(taskId) {
    const id = String(taskId || "").trim();
    if (!id) {
      return false;
    }
    return Boolean(state.translatingTaskIds && state.translatingTaskIds[id]);
  }

  function setTaskTranslationInProgress(taskId, inProgress) {
    const id = String(taskId || "").trim();
    if (!id) {
      return;
    }
    if (!state.translatingTaskIds || typeof state.translatingTaskIds !== "object") {
      state.translatingTaskIds = {};
    }
    if (inProgress) {
      state.translatingTaskIds[id] = true;
      return;
    }
    delete state.translatingTaskIds[id];
  }

  async function translateTaskDescriptionForLanguage(task, targetLang, options) {
    if (!task || typeof task !== "object") {
      return false;
    }
    const opts = options && typeof options === "object" ? options : {};
    const target = normalizeUiLanguage(targetLang);
    const source = String(task.description || "").trim();

    if (!source) {
      setTaskTranslationEntry(task, target, "");
      return true;
    }

    if (!shouldTranslateTextForTarget(source, target)) {
      setTaskTranslationEntry(task, target, source);
      return true;
    }

    if (!opts.force && hasTaskTranslationCacheForLanguage(task, target)) {
      return false;
    }

    const translatedList = await requestCloudTranslations(target, [source]);
    const translated = String((translatedList && translatedList[0]) || "").trim();
    if (!translated || translated === source) {
      throw new Error("TRANSLATION_EMPTY_OR_UNCHANGED");
    }
    setTaskTranslationEntry(task, target, translated);
    return true;
  }

  async function cacheTaskTranslationsOnSave(task) {
    if (!task || typeof task !== "object") {
      return { changed: false, hasError: false };
    }
    let changed = false;
    let hasError = false;
    const targets = ["zh", "en"];

    for (const target of targets) {
      try {
        const updated = await translateTaskDescriptionForLanguage(task, target, { force: true });
        if (updated) {
          changed = true;
        }
      } catch (error) {
        hasError = true;
        console.error("cacheTaskTranslationsOnSave error", target, error);
      }
    }

    return { changed, hasError };
  }

  async function handleTranslateSingleTask(task) {
    if (!task || !task.id) {
      return;
    }
    const target = normalizeUiLanguage(state.uiLanguage);
    if (!needsTaskManualTranslationForLanguage(task, target)) {
      showToast(
        target === "en" ? "This item has no pending translation request." : "此筆目前不需要手動翻譯。",
      );
      return;
    }
    if (isTaskTranslationInProgress(task.id)) {
      return;
    }

    setTaskTranslationInProgress(task.id, true);
    renderAll();

    try {
      await translateTaskDescriptionForLanguage(task, target, { force: true });
      touchTask(task);
      saveTasks();
      renderAll();
      showToast(target === "en" ? "Translated and cached." : "已翻譯並寫入快取。");
    } catch (error) {
      console.error("handleTranslateSingleTask error", error);
      showToast(
        target === "en"
          ? "Translation failed. Try again later."
          : "翻譯失敗，請稍後再試。",
      );
    } finally {
      setTaskTranslationInProgress(task.id, false);
      renderAll();
    }
  }

  async function handlePasswordSubmit(event) {
    event.preventDefault();
    const entered = normalizeServerInput((els.serverInput && els.serverInput.value) || "");
    const server = resolveServerByInput(entered);
    if (!server) {
      showPasswordError(
        isEnglishUi()
          ? "Invalid server. Use letters/numbers/-/_ and at least 3 characters."
          : "伺服器名稱錯誤（請用英數、-、_，至少 3 碼）。",
      );
      return;
    }

    const password = String((els.passwordInput && els.passwordInput.value) || "");
    const requiresPassword = normalizeServerInput(server.serverId || "") === "test";

    if (requiresPassword && !password.trim()) {
      showPasswordError(isEnglishUi() ? "Password required for test server." : "test 伺服器需要輸入密碼。");
      return;
    }

    try {
      if (requiresPassword) {
        const loginResult = await authenticateServerWithPassword(server, password);
        state.authToken = String(loginResult.token || "");
        state.authSession = {
          serverId: String(loginResult.serverId || server.serverId || ""),
          username: String(loginResult.username || ""),
          role: String(loginResult.role || ""),
          expiresAt: String(loginResult.expiresAt || ""),
        };
      } else {
        state.authToken = "";
        state.authSession = null;
      }
      await loadAssignableAccounts(server);
      if (!requiresPassword) {
        await disableCloudPushSubscription(server);
      }
    } catch (error) {
      showPasswordError(extractAuthErrorMessage(error));
      return;
    }

    state.currentServerId = server.serverId;
    state.currentServerConfig = server;
    state.currentProfile = resolveProfileForServer(server.serverId);
    unlockAccessGate();
    init();
    if (requiresPassword) {
      ensureCloudPushSubscription(server, { requestPermission: false }).catch(function () {
        return null;
      });
    }

    if (requiresPassword && state.authSession && state.authSession.username) {
      showToast(
        isEnglishUi()
          ? "Signed in: " + state.authSession.username + " @ " + server.displayName
          : "已登入：" + state.authSession.username + "（" + server.displayName + "）",
      );
      return;
    }

    showToast(isEnglishUi() ? "Entered server: " + server.displayName : "已進入 " + server.displayName + " 伺服器。");
  }

  async function authenticateServerWithPassword(server, password) {
    const cloudBase = getCloudApiBase(server);
    const loginUrl = buildCloudAuthLoginUrl(cloudBase);
    if (!loginUrl) {
      throw new Error(isEnglishUi() ? "Cloud API is not configured." : "尚未設定 Cloud API。");
    }

    const response = await fetchWithTimeout(
      loginUrl,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          serverId: server && server.serverId ? server.serverId : "",
          password: String(password || ""),
        }),
      },
      CLOUD_REQUEST_TIMEOUT_MS,
    );

    let parsed = null;
    try {
      parsed = await response.json();
    } catch (error) {
      parsed = null;
    }

    if (!response.ok || !parsed || !parsed.ok || !parsed.token) {
      const apiError = parsed && parsed.error ? String(parsed.error) : "LOGIN_FAILED";
      throw new Error(apiError);
    }

    return parsed;
  }

  function extractAuthErrorMessage(error) {
    const raw = String(error && error.message ? error.message : "LOGIN_FAILED");
    const code = raw.toUpperCase();
    if (code.includes("INVALID_CREDENTIALS")) {
      return isEnglishUi() ? "Wrong password." : "密碼錯誤。";
    }
    if (code.includes("FORBIDDEN_SERVER")) {
      return isEnglishUi() ? "Server is not allowed." : "此伺服器未開放。";
    }
    if (code.includes("INVALID_SERVER_ID")) {
      return isEnglishUi() ? "Invalid server id." : "伺服器代號錯誤。";
    }
    if (code.includes("ACCOUNT_DISABLED")) {
      return isEnglishUi() ? "Account is disabled." : "此帳號已停用。";
    }
    return isEnglishUi() ? "Login failed, please retry." : "登入失敗，請稍後重試。";
  }

  function unlockAccessGate() {
    document.body.classList.remove("gate-locked");
    if (els.passwordGate) {
      els.passwordGate.classList.add("hidden");
    }
    if (els.serverInput) {
      els.serverInput.value = "";
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
    els.passwordError.textContent = message || (isEnglishUi() ? "Invalid user." : "使用者錯誤。");
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
    els.serverInput = document.getElementById("server-input");
    els.passwordInput = document.getElementById("password-input");
    els.passwordError = document.getElementById("password-error");
    els.uiLanguageSelect = document.getElementById("ui-language-select");
    els.todayDateLabel = document.getElementById("today-date-label");
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
    els.taskAssignee = document.getElementById("task-assignee");
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
    els.mobileAddBtn = document.getElementById("mobile-add-btn");
    els.mobileUpcomingToggle = document.getElementById("mobile-upcoming-toggle");
    els.mobileUpcomingBadge = document.getElementById("mobile-upcoming-badge");
    els.mobileUpcomingClose = document.getElementById("mobile-upcoming-close");
    els.mobileUpcomingOverlay = document.getElementById("mobile-upcoming-overlay");
    els.mobileTopBtn = document.getElementById("mobile-top-btn");
    els.exportDate = document.getElementById("export-date");
    els.exportStatus = document.getElementById("export-status");
    els.taskListFilter = document.getElementById("task-list-filter");
    els.todayAutoDate = document.getElementById("today-auto-date");
    els.todayExpectedCheckin = document.getElementById("today-expected-checkin");
    els.todayExpectedCheckout = document.getElementById("today-expected-checkout");
    els.todayOccupancyRate = document.getElementById("today-occupancy-rate");
    els.todayOverviewSaveBtn = document.getElementById("today-overview-save-btn");
    els.panelToggleButtons = Array.prototype.slice.call(document.querySelectorAll(".panel-toggle-btn"));
    els.tableHeaderCells = Array.prototype.slice.call(document.querySelectorAll("#task-list-panel thead th"));
  }

  function bindEvents() {
    if (els.uiLanguageSelect) {
      els.uiLanguageSelect.addEventListener("change", handleUiLanguageChange);
    }
    els.taskForm.addEventListener("submit", handleAddTask);
    els.taskCategory.addEventListener("change", updateSubcategoryOptions);
    els.taskSubcategory.addEventListener("change", updateFormLockState);
    els.allDayBtn.addEventListener("click", toggleAllDayMode);
    if (els.taskStartAt) {
      els.taskStartAt.addEventListener("input", handleTimeDigitsInput);
      els.taskStartAt.addEventListener("blur", handleTimeDigitsBlur);
    }
    if (els.taskEndAt) {
      els.taskEndAt.addEventListener("input", handleTimeDigitsInput);
      els.taskEndAt.addEventListener("blur", handleTimeDigitsBlur);
    }
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
    if (els.requestNotificationBtn) {
      els.requestNotificationBtn.addEventListener("click", requestNotificationPermission);
    }
    if (els.notificationToggle) {
      els.notificationToggle.addEventListener("change", handleNotificationToggleChange);
    }
    if (els.mobileUpcomingToggle) {
      els.mobileUpcomingToggle.addEventListener("click", handleMobileUpcomingToggle);
    }
    if (els.mobileUpcomingClose) {
      els.mobileUpcomingClose.addEventListener("click", handleMobileUpcomingClose);
    }
    if (els.mobileUpcomingOverlay) {
      els.mobileUpcomingOverlay.addEventListener("click", handleMobileUpcomingClose);
    }
    if (els.mobileTopBtn) {
      els.mobileTopBtn.addEventListener("click", handleMobileTopClick);
    }
    if (els.mobileAddBtn) {
      els.mobileAddBtn.addEventListener("click", handleMobileAddClick);
    }
    window.addEventListener("resize", syncMobileReminderUi);
    window.addEventListener("scroll", updateMobileFloatingVisibility, { passive: true });
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
    document.addEventListener("keydown", function (event) {
      if (event.key === "Escape" && state.mobileUpcomingOpen) {
        handleMobileUpcomingClose();
      }
    });
  }

  function handleUiLanguageChange() {
    const nextLang = normalizeUiLanguage(els.uiLanguageSelect ? els.uiLanguageSelect.value : state.uiLanguage);
    if (nextLang === normalizeUiLanguage(state.uiLanguage)) {
      return;
    }
    state.uiLanguage = nextLang;
    state.showOriginalByTaskId = {};
    saveUiLanguage();
    applyUiLanguageToStatic();
    setupCategorySelectOptions();
    updateSubcategoryOptions();
    updateFormLockState();
    renderTodayOverviewBar();
    renderAll();
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
      input.type = "text";
      input.setAttribute("inputmode", "numeric");
      input.setAttribute("maxlength", "4");
      input.setAttribute("pattern", "([0-9]{4}|[0-9]{2}:[0-9]{2})");
      input.setAttribute("placeholder", "HHMM");
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
    btn.textContent = isCollapsed ? getUiText("expand") : getUiText("collapse");
    btn.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
  }

  function isMobileViewport() {
    if (typeof window === "undefined" || typeof window.matchMedia !== "function") {
      return false;
    }
    return window.matchMedia("(max-width: " + MOBILE_MAX_WIDTH + "px)").matches;
  }

  function syncMobileReminderUi() {
    const isMobile = isMobileViewport();
    if (!isMobile && state.mobileUpcomingOpen) {
      state.mobileUpcomingOpen = false;
    }
    if (document.body) {
      document.body.classList.toggle("mobile-upcoming-open", isMobile && state.mobileUpcomingOpen);
    }
    if (els.mobileUpcomingToggle) {
      els.mobileUpcomingToggle.setAttribute("aria-expanded", state.mobileUpcomingOpen ? "true" : "false");
    }
    if (els.mobileUpcomingOverlay) {
      els.mobileUpcomingOverlay.setAttribute("aria-hidden", state.mobileUpcomingOpen ? "false" : "true");
    }
    if (els.upcomingBoard) {
      els.upcomingBoard.setAttribute("aria-hidden", isMobile && !state.mobileUpcomingOpen ? "true" : "false");
    }
    updateMobileFloatingVisibility();
  }

  function setMobileUpcomingOpen(open) {
    state.mobileUpcomingOpen = Boolean(open);
    syncMobileReminderUi();
  }

  function handleMobileUpcomingToggle() {
    if (!isMobileViewport()) {
      return;
    }
    setMobileUpcomingOpen(!state.mobileUpcomingOpen);
  }

  function handleMobileUpcomingClose() {
    setMobileUpcomingOpen(false);
  }

  function updateMobileReminderBadge(count) {
    if (!els.mobileUpcomingBadge || !els.mobileUpcomingToggle) {
      return;
    }
    const total = Number.isFinite(count) ? Math.max(0, Math.floor(count)) : 0;
    if (total > 0) {
      els.mobileUpcomingBadge.textContent = String(total);
      els.mobileUpcomingBadge.classList.remove("hidden");
      els.mobileUpcomingToggle.classList.add("has-alert");
      return;
    }
    els.mobileUpcomingBadge.textContent = "0";
    els.mobileUpcomingBadge.classList.add("hidden");
    els.mobileUpcomingToggle.classList.remove("has-alert");
  }

  function handleMobileTopClick() {
    try {
      window.scrollTo({ top: 0, behavior: "smooth" });
    } catch (error) {
      window.scrollTo(0, 0);
    }
    updateMobileFloatingVisibility();
  }

  function handleMobileAddClick() {
    setMobileUpcomingOpen(false);
    scrollToTaskFormPanel();
  }

  function updateMobileFloatingVisibility() {
    if (typeof document === "undefined" || !document.body) {
      return;
    }
    if (!isMobileViewport()) {
      document.body.classList.remove("mobile-top-visible");
      document.body.classList.remove("mobile-version-visible");
      return;
    }
    const root = document.documentElement;
    const body = document.body;
    const scrollTop = Math.max(
      window.pageYOffset || 0,
      root && root.scrollTop ? root.scrollTop : 0,
      body && body.scrollTop ? body.scrollTop : 0,
    );
    const viewportHeight = window.innerHeight || (root ? root.clientHeight : 0) || 0;
    const docHeight = Math.max(
      root && root.scrollHeight ? root.scrollHeight : 0,
      body && body.scrollHeight ? body.scrollHeight : 0,
    );
    const nearBottom = scrollTop + viewportHeight >= docHeight - 32;
    const shouldShowTop = scrollTop > 140 && !state.mobileUpcomingOpen;
    body.classList.toggle("mobile-top-visible", shouldShowTop);
    body.classList.toggle("mobile-version-visible", nearBottom);
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

  function sanitizeTimeDigitsValue(value) {
    return String(value || "")
      .replace(/\D/g, "")
      .slice(0, 4);
  }

  function normalizeCompactTimeText(value) {
    const text = String(value || "").trim();
    if (!text) {
      return "";
    }

    const colonMatch = text.match(/^(\d{1,2}):(\d{1,2})$/);
    if (colonMatch) {
      const hh = Number(colonMatch[1]);
      const mm = Number(colonMatch[2]);
      if (hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59) {
        return String(hh).padStart(2, "0") + ":" + String(mm).padStart(2, "0");
      }
      return "";
    }

    const digits = sanitizeTimeDigitsValue(text);
    if (!digits) {
      return "";
    }

    let hh = Number.NaN;
    let mm = Number.NaN;
    if (digits.length <= 2) {
      hh = Number(digits);
      mm = 0;
    } else if (digits.length === 3) {
      hh = Number(digits.slice(0, 1));
      mm = Number(digits.slice(1));
    } else {
      hh = Number(digits.slice(0, 2));
      mm = Number(digits.slice(2));
    }

    if (!Number.isFinite(hh) || !Number.isFinite(mm) || hh < 0 || hh > 23 || mm < 0 || mm > 59) {
      return "";
    }
    return String(hh).padStart(2, "0") + ":" + String(mm).padStart(2, "0");
  }

  function handleTimeDigitsInput(event) {
    const input = event && event.target ? event.target : null;
    if (!input) {
      return;
    }
    input.value = sanitizeTimeDigitsValue(input.value);
  }

  function handleTimeDigitsBlur(event) {
    const input = event && event.target ? event.target : null;
    if (!input) {
      return;
    }
    const part = input === els.taskEndAt ? "end" : "start";
    input.value = normalizeTimeInputFromAny(input.value, part);
  }

  function normalizeTimeInputFromAny(value, part) {
    const currentValue = String(value || "").trim();
    if (!currentValue) {
      return "";
    }
    const compact = normalizeCompactTimeText(currentValue);
    if (compact) {
      return compact;
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
      els.taskSubcategory.innerHTML = '<option value="">' + escapeHtml(getUiText("subcategoryChooseCategory")) + "</option>";
      els.taskSubcategory.disabled = true;
      els.taskSubcategory.required = false;
      updateFormLockState();
      return;
    }

    const options = getSubcategoryOptions(category);
    if (options.length === 0) {
      els.taskSubcategory.innerHTML = '<option value="">' + escapeHtml(getUiText("subcategoryNone")) + "</option>";
      els.taskSubcategory.disabled = true;
      els.taskSubcategory.required = false;
      updateFormLockState();
      return;
    }

    els.taskSubcategory.disabled = false;
    els.taskSubcategory.required = true;
    els.taskSubcategory.innerHTML =
      '<option value="" disabled selected>' +
      escapeHtml(getUiText("subcategoryChoose")) +
      "</option>" +
      options
        .map(function (item) {
          return '<option value="' + item + '">' + escapeHtml(getSubcategoryDisplayName(item)) + "</option>";
        })
        .join("");

    els.taskSubcategory.value = options.indexOf(current) !== -1 ? current : "";
    updateFormLockState();
  }

  function setupCategorySelectOptions() {
    setupTaskCategoryOptions();
    setupTodayCategoryOptions();
    setupQueryCategoryOptions();
  }

  function setupTaskCategoryOptions() {
    if (!els.taskCategory) {
      return;
    }
    const categories = getActiveCategories();
    const current = String(els.taskCategory.value || "").trim();
    els.taskCategory.innerHTML =
      '<option value="" disabled selected>' +
      escapeHtml(isEnglishUi() ? "Choose category" : "請選擇主分類") +
      "</option>" +
      categories
        .map(function (category) {
          return '<option value="' + category + '">' + escapeHtml(getCategoryDisplayName(category)) + "</option>";
        })
        .join("");
    if (isValidCategory(current)) {
      els.taskCategory.value = current;
    } else {
      els.taskCategory.value = "";
    }
  }

  function setupTodayCategoryOptions() {
    if (!els.todayCategoryFilter) {
      return;
    }
    const categories = getActiveCategories();
    const current = normalizeQueryCategory(state.todayCategory || els.todayCategoryFilter.value || "all");
    els.todayCategoryFilter.innerHTML = ['<option value="all">' + escapeHtml(getUiText("allCategories")) + "</option>"]
      .concat(
        categories.map(function (category) {
          return '<option value="' + category + '">' + escapeHtml(getCategoryDisplayName(category)) + "</option>";
        }),
      )
      .join("");
    els.todayCategoryFilter.value = current === "all" ? "all" : isValidCategory(current) ? current : "all";
    state.todayCategory = els.todayCategoryFilter.value;
  }

  function setupQueryCategoryOptions() {
    if (!els.queryCategory) {
      return;
    }
    const categories = getActiveCategories();
    const current = normalizeQueryCategory(els.queryCategory.value || state.queryCategory);
    const optionsHtml = ['<option value="all">' + escapeHtml(getUiText("allCategories")) + "</option>"]
      .concat(
        categories.map(function (category) {
          return '<option value="' + category + '">' + escapeHtml(getCategoryDisplayName(category)) + "</option>";
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

    [els.taskTitle, els.taskOwner, els.taskAssignee, els.taskStartDate, els.taskEndDate, els.taskStartAt, els.taskEndAt, els.taskPinned, els.taskDescription].forEach(function (el) {
      if (el) {
        el.disabled = locked;
      }
    });
    els.allDayBtn.disabled = locked;
    els.addTaskBtn.disabled = locked;

    if (!isValidCategory(category)) {
      els.categoryTip.textContent = getUiText("categoryTipSelectCategory");
      return;
    }

    if (hasSubcategoryOptions(category) && !normalizeSubcategory(category, subcategory)) {
      els.categoryTip.textContent = getUiText("categoryTipNeedSub");
      return;
    }

    if (hasSubcategoryOptions(category)) {
      els.categoryTip.textContent = getUiText("categoryTipCurrentSub", {
        category: getCategoryDisplayName(category),
        subcategory: getSubcategoryDisplayName(subcategory),
      });
      return;
    }

    els.categoryTip.textContent = getUiText("categoryTipCurrent", {
      category: getCategoryDisplayName(category),
    });
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
      assignee: normalizeAccountName(input.assignee || input.assignedTo || ""),
      completedBy: String(input.completedBy || "").trim(),
      description: String(input.description || "").trim(),
      translations: normalizeTaskTranslations(input.translations),
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

  async function handleAddTask(event) {
    event.preventDefault();

    if (!isCategorySelectionReady()) {
      showToast(isEnglishUi() ? "Please finish category selection first." : "請先完成分類選擇。");
      return;
    }

    const category = String(els.taskCategory.value || "").trim();
    const subcategory = String(els.taskSubcategory.value || "").trim();
    const title = els.taskTitle.value.trim();
    const owner = els.taskOwner.value.trim();
    const assignee = normalizeAccountName(els.taskAssignee ? els.taskAssignee.value : "");
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
      showToast(isEnglishUi() ? "Please fill required fields." : "請填寫完整資料後再新增。");
      return;
    }
    if (hasSubcategoryOptions(category) && !normalizedSubcategory) {
      showToast(isEnglishUi() ? "Please choose a subcategory." : "請先選擇子分類。");
      return;
    }

    if (state.editingTaskId) {
      const task = state.tasks.find(function (item) {
        return item.id === state.editingTaskId;
      });
      if (!task) {
        showToast(isEnglishUi() ? "Task not found." : "找不到要修改的待辦。");
        resetTaskForm();
        return;
      }
      const previousDescription = String(task.description || "").trim();
      task.category = category;
      task.subcategory = normalizedSubcategory;
      task.title = title;
      task.owner = owner;
      task.assignee = assignee;
      task.description = description;
      const noteChanged = previousDescription !== description;
      let translateResult = { changed: false, hasError: false };
      if (noteChanged) {
        task.translations = {};
        translateResult = await cacheTaskTranslationsOnSave(task);
      }
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
      showToast(
        translateResult.hasError
          ? isEnglishUi()
            ? 'Task updated. Translation may be incomplete, use "Translate This".'
            : "待辦已修改，翻譯若未完成可按「翻譯此筆」。"
          : isEnglishUi()
            ? "Task updated."
            : "待辦已修改。",
      );
      return;
    }

    const nowIso = new Date().toISOString();
    const newTask = {
      id: buildId(),
      category: category,
      subcategory: normalizedSubcategory,
      title: title,
      owner: owner,
      assignee: assignee,
      completedBy: "",
      description: description,
      translations: {},
      startAt: range.startAtIso,
      endAt: range.endAtIso,
      dueAt: range.startAtIso,
      allDay: allDay,
      status: "pending",
      pinned: pinned,
      remindedAt: null,
      createdAt: nowIso,
      updatedAt: nowIso,
    };
    const translateResult = await cacheTaskTranslationsOnSave(newTask);
    state.tasks.push(newTask);

    state.tasks.sort(sortByDueTime);
    saveTasks();
    renderAll();
    resetTaskForm();
    showToast(
      translateResult.hasError
        ? isEnglishUi()
          ? 'Task added. Translation may be incomplete, use "Translate This".'
          : "待辦已新增，翻譯若未完成可按「翻譯此筆」。"
        : isEnglishUi()
          ? "Task added."
          : "待辦已新增。",
    );
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

  async function runTaskAction(rawAction, taskId) {
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
        const input = window.prompt(
          normalizeUiLanguage(state.uiLanguage) === "en" ? "Please enter completed by:" : "請輸入完成人：",
          task.completedBy || "",
        );
        if (input === null) {
          return;
        }
        const completedBy = input.trim();
        if (!completedBy) {
          showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Completed by is required." : "請填寫完成人後才能完成。");
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
      showToast(
        task.status === "done"
          ? normalizeUiLanguage(state.uiLanguage) === "en"
            ? "Marked as done."
            : "已標記完成。"
          : normalizeUiLanguage(state.uiLanguage) === "en"
            ? "Restored to pending."
            : "已恢復為待辦。",
      );
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

    if (action === "original") {
      toggleTaskOriginalDescription(task);
      return;
    }

    if (action === "translate") {
      await handleTranslateSingleTask(task);
      return;
    }

    if (action === "pin") {
      task.pinned = !task.pinned;
      touchTask(task);
      saveTasks();
      renderAll();
      showToast(
        task.pinned
          ? normalizeUiLanguage(state.uiLanguage) === "en"
            ? "Pinned."
            : "已設為置頂。"
          : normalizeUiLanguage(state.uiLanguage) === "en"
            ? "Unpinned."
            : "已取消置頂。",
      );
      return;
    }

    if (action === "delete") {
      const ok = window.confirm(
        normalizeUiLanguage(state.uiLanguage) === "en"
          ? "Are you sure you want to cancel and delete this task?"
          : "確定要取消並刪除此待辦嗎？",
      );
      if (!ok) {
        return;
      }
      state.tasks = state.tasks.filter(function (item) {
        return item.id !== taskId;
      });
      if (state.showOriginalByTaskId && typeof state.showOriginalByTaskId === "object") {
        delete state.showOriginalByTaskId[taskId];
      }
      if (state.translatingTaskIds && typeof state.translatingTaskIds === "object") {
        delete state.translatingTaskIds[taskId];
      }
      if (state.editingTaskId === taskId) {
        resetTaskForm();
      }
      state.deletedTaskIds[taskId] = new Date().toISOString();
      saveTasks();
      saveDeletedTaskIds();
      renderAll();
      showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Task deleted." : "待辦已刪除。");
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
    if (els.taskAssignee) {
      refreshAssigneeSelectOptions();
      const assignee = normalizeAccountName(task.assignee || "");
      if (assignee && findAccountByUsername(assignee)) {
        els.taskAssignee.value = assignee;
      } else {
        els.taskAssignee.value = "";
      }
    }
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
    els.addTaskBtn.textContent = getUiText("save");
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
    showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Task loaded. You can edit now." : "已載入待辦，可開始修改。");
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
    showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Edit canceled." : "已取消修改。");
  }

  function handleClearForm() {
    resetTaskForm();
    showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Form cleared." : "表單已清空。");
  }

  function resetTaskForm() {
    state.editingTaskId = null;
    els.taskForm.reset();
    refreshAssigneeSelectOptions();
    updateSubcategoryOptions();
    setDefaultDueTime();
    els.addTaskBtn.textContent = getUiText("addTask");
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
      els.queryResultText.textContent = getUiText("queryAllText", { count: list.length });
      return;
    }
    els.queryResultText.textContent = getUiText("queryConditionText", {
      condition: buildQueryConditionText(),
      count: list.length,
    });
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
    const displayCategoryLabel = selectedCategory === "all" ? getUiText("allCategories") : getCategoryDisplayName(selectedCategory);
    els.todayPinnedList.innerHTML =
      pinnedList.length === 0
        ? '<li class="today-pinned-empty">' + escapeHtml(getUiText("todayNoPinned")) + "</li>"
        : pinnedList
            .map(function (task) {
              return renderTodayTaskItem(task, true);
            })
            .join("");

    if (todayAllList.length === 0) {
      els.todayTaskSummary.textContent = getUiText("todaySummaryZero", { category: displayCategoryLabel });
      els.todayTaskList.innerHTML = '<li class="today-empty">' + escapeHtml(getUiText("todayNoSchedule")) + "</li>";
      return;
    }

    const pendingCount = todayAllList.filter(function (task) {
      return task.status === "pending";
    }).length;
    const doneCount = todayAllList.length - pendingCount;
    els.todayTaskSummary.textContent = getUiText("todaySummary", {
      total: todayAllList.length,
      pending: pendingCount,
      done: doneCount,
      category: displayCategoryLabel,
    });

    els.todayTaskList.innerHTML =
      normalList.length === 0
        ? '<li class="today-empty">' + escapeHtml(getUiText("todayNoNormal")) + "</li>"
        : normalList
            .map(function (task) {
              return renderTodayTaskItem(task, false);
            })
            .join("");
  }

  function renderTodayTaskItem(task, inPinnedZone) {
    const isDone = task.status === "done";
    const statusClass = isDone ? "status-done" : "status-pending";
    const statusText = isDone ? getUiText("statusDone") : getUiText("statusPending");
    const pinTag = task.pinned ? '<span class="today-pill">' + escapeHtml(getUiText("pinned")) + "</span>" : "";
    const pinBtnText = task.pinned ? getUiText("actionUnpin") : getUiText("actionPin");
    const doneBtnText = isDone ? getUiText("actionUndoComplete") : getUiText("actionComplete");
    const countdownText = formatCountdown(getTaskStartAt(task), task.allDay, isDone);
    const displayCategory = getCategoryDisplayName(task.category || getUiText("uncategorized"));
    const displaySubcategory = getSubcategoryDisplayName(getTaskDisplayField(task, "subcategory") || task.subcategory);
    const displayTitle = getTaskDisplayField(task, "title") || task.title || "-";
    const displayDescription = getTaskDisplayDescription(task);
    const canToggleOriginal = hasTranslatedDescriptionForCurrentLanguage(task);
    const showingOriginal = isTaskShowingOriginalDescription(task);
    const translationBusy = isTaskTranslationInProgress(task.id);
    const needsManualTranslate = needsTaskManualTranslationForLanguage(task, state.uiLanguage);
    const originalBtnText = showingOriginal ? getUiText("actionShowTranslated") : getUiText("actionShowOriginal");
    const translateBtnText = translationBusy ? getUiText("actionTranslating") : getUiText("actionTranslateTask");
    const completionText = isDone && task.completedBy ? " | " + getUiText("completedBy") + "：" + escapeHtml(task.completedBy) : "";
    const assigneeDisplay = task.assignee
      ? escapeHtml(task.assignee)
      : escapeHtml(getUiText("taskAssigneeUnassigned"));
    const metaLineText =
      escapeHtml(displayCategory) +
      (displaySubcategory ? " / " + escapeHtml(displaySubcategory) : "") +
      " | " +
      getUiText("owner") +
      "：" +
      escapeHtml(task.owner || getUiText("notFilled")) +
      " | " +
      getUiText("assignee") +
      "：" +
      assigneeDisplay +
      completionText +
      " | " +
      getUiText("countdown") +
      "：" +
      escapeHtml(countdownText);
    const descriptionHtml = displayDescription
      ? '<p class="today-desc"><span class="today-desc-label">' +
        escapeHtml(getUiText("content")) +
        "：</span>" +
        escapeHtml(displayDescription).replace(/\n/g, "<br>") +
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
      escapeHtml(displayCategory || getUiText("uncategorized")) +
      "</span>" +
      pinTag +
      '<span class="today-time">' +
      formatTimeRange(task) +
      "</span>" +
      "</div>" +
      '<p class="today-title">' +
      escapeHtml(displayTitle) +
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
      '">' +
      getUiText("actionEdit") +
      "</button>" +
      '<button class="action-btn action-copy" type="button" data-action="copy" data-id="' +
      escapeHtml(task.id) +
      '">' +
      getUiText("actionCopy") +
      "</button>" +
      (translationBusy || needsManualTranslate
        ? '<button class="action-btn action-translate" type="button" data-action="translate" data-id="' +
          escapeHtml(task.id) +
          '"' +
          (translationBusy ? " disabled" : "") +
          ">" +
          escapeHtml(translateBtnText) +
          "</button>"
        : "") +
      '<button class="action-btn action-original' +
      (showingOriginal ? " active" : "") +
      '" type="button" data-action="original" data-id="' +
      escapeHtml(task.id) +
      '"' +
      (canToggleOriginal ? "" : " disabled") +
      ">" +
      escapeHtml(originalBtnText) +
      "</button>" +
      '<button class="action-btn action-cancel" type="button" data-action="cancel" data-id="' +
      escapeHtml(task.id) +
      '">' +
      getUiText("actionCancel") +
      "</button>" +
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
        '<tr class="empty-row"><td colspan="10">' + escapeHtml(getUiText("taskTableEmpty")) + "</td></tr>";
      return;
    }

    els.tableBody.innerHTML = list
      .map(function (task) {
        const isDone = task.status === "done";
        const statusClass = isDone ? "status-done" : "status-pending";
        const statusText = isDone ? getUiText("statusDone") : getUiText("statusPending");
        const pinnedText = task.pinned ? getUiText("pinned") : "-";
        const pinBtnText = task.pinned ? getUiText("actionUnpin") : getUiText("actionPin");
        const displayCategory = getCategoryDisplayName(task.category || getUiText("uncategorized"));
        const displaySubcategoryRaw = getTaskDisplayField(task, "subcategory") || task.subcategory;
        const displaySubcategory = getSubcategoryDisplayName(displaySubcategoryRaw);
        const displayTitle = getTaskDisplayField(task, "title") || task.title || "-";
        const displayDescription = getTaskDisplayDescription(task);
        const canToggleOriginal = hasTranslatedDescriptionForCurrentLanguage(task);
        const showingOriginal = isTaskShowingOriginalDescription(task);
        const translationBusy = isTaskTranslationInProgress(task.id);
        const needsManualTranslate = needsTaskManualTranslationForLanguage(task, state.uiLanguage);
        const originalBtnText = showingOriginal ? getUiText("actionShowTranslated") : getUiText("actionShowOriginal");
        const translateBtnText = translationBusy ? getUiText("actionTranslating") : getUiText("actionTranslateTask");
        const completedByText = task.status === "done" && task.completedBy ? escapeHtml(task.completedBy) : "-";
        const ownerCellText = escapeHtml(task.owner) + "<br><span style=\"color:#766e61;\">" + escapeHtml(getUiText("assignee")) + "：" + escapeHtml(task.assignee || getUiText("taskAssigneeUnassigned")) + "</span>";
        const subcategoryHtml = displaySubcategory
          ? escapeHtml(displaySubcategory)
          : '<span style="color:#7a9198;">-</span>';
        const descHtml = displayDescription
          ? escapeHtml(displayDescription).replace(/\n/g, "<br>")
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
          escapeHtml(displayCategory) +
          "</td>" +
          "<td>" +
          subcategoryHtml +
          "</td>" +
          "<td>" +
          escapeHtml(displayTitle) +
          "</td>" +
          "<td>" +
          ownerCellText +
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
          (isDone ? getUiText("actionUndoComplete") : getUiText("actionComplete")) +
          "</button>" +
          '<button class="action-btn action-edit" type="button" data-action="edit" data-id="' +
          task.id +
          '">' +
          getUiText("actionEdit") +
          "</button>" +
          '<button class="action-btn action-copy" type="button" data-action="copy" data-id="' +
          task.id +
          '">' +
          getUiText("actionCopy") +
          "</button>" +
          (translationBusy || needsManualTranslate
            ? '<button class="action-btn action-translate" type="button" data-action="translate" data-id="' +
              task.id +
              '"' +
              (translationBusy ? " disabled" : "") +
              ">" +
              escapeHtml(translateBtnText) +
              "</button>"
            : "") +
          '<button class="action-btn action-original' +
          (showingOriginal ? " active" : "") +
          '" type="button" data-action="original" data-id="' +
          task.id +
          '"' +
          (canToggleOriginal ? "" : " disabled") +
          ">" +
          escapeHtml(originalBtnText) +
          "</button>" +
          '<button class="action-btn action-pin" type="button" data-action="pin" data-id="' +
          task.id +
          '">' +
          pinBtnText +
          "</button>" +
          '<button class="action-btn action-delete" type="button" data-action="delete" data-id="' +
          task.id +
          '">' +
          getUiText("actionDelete") +
          "</button>" +
          "</div></td>" +
          "</tr>"
        );
      })
      .join("");
  }

  function renderUpcomingBoard() {
    if (!els.upcomingSummary || !els.upcomingTaskList) {
      return;
    }
    const now = Date.now();
    const dueNow = [];
    const within30 = [];
    const within60 = [];

    state.tasks
      .filter(function (task) {
        return task.status === "pending" && Boolean(getTaskStartAt(task)) && !task.allDay;
      })
      .sort(sortByDueTime)
      .forEach(function (task) {
        const diffMs = new Date(getTaskStartAt(task)).getTime() - now;
        if (diffMs <= 0) {
          dueNow.push(task);
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

    updateMobileReminderBadge(dueNow.length);

    const total = dueNow.length + within30.length + within60.length;
    if (total === 0) {
      els.upcomingSummary.textContent = getUiText("upcomingSummaryEmpty");
      els.upcomingTaskList.innerHTML = '<li class="upcoming-empty">' + escapeHtml(getUiText("upcomingListEmpty")) + "</li>";
      return;
    }

    const summaryParts = [];
    if (dueNow.length > 0) {
      summaryParts.push(getUiText("timeWindowDue") + " " + dueNow.length);
    }
    summaryParts.push(getUiText("timeWindow30") + " " + within30.length);
    summaryParts.push(getUiText("timeWindow60") + " " + within60.length);
    const isEnglish = normalizeUiLanguage(state.uiLanguage) === "en";
    els.upcomingSummary.textContent = summaryParts.join(isEnglish ? ", " : "，") + (isEnglish ? "." : "。");

    const items = [];
    dueNow.forEach(function (task) {
      items.push(renderUpcomingItem(task, getUiText("timeWindowDue")));
    });
    within30.forEach(function (task) {
      items.push(renderUpcomingItem(task, getUiText("timeWindow30")));
    });
    within60.forEach(function (task) {
      items.push(renderUpcomingItem(task, getUiText("timeWindow60")));
    });
    els.upcomingTaskList.innerHTML = items.join("");
  }

  function renderUpcomingItem(task, windowTag) {
    const displayTitle = getTaskDisplayField(task, "title") || task.title || "-";
    const displayDescription = getTaskDisplayDescription(task);
    const displaySubcategory = getSubcategoryDisplayName(getTaskDisplayField(task, "subcategory") || task.subcategory);
    const detailText = displayDescription
      ? escapeHtml(displayDescription).replace(/\n/g, "<br>")
      : escapeHtml(getUiText("noData"));
    const countdownText = formatCountdown(getTaskStartAt(task), task.allDay, false);
    const metaLineText =
      escapeHtml(getCategoryDisplayName(task.category || getUiText("uncategorized"))) +
      (displaySubcategory ? " / " + escapeHtml(displaySubcategory) : "") +
      " | " +
      getUiText("owner") +
      "：" +
      escapeHtml(task.owner || getUiText("notFilled")) +
      " | " +
      getUiText("assignee") +
      "：" +
      escapeHtml(task.assignee || getUiText("taskAssigneeUnassigned")) +
      " | " +
      getUiText("countdown") +
      "：" +
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
      escapeHtml(displayTitle) +
      "</p>" +
      '<p class="upcoming-meta upcoming-meta-line">' +
      metaLineText +
      "</p>" +
      '<p class="upcoming-detail"><span class="upcoming-detail-label">' +
      escapeHtml(getUiText("time")) +
      "：</span>" +
      escapeHtml(formatDueDisplay(task)) +
      "</p>" +
      '<p class="upcoming-detail"><span class="upcoming-detail-label">' +
      escapeHtml(getUiText("content")) +
      "：</span>" +
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
    const taskTitle = getTaskDisplayField(task, "title") || task.title || "-";
    const text =
      normalizeUiLanguage(state.uiLanguage) === "en"
        ? 'Task "' + taskTitle + '" is now due.'
        : "待辦「" + taskTitle + "」已到時間。";
    showToast(text);

    if (!els.notificationToggle.checked) {
      return;
    }
    if (!("Notification" in window)) {
      return;
    }
    if (Notification.permission === "granted") {
      const categoryText =
        getCategoryDisplayName(task.category) +
        (task.subcategory ? "/" + getSubcategoryDisplayName(getTaskDisplayField(task, "subcategory") || task.subcategory) : "");
      new Notification(normalizeUiLanguage(state.uiLanguage) === "en" ? "Handover Reminder" : "工作交接提醒", {
        body:
          taskTitle +
          (normalizeUiLanguage(state.uiLanguage) === "en" ? " (Category: " : "（分類：") +
          categoryText +
          (normalizeUiLanguage(state.uiLanguage) === "en" ? ", Owner: " : "，填寫人：") +
          (task.owner || getUiText("notFilled")) +
          (normalizeUiLanguage(state.uiLanguage) === "en" ? ")" : "）"),
      });
    }
  }

  async function requestNotificationPermission() {
    if (!("Notification" in window)) {
      showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Browser notifications are not supported." : "目前瀏覽器不支援通知功能。");
      return;
    }
    try {
      const result = Notification.permission === "granted" ? "granted" : await Notification.requestPermission();
      if (result === "granted") {
        const syncResult = await ensureCloudPushSubscription(getCurrentServerConfig(), {
          requestPermission: false,
        });
        if (syncResult.ok) {
          showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Notifications enabled (offline push ready)." : "已授權通知（離線推播已啟用）。");
          return;
        }
        showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Notifications enabled." : "已授權瀏覽器通知。");
      } else if (result === "denied") {
        showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Notifications blocked. Enable in browser settings." : "通知已被封鎖，可到瀏覽器設定開啟。");
      } else {
        showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Notification permission not granted." : "通知授權尚未開啟。");
      }
    } catch (error) {
      console.error("requestPermission error", error);
      showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Notification authorization failed." : "通知授權失敗。");
    }
  }

  function handleNotificationToggleChange() {
    if (!els.notificationToggle) {
      return;
    }
    if (!els.notificationToggle.checked) {
      disableCloudPushSubscription(getCurrentServerConfig()).catch(function () {
        return null;
      });
      return;
    }
    ensureCloudPushSubscription(getCurrentServerConfig(), {
      requestPermission: false,
    }).catch(function () {
      return null;
    });
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
      const tasksForExcel = context.includeDatePrefix
        ? filterExcelFutureTasksAfterToday(context.list)
        : context.list;
      exportExcelFile(tasksForExcel, context.effectiveDate, context.includeDatePrefix);
      showToast("Excel file exported.");
    } catch (error) {
      console.error("exportExcel error", error);
      showToast("Excel 匯出失敗，請稍後重試。");
    }
  }

  function filterExcelFutureTasksAfterToday(tasks) {
    const todayKey = toDateKey(new Date());
    return (Array.isArray(tasks) ? tasks : []).filter(function (task) {
      const key = getExportTaskDateKey(task);
      return Boolean(key) && key > todayKey;
    });
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
      applyExcelWorksheetStyles(worksheet, built.meta, headers.length);
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
    const meta = {
      mergedRows: [],
      headerRows: [],
      dateTitleRows: [],
      dayBlocks: [],
    };

    function pushMergedRow(text, type) {
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
      meta.mergedRows.push(rowIndex);
      if (type === "date-title") {
        meta.dateTitleRows.push(rowIndex);
      }
      return rowIndex;
    }

    function pushHeaderRow() {
      const rowIndex = aoa.length;
      aoa.push(headers.slice());
      meta.headerRows.push(rowIndex);
      return rowIndex;
    }

    function pushTaskRows(taskList) {
      const startRow = aoa.length;
      const rows = buildExcelRows(taskList);
      if (rows.length === 0) {
        const emptyRow = [];
        for (let i = 0; i < cols; i += 1) {
          emptyRow.push("-");
        }
        aoa.push(emptyRow);
        return { startRow: startRow, endRow: aoa.length - 1 };
      }
      rows.forEach(function (row) {
        aoa.push(
          headers.map(function (key) {
            return row[key];
          }),
        );
      });
      return { startRow: startRow, endRow: aoa.length - 1 };
    }

    if (!includeDatePrefix) {
      pushMergedRow(buildArrDepOccLine(exportDate), "summary");
    }
    pushMergedRow(includeDatePrefix ? "Future To-Do / 未來待辦事項" : "Daily Briefing / 每日報告", "section-title");
    aoa.push([]);

    if (!includeDatePrefix) {
      const singleBlockStart = aoa.length;
      pushHeaderRow();
      const singleRows = pushTaskRows(list);
      meta.dayBlocks.push({
        startRow: singleBlockStart,
        endRow: Math.max(singleRows.endRow, singleBlockStart),
        color: EXCEL_DAY_BLOCK_COLORS[0],
      });
      return { aoa: aoa, merges: merges, meta: meta };
    }

    const grouped = groupExportTasksByDate(list);
    if (grouped.length === 0) {
      const emptyBlockStart = aoa.length;
      pushHeaderRow();
      const emptyRows = pushTaskRows([]);
      meta.dayBlocks.push({
        startRow: emptyBlockStart,
        endRow: Math.max(emptyRows.endRow, emptyBlockStart),
        color: EXCEL_DAY_BLOCK_COLORS[0],
      });
      return { aoa: aoa, merges: merges, meta: meta };
    }

    grouped.forEach(function (group, index) {
      const blockStart = aoa.length;
      pushMergedRow("日期：" + group.label, "date-title");
      pushHeaderRow();
      const taskRows = pushTaskRows(group.tasks);
      meta.dayBlocks.push({
        startRow: blockStart,
        endRow: Math.max(taskRows.endRow, blockStart),
        color: EXCEL_DAY_BLOCK_COLORS[index % EXCEL_DAY_BLOCK_COLORS.length],
      });
      if (index < grouped.length - 1) {
        aoa.push([]);
      }
    });
    return { aoa: aoa, merges: merges, meta: meta };
  }

  function applyExcelWorksheetStyles(worksheet, meta, colCount) {
    if (!worksheet || !window.XLSX || !window.XLSX.utils) {
      return;
    }
    const cols = Number(colCount) > 0 ? Number(colCount) : 5;
    const styleMeta = meta && typeof meta === "object" ? meta : {};
    const dayBlocks = Array.isArray(styleMeta.dayBlocks) ? styleMeta.dayBlocks : [];
    const headerRows = Array.isArray(styleMeta.headerRows) ? styleMeta.headerRows : [];
    const mergedRows = Array.isArray(styleMeta.mergedRows) ? styleMeta.mergedRows : [];
    const dateTitleRowSet = new Set(Array.isArray(styleMeta.dateTitleRows) ? styleMeta.dateTitleRows : []);

    function encodeCell(rowIndex, colIndex) {
      return window.XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    }

    function ensureCell(rowIndex, colIndex) {
      const address = encodeCell(rowIndex, colIndex);
      if (!worksheet[address]) {
        worksheet[address] = { t: "s", v: "" };
      }
      return worksheet[address];
    }

    function createBorder(color) {
      const rgb = String(color || "").trim() || EXCEL_BLOCK_BORDER_COLOR;
      return {
        top: { style: "thin", color: { rgb: rgb } },
        bottom: { style: "thin", color: { rgb: rgb } },
        left: { style: "thin", color: { rgb: rgb } },
        right: { style: "thin", color: { rgb: rgb } },
      };
    }

    function createFont(options) {
      const style = options && typeof options === "object" ? options : {};
      return {
        name: EXPORT_FONT_EAST_ASIA,
        sz: Number(style.size) > 0 ? Number(style.size) : 12,
        bold: Boolean(style.bold),
        color: { rgb: String(style.color || "1F2A2A").replace(/^#/, "") },
      };
    }

    function createAlignment(options) {
      const style = options && typeof options === "object" ? options : {};
      return {
        horizontal: style.horizontal || "center",
        vertical: style.vertical || "center",
        wrapText: style.wrapText !== false,
      };
    }

    function applyStyle(rowIndex, colIndex, options) {
      const style = options && typeof options === "object" ? options : {};
      const cell = ensureCell(rowIndex, colIndex);
      const nextStyle = {
        font: createFont(style.font),
        alignment: createAlignment(style.alignment),
      };
      if (style.fillColor) {
        nextStyle.fill = {
          patternType: "solid",
          fgColor: { rgb: String(style.fillColor).replace(/^#/, "") },
        };
      }
      if (style.useBorder !== false) {
        nextStyle.border = createBorder(style.borderColor);
      }
      cell.s = nextStyle;
    }

    dayBlocks.forEach(function (block) {
      const startRow = Number(block && block.startRow);
      const endRow = Number(block && block.endRow);
      if (!Number.isFinite(startRow) || !Number.isFinite(endRow) || endRow < startRow) {
        return;
      }
      const fillColor = String(block && block.color ? block.color : EXCEL_DAY_BLOCK_COLORS[0]).replace(/^#/, "");
      for (let r = startRow; r <= endRow; r += 1) {
        for (let c = 0; c < cols; c += 1) {
          applyStyle(r, c, {
            fillColor: fillColor,
            borderColor: EXCEL_BLOCK_BORDER_COLOR,
            font: {
              bold: headerRows.indexOf(r) >= 0 || dateTitleRowSet.has(r),
              size: 12,
            },
          });
        }
      }
    });

    mergedRows.forEach(function (rowIndex) {
      for (let c = 0; c < cols; c += 1) {
        applyStyle(rowIndex, c, {
          fillColor: dateTitleRowSet.has(rowIndex) ? "F1DFC0" : "EFE8DB",
          borderColor: EXCEL_BLOCK_BORDER_COLOR,
          font: {
            bold: true,
            size: dateTitleRowSet.has(rowIndex) ? 12 : 13,
          },
        });
      }
    });

    headerRows.forEach(function (rowIndex) {
      for (let c = 0; c < cols; c += 1) {
        applyStyle(rowIndex, c, {
          fillColor: "E6D6BC",
          borderColor: EXCEL_BLOCK_BORDER_COLOR,
          font: {
            bold: true,
            size: 12,
          },
        });
      }
    });
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
    state.translatingTaskIds = {};
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
      return task.category === "待回覆信件";
    });

    const daily = (Array.isArray(currentList) ? currentList.slice() : []).sort(sortForTaskTable);

    const banquets = statusPool.filter(function (task) {
      return task.category === "會議室" || task.category === "團桌" || task.category === "預訂";
    });

    const future = statusPool.filter(function (task) {
      const startAt = getTaskStartAt(task);
      if (!startAt) {
        return false;
      }
      return new Date(startAt).getTime() > endOfTargetDay;
    });

    const transfer = statusPool.filter(function (task) {
      return task.category === "包裹代收";
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
    return getActiveExportColorMap()[key] || EXPORT_DEFAULT_COLOR;
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
      return getUiText("pending");
    }
    if (status === "done") {
      return getUiText("done");
    }
    return getUiText("allStatus");
  }

  function getQueryCategoryLabel(category) {
    const normalized = normalizeQueryCategory(category);
    return normalized === "all" ? getUiText("allCategories") : getCategoryDisplayName(normalized);
  }

  function buildQueryConditionText() {
    return buildConditionText(state.queryDate, state.queryStatus, state.queryCategory, state.queryKeyword);
  }

  function buildConditionText(dateValue, statusValue, categoryValue, keywordValue) {
    const dateText = dateValue ? dateValue.replace(/-/g, "/") : normalizeUiLanguage(state.uiLanguage) === "en" ? "All Dates" : "全部日期";
    const categoryText = getQueryCategoryLabel(categoryValue);
    const keywordText = normalizeQueryKeyword(keywordValue);
    const keywordPrefix = normalizeUiLanguage(state.uiLanguage) === "en" ? "Keyword: " : "關鍵字：";
    const keywordLabel = keywordText ? keywordPrefix + keywordText : keywordPrefix + (normalizeUiLanguage(state.uiLanguage) === "en" ? "All" : "全部");
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
        const translatedTitle = getTaskTranslatedField(task, "title", state.uiLanguage);
        const translatedDesc = getTaskTranslatedField(task, "description", state.uiLanguage);
        const translatedSub = getTaskTranslatedField(task, "subcategory", state.uiLanguage);
        const haystack = [
          task.category,
          task.subcategory,
          task.title,
          task.owner,
          task.completedBy,
          task.description,
          getCategoryDisplayName(task.category),
          getSubcategoryDisplayName(task.subcategory),
          translatedSub,
          translatedTitle,
          translatedDesc,
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
      showToast(normalizeUiLanguage(state.uiLanguage) === "en" ? "Nothing to copy." : "無可複製內容。");
      return;
    }
    copyTextToClipboard(text).then(function (ok) {
      showToast(
        ok
          ? normalizeUiLanguage(state.uiLanguage) === "en"
            ? "Task text copied."
            : "已複製事項文字。"
          : normalizeUiLanguage(state.uiLanguage) === "en"
            ? "Copy failed. Please copy manually."
            : "複製失敗，請手動複製。",
      );
    });
  }

  function buildTaskCopyText(task) {
    if (!task) {
      return "";
    }
    const isDone = task.status === "done";
    const displayCategory = getCategoryDisplayName(task.category || getUiText("uncategorized")) || "-";
    const displaySubcategory =
      getSubcategoryDisplayName(getTaskDisplayField(task, "subcategory") || task.subcategory) || "-";
    const displayTitle = getTaskDisplayField(task, "title") || task.title || "-";
    const displayDescription = getTaskDisplayDescription(task) || task.description || "-";
    const lines = [];
    lines.push(getUiText("taskCategoryLabel") + "：" + displayCategory);
    lines.push(getUiText("taskSubcategoryLabel") + "：" + displaySubcategory);
    lines.push(getUiText("taskTitleLabel") + "：" + displayTitle);
    lines.push(getUiText("owner") + "：" + (task.owner || "-"));
    lines.push(getUiText("assignee") + "：" + (task.assignee || getUiText("taskAssigneeUnassigned")));
    if (isDone) {
      lines.push(getUiText("completedBy") + "：" + (task.completedBy || "-"));
    }
    lines.push(getUiText("time") + "：" + formatTimeRange(task));
    lines.push(getUiText("content") + "：" + displayDescription);
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
      const useIncoming = getTaskRevisionMs(normalized) >= getTaskRevisionMs(existing);
      const winner = useIncoming ? normalized : existing;
      const loser = useIncoming ? existing : normalized;
      mergeTaskTranslationCache(winner, loser);
      merged.set(normalized.id, winner);
    });

    return Array.from(merged.values()).sort(sortByDueTime);
  }

  function mergeTaskTranslationCache(targetTask, sourceTask) {
    if (!targetTask || !sourceTask) {
      return;
    }
    if (!targetTask.translations || typeof targetTask.translations !== "object" || Array.isArray(targetTask.translations)) {
      targetTask.translations = {};
    }
    const sourceTranslations =
      sourceTask.translations && typeof sourceTask.translations === "object" && !Array.isArray(sourceTask.translations)
        ? sourceTask.translations
        : {};
    const targetSig = getTaskTranslationSignature(targetTask);
    const targetDesc = String(targetTask.description || "").trim();
    const sourceDesc = String(sourceTask.description || "").trim();
    if (targetDesc !== sourceDesc) {
      return;
    }

    ["en", "zh"].forEach(function (lang) {
      const current = targetTask.translations[lang];
      if (
        current &&
        typeof current === "object" &&
        String(current._sig || "") === targetSig &&
        isTaskTranslationUsable(targetTask, current, lang)
      ) {
        return;
      }
      const candidate = sourceTranslations[lang];
      if (!candidate || typeof candidate !== "object") {
        return;
      }
      const candidateDesc = String(candidate.description || "").trim();
      if (!candidateDesc) {
        return;
      }
      if (shouldTranslateTextForTarget(targetDesc, lang) && candidateDesc === targetDesc) {
        return;
      }
      targetTask.translations[lang] = {
        title: String(targetTask.title || "").trim(),
        description: candidateDesc,
        subcategory: String(targetTask.subcategory || "").trim(),
        _sig: targetSig,
      };
    });
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
    const idx = getActiveCategories().indexOf(String(category || "").trim());
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
    const categories = getActiveCategories();
    const value = String(raw || "").trim();
    if (categories.indexOf(value) === -1) {
      return categories[0];
    }
    return value;
  }

  function isValidCategory(value) {
    return getActiveCategories().indexOf(value) !== -1;
  }

  function getSubcategoryOptions(category) {
    const list = getActiveSubcategoryMap()[category];
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
        message: isEnglishUi() ? "Please select start date first." : "請先選擇開始日期。",
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
          message: isEnglishUi() ? "Invalid date format." : "日期格式無效，請重新選擇。",
        };
      }
      if (endMs < startMs) {
        return {
          ok: false,
          message: isEnglishUi() ? "End date cannot be earlier than start date." : "結束日期不可早於開始日期。",
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
        message: isEnglishUi() ? "Invalid start time format." : "開始時間格式無效，請重新輸入。",
      };
    }
    if (endText && Number.isNaN(endMs)) {
      return {
        ok: false,
        message: isEnglishUi() ? "Invalid end time format." : "結束時間格式無效，請重新輸入。",
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
        message: isEnglishUi() ? "End time cannot be earlier than start time." : "結束時間不可早於開始時間。",
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
    const compactTime = normalizeCompactTimeText(normalized);
    if (compactTime) {
      const timeOnly = compactTime.match(/^(\d{2}):(\d{2})$/);
      if (!timeOnly) {
        return Number.NaN;
      }
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
          return formatDateOnly(startAt) + " " + (isEnglishUi() ? "All Day" : "全天");
        }
        return (
          formatDateOnly(startAt) +
          " " +
          (isEnglishUi() ? "to" : "至") +
          " " +
          formatDateOnly(endAt) +
          " " +
          (isEnglishUi() ? "All Day" : "全天")
        );
      }
      const allDayBase = startAt || endAt;
      return formatDateOnly(allDayBase) + " " + (isEnglishUi() ? "All Day" : "全天");
    }
    if (startAt && endAt) {
      const sameDate = toDateKey(new Date(startAt)) === toDateKey(new Date(endAt));
      if (sameDate) {
        return (
          formatDateOnly(startAt) +
          " " +
          formatTime(startAt, false) +
          " " +
          (isEnglishUi() ? "to" : "至") +
          " " +
          formatTime(endAt, false)
        );
      }
      return formatDateTime(startAt) + " " + (isEnglishUi() ? "to" : "至") + " " + formatDateTime(endAt);
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
      return isEnglishUi() ? "All Day" : "全天";
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
      return isEnglishUi() ? "All Day" : "全天";
    }
    const startAt = getTaskStartAt(task);
    const endAt = getTaskEndAt(task);
    if (!startAt && !endAt) {
      return "--:--";
    }
    if (startAt && endAt) {
      return formatTime(startAt, false) + " " + (isEnglishUi() ? "to" : "至") + " " + formatTime(endAt, false);
    }
    return formatTime(startAt || endAt, false);
  }

  function formatCountdown(input, allDay, isDone) {
    if (isDone) {
      return isEnglishUi() ? "Completed" : "已完成";
    }
    if (allDay) {
      return isEnglishUi() ? "All-day task" : "全天事項";
    }
    if (!input) {
      return isEnglishUi() ? "No time" : "無時間";
    }
    const dueMs = input instanceof Date ? input.getTime() : new Date(input).getTime();
    if (Number.isNaN(dueMs)) {
      return isEnglishUi() ? "No time" : "無時間";
    }
    const diffSec = Math.floor((dueMs - Date.now()) / 1000);
    if (diffSec >= 0) {
      return (isEnglishUi() ? "Remaining " : "剩餘 ") + formatDuration(diffSec);
    }
    return (isEnglishUi() ? "Overdue " : "逾時 ") + formatDuration(Math.abs(diffSec));
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
      return isEnglishUi() ? days + "d " + hh + ":" + mm + ":" + ss : days + "天 " + hh + ":" + mm + ":" + ss;
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

