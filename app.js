(function () {
  "use strict";

  const STORAGE_KEY = "handover_tasks_v1";
  const TODAY_OVERVIEW_KEY = "handover_today_overview_v1";
  const ACCESS_PASSWORD = "caesarmetro";
  const BACKUP_TYPE = "handover-backup";
  const BACKUP_VERSION = "0.7";
  const REMINDER_CHECK_MS = 30 * 1000;
  const COUNTDOWN_REFRESH_MS = 1000;
  const AUTO_SAVE_INTERVAL_MS = 10 * 1000;
  const AUTO_LOAD_INTERVAL_MS = 30 * 1000;
  const TOAST_MS = 3000;
  const CATEGORIES = ["廣場", "包裹代收", "車輛安排", "大廳", "會議室", "團桌", "客房", "餐飲部", "待回覆信件", "公告"];
  const SUBCATEGORY_MAP = {
    廣場: ["保留車位", "其他"],
    包裹代收: ["團體", "散客", "其他"],
    車輛安排: ["禮賓車", "計程車", "其他"],
    大廳: ["行李寄放", "其他"],
    客房: ["下行李", "房務相關事項", "其他"],
  };

  const state = {
    tasks: [],
    queryDate: "",
    queryStatus: "all",
    taskListFilter: "all",
    todayOverview: {
      checkin: "",
      checkout: "",
      occupancy: "",
    },
    todayCategory: "all",
    formAllDay: false,
    editingTaskId: null,
    reminderTimer: null,
    countdownTimer: null,
    autoSaveTimer: null,
    autoLoadTimer: null,
    autoSaveFileHandle: null,
    lastAutoSaveAt: null,
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
    state.initialized = true;
    state.tasks = loadTasks();
    state.todayOverview = loadTodayOverview();
    bindEvents();
    syncCollapsiblePanels();
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
    updateAutoSaveStatus();
    renderTodayOverviewBar();
    renderAll();
    startReminderLoop();
    startCountdownLoop();
    startAutoSaveLoop();
    startAutoLoadLoop();
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

  function handlePasswordSubmit(event) {
    event.preventDefault();
    const entered = String(els.passwordInput.value || "").trim();
    if (entered !== ACCESS_PASSWORD) {
      showPasswordError("使用者錯誤，請再試一次。");
      return;
    }
    unlockAccessGate();
    init();
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
    els.toast = document.getElementById("toast");
    els.notificationToggle = document.getElementById("notification-toggle");
    els.requestNotificationBtn = document.getElementById("request-notification-btn");
    els.exportDate = document.getElementById("export-date");
    els.exportStatus = document.getElementById("export-status");
    els.pickAutoSaveFileBtn = document.getElementById("pick-auto-save-file-btn");
    els.autoSaveStatus = document.getElementById("auto-save-status");
    els.saveDataBtn = document.getElementById("save-data-btn");
    els.loadDataBtn = document.getElementById("load-data-btn");
    els.loadDataInput = document.getElementById("load-data-input");
    els.taskListFilter = document.getElementById("task-list-filter");
    els.todayAutoDate = document.getElementById("today-auto-date");
    els.todayExpectedCheckin = document.getElementById("today-expected-checkin");
    els.todayExpectedCheckout = document.getElementById("today-expected-checkout");
    els.todayOccupancyRate = document.getElementById("today-occupancy-rate");
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
      els.todayOccupancyRate.addEventListener("input", handleTodayOverviewInput);
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
    els.requestNotificationBtn.addEventListener("click", requestNotificationPermission);
    if (els.saveDataBtn) {
      els.saveDataBtn.addEventListener("click", handleSaveData);
    }
    if (els.pickAutoSaveFileBtn) {
      els.pickAutoSaveFileBtn.addEventListener("click", handlePickAutoSaveFile);
    }
    if (els.loadDataBtn) {
      els.loadDataBtn.addEventListener("click", handleLoadDataClick);
    }
    if (els.loadDataInput) {
      els.loadDataInput.addEventListener("change", handleLoadDataChange);
    }
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
    els.taskStartAt.value = "";
    els.taskEndAt.value = "";
  }

  function apply24HourInputMode() {
    [els.taskStartAt, els.taskEndAt].forEach(function (input) {
      if (!input) {
        return;
      }
      if (input.type === "datetime-local") {
        input.setAttribute("step", "60");
        input.setAttribute("lang", "en-GB");
      }
    });
  }

  function handleTodayOverviewInput() {
    state.todayOverview.checkin = normalizeTodayOverviewValue(els.todayExpectedCheckin ? els.todayExpectedCheckin.value : "");
    state.todayOverview.checkout = normalizeTodayOverviewValue(els.todayExpectedCheckout ? els.todayExpectedCheckout.value : "");
    state.todayOverview.occupancy = normalizeOccupancyRateValue(els.todayOccupancyRate ? els.todayOccupancyRate.value : "");
    if (els.todayOccupancyRate) {
      els.todayOccupancyRate.value = state.todayOverview.occupancy;
    }
    saveTodayOverview();
  }

  function renderTodayOverviewBar() {
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
    const currentStart = String(els.taskStartAt.value || "").trim();
    const currentEnd = String(els.taskEndAt.value || "").trim();
    state.formAllDay = next;

    els.allDayBtn.setAttribute("aria-pressed", next ? "true" : "false");
    els.allDayBtn.classList.toggle("active", next);
    els.allDayBtn.textContent = next ? "全日中" : "全日";
    els.taskStartAt.type = next ? "date" : "datetime-local";
    els.taskEndAt.type = next ? "date" : "datetime-local";
    apply24HourInputMode();

    els.taskStartAt.value = normalizeRangeInputValue(currentStart, next, "start");
    els.taskEndAt.value = normalizeRangeInputValue(currentEnd, next, "end");
  }

  function normalizeRangeInputValue(value, allDay, part) {
    const currentValue = String(value || "").trim();
    if (!currentValue) {
      return "";
    }
    if (allDay) {
      return currentValue.includes("T") ? currentValue.slice(0, 10) : currentValue;
    }
    if (currentValue.includes("T")) {
      return currentValue;
    }
    return currentValue + (part === "end" ? "T18:00" : "T09:00");
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

    [els.taskTitle, els.taskOwner, els.taskStartAt, els.taskEndAt, els.taskPinned, els.taskDescription].forEach(function (el) {
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
      const raw = localStorage.getItem(STORAGE_KEY);
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
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state.tasks));
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
      createdAt: Number.isNaN(createdAtMs) ? new Date().toISOString() : new Date(createdAtMs).toISOString(),
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
    const startAtInput = String(els.taskStartAt.value || "").trim();
    const endAtInput = String(els.taskEndAt.value || "").trim();
    const pinned = Boolean(els.taskPinned.checked);
    const allDay = Boolean(state.formAllDay);
    const normalizedSubcategory = normalizeSubcategory(category, subcategory);
    const range = parseTaskTimeRange(startAtInput, endAtInput, allDay);
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
      state.tasks.sort(sortByDueTime);
      saveTasks();
      renderAll();
      resetTaskForm();
      showToast("待辦已修改。");
      return;
    }

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
      createdAt: new Date().toISOString(),
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
    if (els.queryStatus) {
      els.queryStatus.value = state.queryStatus;
    }
    renderAll();
  }

  function clearDateQuery() {
    state.queryDate = "";
    state.queryStatus = "all";
    els.queryDate.value = "";
    if (els.queryStatus) {
      els.queryStatus.value = "all";
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
      } else {
        task.status = "pending";
        task.remindedAt = null;
        task.completedBy = "";
      }
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
      saveTasks();
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
    if (startAt) {
      els.taskStartAt.value = task.allDay ? toDateKey(new Date(startAt)) : toDateTimeLocalValue(new Date(startAt));
    } else {
      els.taskStartAt.value = "";
    }
    if (endAt) {
      els.taskEndAt.value = task.allDay ? toDateKey(new Date(endAt)) : toDateTimeLocalValue(new Date(endAt));
    } else {
      els.taskEndAt.value = "";
    }
    els.taskPinned.checked = Boolean(task.pinned);
    els.taskDescription.value = task.description || "";
    els.addTaskBtn.textContent = "儲存修改";
    els.cancelEditBtn.classList.remove("hidden");
    updateFormLockState();
    showToast("已載入待辦，可開始修改。");
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
    if (!state.queryDate && state.queryStatus === "all") {
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
        return Boolean(task.pinned);
      })
      .sort(sortByDueTime);

    const normalList = todayAllList.filter(function (task) {
      return !task.pinned;
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
    return getTasksByDateAndStatus(state.queryDate, state.queryStatus);
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
      renderTodayOverviewBar();
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

  async function handleExportWord() {
    const exportDate = els.exportDate ? String(els.exportDate.value || "").trim() : "";
    const exportStatus = normalizeQueryStatus(els.exportStatus ? els.exportStatus.value : "all");
    const list = getTasksByDateAndStatus(exportDate, exportStatus);
    const conditionText = buildConditionText(exportDate, exportStatus);
    const statusPool = getExportStatusPool(exportStatus);

    if (list.length === 0 && statusPool.length === 0) {
      showToast("No data available for export.");
      return;
    }

    if (window.docx && window.saveAs) {
      try {
        await exportDocx(list, conditionText, exportDate, exportStatus);
        showToast("Word file exported.");
        return;
      } catch (error) {
        console.error("exportDocx error", error);
      }
    }

    exportLegacyDoc(list, conditionText, exportDate, exportStatus);
    showToast("Exported via compatibility Word mode.");
  }

  async function handleSaveData() {
    if (canUseFileSystemAccessApi()) {
      try {
        let handle = state.autoSaveFileHandle;
        if (!handle) {
          handle = await window.showSaveFilePicker(buildBackupPickerOptions());
        }
        const ok = await writeBackupToFileHandle(handle, true);
        if (!ok) {
          showToast("儲存失敗，請確認檔案權限。");
          return;
        }
        state.autoSaveFileHandle = handle;
        state.lastAutoSaveAt = new Date().toISOString();
        updateAutoSaveStatus();
        showToast("已儲存並覆蓋原本檔案。");
        return;
      } catch (error) {
        if (error && error.name === "AbortError") {
          return;
        }
        console.error("handleSaveData (overwrite) error", error);
      }
    }
    try {
      const filename = buildBackupFileName();
      downloadBackupFile(filename);
      showToast("已下載備份檔。");
    } catch (error) {
      console.error("handleSaveData error", error);
      showToast("儲存資料失敗。");
    }
  }

  async function handlePickAutoSaveFile() {
    if (!canUseFileSystemAccessApi()) {
      showToast("目前瀏覽器不支援直接存到本機檔案。");
      return;
    }
    try {
      const handle = await window.showSaveFilePicker(buildBackupPickerOptions());
      const ok = await writeBackupToFileHandle(handle, true);
      if (!ok) {
        showToast("無法寫入自動儲存檔案。");
        return;
      }
      state.autoSaveFileHandle = handle;
      state.lastAutoSaveAt = new Date().toISOString();
      updateAutoSaveStatus();
      showToast("已啟用自動儲存（每10秒）與自動讀取（每30秒）。");
    } catch (error) {
      if (error && error.name === "AbortError") {
        return;
      }
      console.error("handlePickAutoSaveFile error", error);
      showToast("選取自動儲存檔案失敗。");
    }
  }

  function startAutoSaveLoop() {
    if (state.autoSaveTimer) {
      clearInterval(state.autoSaveTimer);
    }
    state.autoSaveTimer = setInterval(function () {
      autoSaveToSelectedFile();
    }, AUTO_SAVE_INTERVAL_MS);
  }

  function startAutoLoadLoop() {
    if (state.autoLoadTimer) {
      clearInterval(state.autoLoadTimer);
    }
    state.autoLoadTimer = setInterval(function () {
      autoLoadFromSelectedFile();
    }, AUTO_LOAD_INTERVAL_MS);
  }

  async function autoSaveToSelectedFile() {
    if (!state.autoSaveFileHandle) {
      return;
    }
    const ok = await writeBackupToFileHandle(state.autoSaveFileHandle, false);
    if (!ok) {
      updateAutoSaveStatus();
      return;
    }
    state.lastAutoSaveAt = new Date().toISOString();
    updateAutoSaveStatus();
  }

  async function autoLoadFromSelectedFile() {
    if (!state.autoSaveFileHandle || typeof state.autoSaveFileHandle.getFile !== "function") {
      return;
    }
    try {
      const file = await state.autoSaveFileHandle.getFile();
      const raw = await file.text();
      if (!raw) {
        return;
      }
      const parsed = JSON.parse(raw);
      const imported = parseBackupPayload(parsed);
      if (!imported) {
        return;
      }
      const changed = applyImportedBackup(imported, true);
      if (changed) {
        showToast("已自動讀取最新資料。");
      }
    } catch (error) {
      console.error("autoLoadFromSelectedFile error", error);
    }
  }

  function canUseFileSystemAccessApi() {
    return typeof window.showSaveFilePicker === "function" && window.isSecureContext;
  }

  async function writeBackupToFileHandle(handle, allowPermissionPrompt) {
    if (!handle || typeof handle.createWritable !== "function") {
      return false;
    }
    const granted = await ensureFileHandlePermission(handle, allowPermissionPrompt);
    if (!granted) {
      return false;
    }
    try {
      const json = JSON.stringify(buildBackupPayload(), null, 2);
      const writable = await handle.createWritable();
      await writable.write(json);
      await writable.close();
      return true;
    } catch (error) {
      console.error("writeBackupToFileHandle error", error);
      return false;
    }
  }

  function buildBackupPickerOptions() {
    return {
      suggestedName: "handover_autosave.json",
      types: [
        {
          description: "JSON",
          accept: {
            "application/json": [".json"],
          },
        },
      ],
    };
  }

  async function ensureFileHandlePermission(handle, allowPermissionPrompt) {
    if (!handle || typeof handle.queryPermission !== "function") {
      return true;
    }
    try {
      const query = await handle.queryPermission({ mode: "readwrite" });
      if (query === "granted") {
        return true;
      }
      if (!allowPermissionPrompt || typeof handle.requestPermission !== "function") {
        return false;
      }
      const request = await handle.requestPermission({ mode: "readwrite" });
      return request === "granted";
    } catch (error) {
      return false;
    }
  }

  function updateAutoSaveStatus() {
    if (!els.autoSaveStatus) {
      return;
    }
    if (!state.autoSaveFileHandle) {
      els.autoSaveStatus.textContent = "自動儲存：未啟用（10秒存檔 / 30秒讀取）";
      return;
    }
    if (!state.lastAutoSaveAt) {
      els.autoSaveStatus.textContent = "自動儲存：每10秒（自動讀取每30秒）";
      return;
    }
    els.autoSaveStatus.textContent = "自動儲存：每10秒（上次 " + formatDateTime(state.lastAutoSaveAt) + "，自動讀取每30秒）";
  }

  function handleLoadDataClick() {
    if (!els.loadDataInput) {
      return;
    }
    els.loadDataInput.value = "";
    els.loadDataInput.click();
  }

  async function handleLoadDataChange(event) {
    const input = event && event.target ? event.target : null;
    const file = input && input.files && input.files[0] ? input.files[0] : null;
    if (!file) {
      return;
    }
    try {
      const raw = await file.text();
      const parsed = JSON.parse(raw);
      const imported = parseBackupPayload(parsed);
      if (!imported) {
        showToast("資料格式不正確。");
        return;
      }
      const changed = applyImportedBackup(imported, false);
      if (!changed) {
        showToast("資料已是最新。");
      }
    } catch (error) {
      console.error("handleLoadDataChange error", error);
      showToast("讀取資料失敗。");
    } finally {
      if (input) {
        input.value = "";
      }
    }
  }

  function downloadBackupFile(filename) {
    const payload = buildBackupPayload();
    const json = JSON.stringify(payload, null, 2);
    const blob = new Blob([json], { type: "application/json;charset=utf-8" });
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

  function buildBackupPayload() {
    return {
      type: BACKUP_TYPE,
      version: BACKUP_VERSION,
      exportedAt: new Date().toISOString(),
      tasks: state.tasks.slice(),
      todayOverview: {
        checkin: state.todayOverview.checkin || "",
        checkout: state.todayOverview.checkout || "",
        occupancy: state.todayOverview.occupancy || "",
      },
    };
  }

  function applyImportedBackup(imported, silent) {
    if (!imported || typeof imported !== "object") {
      return false;
    }
    const tasksChanged = JSON.stringify(state.tasks) !== JSON.stringify(imported.tasks);
    const overviewChanged = JSON.stringify(state.todayOverview) !== JSON.stringify(imported.todayOverview);
    if (!tasksChanged && !overviewChanged) {
      return false;
    }

    state.tasks = imported.tasks;
    state.todayOverview = imported.todayOverview;
    if (state.editingTaskId) {
      resetTaskForm();
    }
    saveTasks();
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

    if (Array.isArray(payload)) {
      rawTasks = payload;
    } else if (payload && typeof payload === "object") {
      rawTasks = Array.isArray(payload.tasks) ? payload.tasks : [];
      rawTodayOverview =
        payload.todayOverview && typeof payload.todayOverview === "object" ? payload.todayOverview : {};
    } else {
      return null;
    }

    const tasks = rawTasks.map(normalizeTask).filter(Boolean).sort(sortByDueTime);
    return {
      tasks: tasks,
      todayOverview: {
        checkin: normalizeTodayOverviewValue(rawTodayOverview.checkin),
        checkout: normalizeTodayOverviewValue(rawTodayOverview.checkout),
        occupancy: normalizeOccupancyRateValue(rawTodayOverview.occupancy),
      },
    };
  }

  function buildBackupFileName() {
    const now = new Date();
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    return "handover_backup_" + toDateKey(now) + "_" + hh + mm + ss + ".json";
  }

  async function exportDocx(tasks, conditionText, exportDate, exportStatus) {
    const docxLib = window.docx;
    const Document = docxLib.Document;
    const Packer = docxLib.Packer;
    const Paragraph = docxLib.Paragraph;
    const HeadingLevel = docxLib.HeadingLevel;
    const Table = docxLib.Table;
    const TableCell = docxLib.TableCell;
    const TableRow = docxLib.TableRow;
    const WidthType = docxLib.WidthType;
    const AlignmentType = docxLib.AlignmentType;
    const BorderStyle = docxLib.BorderStyle;

    const statusPool = getExportStatusPool(exportStatus);
    const sections = buildExportSections(tasks, statusPool, exportDate);

    const rows = [
      createExportSectionRow("Attention to All\nNotices", sections.attention, {
        Paragraph: Paragraph,
        TableCell: TableCell,
        TableRow: TableRow,
        WidthType: WidthType,
        AlignmentType: AlignmentType,
        BorderStyle: BorderStyle,
      }),
      createExportSectionRow("Daily Briefing\nDaily Tasks", sections.daily, {
        Paragraph: Paragraph,
        TableCell: TableCell,
        TableRow: TableRow,
        WidthType: WidthType,
        AlignmentType: AlignmentType,
        BorderStyle: BorderStyle,
      }),
      createExportSectionRow("Banquets & Sales\nEvents", sections.banquets, {
        Paragraph: Paragraph,
        TableCell: TableCell,
        TableRow: TableRow,
        WidthType: WidthType,
        AlignmentType: AlignmentType,
        BorderStyle: BorderStyle,
      }),
      createExportSectionRow("Future Follow Up\nUpcoming Tasks", sections.future, {
        Paragraph: Paragraph,
        TableCell: TableCell,
        TableRow: TableRow,
        WidthType: WidthType,
        AlignmentType: AlignmentType,
        BorderStyle: BorderStyle,
      }),
      createExportMergedRow(["Internal Transfer Items"], {
        Paragraph: Paragraph,
        TableCell: TableCell,
        TableRow: TableRow,
        WidthType: WidthType,
        AlignmentType: AlignmentType,
        BorderStyle: BorderStyle,
        center: true,
      }),
      createExportMergedRow(buildTransferFyiLines(sections.transfer, sections.fyi), {
        Paragraph: Paragraph,
        TableCell: TableCell,
        TableRow: TableRow,
        WidthType: WidthType,
        AlignmentType: AlignmentType,
        BorderStyle: BorderStyle,
        center: false,
      }),
    ];

    const document = new Document({
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ text: "Handover Report", heading: HeadingLevel.HEADING_1 }),
            new Paragraph({ text: "Export Time: " + formatDateTime(new Date().toISOString()) }),
            new Paragraph({ text: "Filters: " + conditionText }),
            new Paragraph({ text: "" }),
            new Table({
              width: {
                size: 100,
                type: WidthType.PERCENTAGE,
              },
              columnWidths: [1900, 9000],
              borders: createExportBorders(BorderStyle),
              rows: rows,
            }),
            new Paragraph({ text: "" }),
            new Paragraph({ text: buildArrDepOccLine(exportDate) }),
          ],
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
          }),
        }),
        new tools.TableCell({
          width: {
            size: 9000,
            type: tools.WidthType.DXA,
          },
          borders: createExportBorders(tools.BorderStyle),
          children: createExportParagraphs(buildExportTaskLines(tasks), tools.Paragraph, {
            alignment: tools.AlignmentType.LEFT,
          }),
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
          }),
        }),
      ],
    });
  }

  function createExportParagraphs(lines, Paragraph, options) {
    const list = Array.isArray(lines) ? lines.filter(Boolean) : [];
    const align = options && options.alignment ? options.alignment : undefined;
    if (list.length === 0) {
      return [new Paragraph({ text: "(none)", alignment: align })];
    }
    return list.map(function (line) {
      return new Paragraph({
        text: String(line),
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

  function buildExportTaskLines(tasks) {
    if (!Array.isArray(tasks) || tasks.length === 0) {
      return [];
    }
    return tasks.map(formatExportTaskLine);
  }

  function formatExportTaskLine(task) {
    const categoryText = task.subcategory ? task.category + "/" + task.subcategory : task.category;
    const statusText = task.status === "done" ? "Done" : "Pending";
    const completedByText = task.status === "done" && task.completedBy ? " | Completed By: " + task.completedBy : "";
    const descText = task.description ? " | " + task.description : "";
    return (
      "- " +
      (task.title || "-") +
      " | " +
      categoryText +
      " | " +
      formatDueDisplay(task) +
      " | Owner: " +
      (task.owner || "-") +
      " | " +
      statusText +
      completedByText +
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
    const dateLabel = formatExportDateLabel(exportDate || toDateKey(new Date()));
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

  function exportLegacyDoc(tasks, conditionText, exportDate, exportStatus) {
    const statusPool = getExportStatusPool(exportStatus);
    const sections = buildExportSections(tasks, statusPool, exportDate);

    function toHtmlLines(lines) {
      if (!Array.isArray(lines) || lines.length === 0) {
        return "(none)";
      }
      return lines
        .map(function (line) {
          return escapeHtml(line);
        })
        .join("<br>");
    }

    const html =
      "<html><head><meta charset='utf-8'><style>" +
      "body{font-family:'DFKai-SB','Noto Serif TC',serif;padding:16px;color:#111;}" +
      "h2{margin:0 0 8px 0;}" +
      "p{margin:4px 0;}" +
      "table{width:100%;border-collapse:collapse;table-layout:fixed;}" +
      "td{border:1px solid #000;padding:8px;vertical-align:top;line-height:1.45;}" +
      ".left{width:160px;text-align:center;font-weight:700;}" +
      ".merge-title{text-align:center;font-weight:700;background:#efefef;}" +
      ".arrdepocc{margin-top:12px;font-weight:700;}" +
      "</style></head><body>" +
      "<h2>Handover Report</h2>" +
      "<p>Export Time: " +
      escapeHtml(formatDateTime(new Date().toISOString())) +
      "</p>" +
      "<p>Filters: " +
      escapeHtml(conditionText) +
      "</p>" +
      "<table>" +
      "<tr><td class='left'>Attention to All<br>Notices</td><td>" +
      toHtmlLines(buildExportTaskLines(sections.attention)) +
      "</td></tr>" +
      "<tr><td class='left'>Daily Briefing<br>Daily Tasks</td><td>" +
      toHtmlLines(buildExportTaskLines(sections.daily)) +
      "</td></tr>" +
      "<tr><td class='left'>Banquets &amp; Sales<br>Events</td><td>" +
      toHtmlLines(buildExportTaskLines(sections.banquets)) +
      "</td></tr>" +
      "<tr><td class='left'>Future Follow Up<br>Upcoming Tasks</td><td>" +
      toHtmlLines(buildExportTaskLines(sections.future)) +
      "</td></tr>" +
      "<tr><td colspan='2' class='merge-title'>Internal Transfer Items</td></tr>" +
      "<tr><td colspan='2'>" +
      toHtmlLines(buildTransferFyiLines(sections.transfer, sections.fyi)) +
      "</td></tr>" +
      "</table>" +
      "<p class='arrdepocc'>" +
      escapeHtml(buildArrDepOccLine(exportDate)) +
      "</p>" +
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
    const datePart = exportDate || state.queryDate || toDateKey(new Date());
    return "工作交接清單_" + datePart + "." + ext;
  }

  function normalizeQueryStatus(raw) {
    const value = String(raw || "").trim();
    if (value === "pending" || value === "done") {
      return value;
    }
    return "all";
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

  function buildQueryConditionText() {
    return buildConditionText(state.queryDate, state.queryStatus);
  }

  function buildConditionText(dateValue, statusValue) {
    const dateText = dateValue ? dateValue.replace(/-/g, "/") : "全部日期";
    return dateText + " / " + getQueryStatusLabel(statusValue);
  }

  function getTasksByDateAndStatus(dateValue, statusValue) {
    let list = state.tasks.slice().sort(sortForTaskTable);
    if (dateValue) {
      list = list.filter(function (task) {
        const startAt = getTaskStartAt(task);
        return startAt && toDateKey(new Date(startAt)) === dateValue;
      });
    }
    if (statusValue !== "all") {
      list = list.filter(function (task) {
        return task.status === statusValue;
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

  function loadTodayOverview() {
    const defaults = {
      checkin: "",
      checkout: "",
      occupancy: "",
    };
    try {
      const raw = localStorage.getItem(TODAY_OVERVIEW_KEY);
      if (!raw) {
        return defaults;
      }
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object") {
        return defaults;
      }
      return {
        checkin: normalizeTodayOverviewValue(parsed.checkin),
        checkout: normalizeTodayOverviewValue(parsed.checkout),
        occupancy: normalizeOccupancyRateValue(parsed.occupancy),
      };
    } catch (error) {
      console.error("loadTodayOverview error", error);
      return defaults;
    }
  }

  function saveTodayOverview() {
    localStorage.setItem(TODAY_OVERVIEW_KEY, JSON.stringify(state.todayOverview));
  }

  function normalizeTodayOverviewValue(raw) {
    return String(raw || "").trim();
  }

  function normalizeOccupancyRateValue(raw) {
    const value = String(raw || "").trim();
    if (!value) {
      return "";
    }
    const number = Number(value);
    if (!Number.isFinite(number)) {
      return "";
    }
    const clamped = Math.min(100, Math.max(0, number));
    const rounded = Math.round(clamped * 100) / 100;
    const normalized = rounded.toFixed(2);
    return normalized.replace(/\.00$/, "").replace(/(\.\d)0$/, "$1");
  }

  function buildId() {
    if (window.crypto && typeof window.crypto.randomUUID === "function") {
      return window.crypto.randomUUID();
    }
    return "task_" + Date.now() + "_" + Math.random().toString(16).slice(2);
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
    return year + "-" + month + "-" + day + "T" + hour + ":" + minute;
  }

  function toDateKey(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return year + "-" + month + "-" + day;
  }

  function parseTaskTimeRange(startInput, endInput, allDay) {
    const startText = String(startInput || "").trim();
    const endText = String(endInput || "").trim();
    const startMs = parseTaskTimeInput(startText, allDay, "start");
    const endMs = parseTaskTimeInput(endText, allDay, "end");

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

  function parseTaskTimeInput(value, allDay, part) {
    const text = String(value || "").trim();
    if (!text) {
      return Number.NaN;
    }
    if (allDay) {
      return new Date(text + (part === "end" ? "T23:59" : "T00:00")).getTime();
    }
    return new Date(text).getTime();
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

