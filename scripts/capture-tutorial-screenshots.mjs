#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { chromium, devices } from "playwright";

const APP_URL = "https://spider10632.github.io/handover-assistant-v06/";
const OUTPUT_DIR = path.resolve(process.cwd(), "assets/tutorial/full");
const LEGACY_TUTORIAL_DIR = path.resolve(process.cwd(), "assets/tutorial");
const QA_PREFIX = "QA_TEST_TUTORIAL_";

const DESKTOP_VIEWPORT = { width: 1720, height: 1120 };

async function main() {
  await fs.mkdir(OUTPUT_DIR, { recursive: true });
  await captureDesktopShots();
  await captureMobileShot();
  await copyPwaInstallShots();
  console.log("[capture] tutorial screenshots completed");
}

function formatDateForInput(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function formatHm(date) {
  const h = String(date.getHours()).padStart(2, "0");
  const m = String(date.getMinutes()).padStart(2, "0");
  return `${h}${m}`;
}

async function openAndLogin(page) {
  page.setDefaultTimeout(9000);
  await page.goto(APP_URL, { waitUntil: "domcontentloaded" });
  await page.waitForTimeout(900);

  const gateHiddenAtStart = await page.locator("#password-gate").evaluate((el) => el.classList.contains("hidden"));
  if (gateHiddenAtStart) {
    await page.evaluate(() => {
      try {
        localStorage.clear();
        sessionStorage.clear();
      } catch (error) {}
    });
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.waitForTimeout(900);
  }

  const hasServerInput = await page.locator("#server-input").count();
  if (!hasServerInput) {
    return;
  }

  await page.waitForSelector("#server-input", { timeout: 15000 });

  await page.screenshot({
    path: path.join(OUTPUT_DIR, "01-login-server-password.png"),
    fullPage: true,
  });

  const passwords = ["testmanager", "testorder"];
  let authed = false;
  for (const password of passwords) {
    const submitVisible = await page.locator("#password-submit").isVisible().catch(() => false);
    if (!submitVisible) {
      authed = await page.locator("#password-gate").evaluate((el) => el.classList.contains("hidden"));
      if (authed) {
        break;
      }
    }
    await page.fill("#server-input", "test");
    await page.fill("#password-input", password);
    await page.click("#password-submit");
    await page.waitForTimeout(1500);
    authed = await page.locator("#password-gate").evaluate((el) => el.classList.contains("hidden"));
    if (authed) {
      break;
    }
  }

  if (!authed) {
    throw new Error("login failed for test server");
  }

  await page.waitForSelector("#task-form", { timeout: 15000 });
  if (await page.locator("#ui-language-select").count()) {
    await page.selectOption("#ui-language-select", "zh");
  }
  await page.waitForTimeout(800);
}

async function ensurePanelExpanded(page, toggleSelector) {
  const toggle = page.locator(toggleSelector);
  if (!(await toggle.count())) {
    return;
  }
  const expanded = await toggle.getAttribute("aria-expanded");
  if (expanded === "false") {
    await toggle.click();
    await page.waitForTimeout(350);
  }
}

async function fillTask(page, data) {
  await page.click("#clear-form-btn");
  await page.waitForTimeout(200);

  await page.selectOption("#task-category", { label: "包裹代收" });
  await page.waitForFunction(() => {
    const sub = document.querySelector("#task-subcategory");
    return sub && !sub.disabled && sub.options.length > 1;
  });
  await page.selectOption("#task-subcategory", { label: "散客" });

  await page.fill("#task-title", data.title);
  await page.fill("#task-owner", data.owner || "tutorial-bot");
  await page.fill("#task-description", data.description || "");

  await page.fill("#task-start-date", data.startDate || "");
  await page.fill("#task-start-at", data.startAt || "");
  await page.fill("#task-end-date", data.endDate || "");
  await page.fill("#task-end-at", data.endAt || "");

  const allDay = page.locator("#all-day-btn");
  const allDayPressed = (await allDay.getAttribute("aria-pressed")) === "true";
  if (Boolean(data.allDay) !== allDayPressed) {
    await allDay.click();
  }

  const pin = page.locator("#task-pinned");
  if ((await pin.isChecked()) !== Boolean(data.pinned)) {
    await pin.click();
  }

  if (data.assignee) {
    const option = page.locator(`#task-assignee option[value="${data.assignee}"]`);
    if (await option.count()) {
      await page.selectOption("#task-assignee", data.assignee);
    }
  } else {
    await page.selectOption("#task-assignee", "");
  }

  await page.click("#add-task-btn");
  await page.waitForTimeout(500);
}

async function screenshotElement(page, selector, filename) {
  const target = page.locator(selector).first();
  await target.scrollIntoViewIfNeeded();
  await page.waitForTimeout(250);
  await target.screenshot({
    path: path.join(OUTPUT_DIR, filename),
  });
}

async function screenshotTaskRow(page, title, filename) {
  const row = page.locator("#task-table-body tr", { hasText: title }).first();
  if (!(await row.count())) {
    return false;
  }
  await row.scrollIntoViewIfNeeded();
  await page.waitForTimeout(200);
  await row.screenshot({
    path: path.join(OUTPUT_DIR, filename),
  });
  return true;
}

async function captureDesktopShots() {
  const browser = await chromium.launch({ headless: true, channel: "chrome" });
  const page = await browser.newPage({ viewport: DESKTOP_VIEWPORT });
  const log = (msg) => console.log("[capture][desktop]", msg);
  try {
    await openAndLogin(page);
    log("logged in");

  const now = new Date();
  const today = formatDateForInput(now);
  const tomorrowDate = new Date(now.getTime() + 24 * 60 * 60 * 1000);
  const tomorrow = formatDateForInput(tomorrowDate);
  const upcoming = new Date(now.getTime() + 25 * 60 * 1000);
  const overdue = new Date(now.getTime() - 12 * 60 * 1000);
  const future = new Date(now.getTime() + 120 * 60 * 1000);

  await fillTask(page, {
    title: `${QA_PREFIX}PENDING`,
    description: "教學圖用：待處理示例",
    startDate: today,
    startAt: formatHm(future),
    endDate: today,
    endAt: "",
    assignee: "teststaff",
    pinned: false,
  });
  log("created sample tasks");

  await fillTask(page, {
    title: `${QA_PREFIX}INPROGRESS`,
    description: "教學圖用：處理中示例",
    startDate: today,
    startAt: formatHm(future),
    endDate: today,
    endAt: "",
    assignee: "teststaff",
  });

  await fillTask(page, {
    title: `${QA_PREFIX}DONE`,
    description: "教學圖用：完成示例",
    startDate: today,
    startAt: formatHm(future),
    endDate: today,
    endAt: "",
    assignee: "teststaff",
  });

  await fillTask(page, {
    title: `${QA_PREFIX}PINNED`,
    description: "教學圖用：置頂示例",
    startDate: today,
    startAt: "1700",
    endDate: today,
    endAt: "1800",
    pinned: true,
    assignee: "teststaff",
  });

  await fillTask(page, {
    title: `${QA_PREFIX}OVERDUE`,
    description: "教學圖用：指派監視示例",
    startDate: today,
    startAt: formatHm(overdue),
    endDate: today,
    endAt: "",
    assignee: "teststaff",
  });

  await fillTask(page, {
    title: `${QA_PREFIX}UPCOMING`,
    description: "教學圖用：下一件待辦提醒示例",
    startDate: today,
    startAt: formatHm(upcoming),
    endDate: today,
    endAt: "",
    assignee: "teststaff",
  });

  await page.waitForTimeout(900);

  await screenshotElement(page, ".top-panels > section.panel", "02-today-overview-kpi.png");
  await screenshotElement(page, ".today-pinned-zone", "03-pinned-zone.png");
  await screenshotElement(page, "#task-form", "04-create-task-basic-fields.png");
  log("captured top and form");

  await page.click("#time-now-btn");
  await screenshotElement(page, ".due-input-row", "05-time-range-cross-day-now-allday.png");
  await screenshotElement(page, ".task-assignee-field", "06-assign-to-user.png");
  log("captured time and assignee");

  await page.fill("#query-date", today);
  await page.click("#search-btn");
  await ensurePanelExpanded(page, "#toggle-task-list-panel");
  await page.waitForTimeout(700);

  const inprogressRow = page.locator("#task-table-body tr", { hasText: `${QA_PREFIX}INPROGRESS` }).first();
  if (await inprogressRow.count()) {
    await inprogressRow.locator('[data-action="progress"]').click();
    await page.waitForTimeout(450);
  }

  const doneRow = page.locator("#task-table-body tr", { hasText: `${QA_PREFIX}DONE` }).first();
  if (await doneRow.count()) {
    await doneRow.locator('[data-action="toggle"]').click();
    await page.waitForTimeout(450);
  }

  await screenshotElement(page, ".query-bar", "07-query-filters.png");
  const pendingCaptured = await screenshotTaskRow(page, `${QA_PREFIX}PENDING`, "08-result-actions-pending.png");
  if (!pendingCaptured) {
    await screenshotElement(page, "#task-list-panel .table-wrap", "08-result-actions-pending.png");
  }
  await screenshotElement(page, "#task-list-panel .table-wrap", "09-result-actions-done-progress.png");
  const translatedCaptured = await screenshotTaskRow(page, `${QA_PREFIX}PENDING`, "10-translation-show-original.png");
  if (!translatedCaptured) {
    await screenshotElement(page, "#task-list-panel .table-wrap", "10-translation-show-original.png");
  }
  log("captured query and result");

  await ensurePanelExpanded(page, "#toggle-assignment-watch-panel");
  await page.waitForTimeout(700);
  await screenshotElement(page, "#assignment-watch-panel .panel-body", "11-assignment-watch-panel.png");
  log("captured assignment watch");

  await page.click("#team-status-toggle");
  await page.waitForTimeout(600);
  await screenshotElement(page, "#team-status-panel", "12-team-panel-collapsed.png");

  const teamCard = page.locator("#team-status-list [data-team-user]").first();
  if (await teamCard.count()) {
    await teamCard.click();
    await page.waitForTimeout(450);
  }
  await screenshotElement(page, "#team-status-panel", "13-team-panel-expanded-items.png");
  log("captured team panel");

    await screenshotElement(page, "#upcoming-board", "14-upcoming-reminder-panel.png");
    await screenshotElement(page, ".export-panel", "15-export-word-excel.png");
    log("captured upcoming and export");
  } finally {
    await browser.close();
  }
}

async function captureMobileShot() {
  const browser = await chromium.launch({ headless: true, channel: "chrome" });
  try {
    const context = await browser.newContext({
      ...devices["iPhone 13"],
      locale: "zh-TW",
      timezoneId: "Asia/Taipei",
    });
    const page = await context.newPage();
    await openAndLogin(page);
    console.log("[capture][mobile] logged in");
    await page.waitForTimeout(1000);

    const addBtn = page.locator("#mobile-add-btn");
    if (await addBtn.count()) {
      await addBtn.click();
      await page.waitForTimeout(350);
    }

    await page.screenshot({
      path: path.join(OUTPUT_DIR, "16-mobile-floating-buttons.png"),
      fullPage: true,
    });
  } finally {
    await browser.close();
  }
}

async function copyPwaInstallShots() {
  const pairs = [
    ["pwa-iphone.png", "17-pwa-install-ios.png"],
    ["pwa-android.png", "18-pwa-install-android-notify.png"],
  ];
  for (const [from, to] of pairs) {
    const src = path.join(LEGACY_TUTORIAL_DIR, from);
    const dst = path.join(OUTPUT_DIR, to);
    await fs.copyFile(src, dst);
  }
}

main().catch((error) => {
  console.error("[capture] failed:", error && error.message ? error.message : error);
  process.exit(1);
});
