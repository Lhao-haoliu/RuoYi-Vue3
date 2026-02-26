import fs from "node:fs";
import path from "node:path";
import { chromium } from "playwright-core";

const BASE_URL = "http://localhost";
const USERNAME = "admin";
const PASSWORD = "admin123";
const EDGE_PATHS = [
  "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
  "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe"
];

const report = [];
const runtimeErrors = [];

function pickEdgePath() {
  for (const p of EDGE_PATHS) {
    if (fs.existsSync(p)) return p;
  }
  throw new Error("Microsoft Edge executable not found.");
}

function pushResult(page, action, status, detail = "") {
  report.push({ page, action, status, detail });
}

async function waitReady(page) {
  await page.waitForLoadState("domcontentloaded");
  await page.waitForTimeout(400);
}

async function findVisibleButton(page, names) {
  for (const name of names) {
    const btn = page.getByRole("button", { name }).first();
    if (await btn.count()) {
      if (await btn.isVisible()) return btn;
    }
  }
  return null;
}

async function closeDialogIfOpen(page) {
  const dialog = page.locator(".el-dialog:visible").last();
  if (!(await dialog.count())) return;
  const cancelBtn = dialog.getByRole("button", { name: /取消|取\s*消|关闭|Close/i }).first();
  if (await cancelBtn.count()) {
    await cancelBtn.click();
  } else {
    const x = dialog.locator(".el-dialog__headerbtn").first();
    if (await x.count()) await x.click();
  }
  await page.waitForTimeout(250);
}

async function clickAction(page, pageName, actionName, names, options = {}) {
  const {
    expectDialog = false,
    expectConfirm = false,
    allowDisabledSkip = true,
    clickTimeout = 8000
  } = options;
  const btn = await findVisibleButton(page, names);
  if (!btn) {
    pushResult(pageName, actionName, "SKIP", "按钮不存在或无权限");
    return false;
  }
  const disabled = await btn.isDisabled();
  if (disabled && allowDisabledSkip) {
    pushResult(pageName, actionName, "SKIP", "按钮禁用（通常是未选中行）");
    return false;
  }
  const beforeErrCount = runtimeErrors.length;
  try {
    await btn.click({ timeout: clickTimeout });
    if (expectDialog) {
      await page.waitForSelector(".el-dialog:visible", { timeout: 6000 });
      await closeDialogIfOpen(page);
    } else if (expectConfirm) {
      await page.waitForSelector(".el-message-box:visible", { timeout: 6000 });
      const cancelBtn = page.getByRole("button", { name: /取消|取\s*消/ }).first();
      if (await cancelBtn.count()) {
        await cancelBtn.click();
      }
    } else {
      await page.waitForTimeout(300);
    }
    const afterErrCount = runtimeErrors.length;
    if (afterErrCount > beforeErrCount) {
      pushResult(pageName, actionName, "FAIL", `点击后出现前端异常: ${runtimeErrors[afterErrCount - 1]}`);
      return false;
    }
    pushResult(pageName, actionName, "PASS");
    return true;
  } catch (err) {
    pushResult(pageName, actionName, "FAIL", String(err.message || err));
    return false;
  }
}

async function selectFirstRow(page) {
  const firstCheckbox = page.locator(".el-table__body-wrapper .el-checkbox__input").first();
  if (await firstCheckbox.count()) {
    await firstCheckbox.click();
    await page.waitForTimeout(200);
    return true;
  }
  return false;
}

async function checkSearchReset(page, pageName) {
  try {
    const searchInput = page.locator(".page-container__search-card input:visible").first();
    if (await searchInput.count()) {
      await searchInput.fill("auto-check");
    }
    await clickAction(page, pageName, "搜索", [/搜索|Search/i], { expectDialog: false });
    await clickAction(page, pageName, "重置", [/重置|Reset/i], { expectDialog: false });
  } catch (err) {
    pushResult(pageName, "搜索/重置", "FAIL", String(err.message || err));
  }
}

async function checkUserPage(page) {
  const pageName = "用户管理";
  await page.goto(`${BASE_URL}/system/user`);
  await waitReady(page);
  await checkSearchReset(page, pageName);
  await clickAction(page, pageName, "新建用户", [/新建用户|新增/i], { expectDialog: true });
  await selectFirstRow(page);
  await clickAction(page, pageName, "工具栏修改", [/修改|Edit/i], { expectDialog: true });
  await clickAction(page, pageName, "工具栏删除", [/删除|Delete/i], { expectConfirm: true });
  await clickAction(page, pageName, "工具栏导出", [/导出|Download/i], { expectDialog: false });
  await clickAction(page, pageName, "工具栏导入", [/导入|Import/i], { expectDialog: true });
  await clickAction(page, pageName, "首行编辑", [/编辑/i], { expectDialog: true });

  const more = page.locator(".operation-text-btn:visible").filter({ hasText: "更多" }).first();
  if (!(await more.count())) {
    pushResult(pageName, "更多菜单", "SKIP", "无更多菜单");
    return;
  }
  try {
    await more.click();
    const resetPwd = page.locator(".el-dropdown-menu__item:visible").filter({ hasText: "重置密码" }).first();
    if (await resetPwd.count()) {
      try {
        await resetPwd.click();
        await page.waitForSelector(".el-message-box:visible", { timeout: 4000 });
        await page.getByRole("button", { name: /取消|取\s*消/ }).click();
        pushResult(pageName, "更多-重置密码", "PASS");
      } catch (err) {
        pushResult(pageName, "更多-重置密码", "FAIL", String(err.message || err));
      }
    } else {
      pushResult(pageName, "更多-重置密码", "SKIP", "无权限或按钮不存在");
    }
  } catch (err) {
    pushResult(pageName, "更多菜单", "FAIL", String(err.message || err));
  }
}

async function checkRoleAuthUserPage(page) {
  const pageName = "角色分配用户";
  await waitReady(page);
  await checkSearchReset(page, pageName);
  await clickAction(page, pageName, "添加用户", [/添加用户/i], { expectDialog: true });
  await clickAction(page, pageName, "关闭", [/关闭|返回|Close/i], { expectDialog: false });
}

async function checkRolePage(page) {
  const pageName = "角色管理";
  await page.goto(`${BASE_URL}/system/role`);
  await waitReady(page);
  await checkSearchReset(page, pageName);
  await clickAction(page, pageName, "新建角色", [/New Role|新增/i], { expectDialog: true });
  await selectFirstRow(page);
  await clickAction(page, pageName, "工具栏修改", [/修改|Edit/i], { expectDialog: true });
  await clickAction(page, pageName, "工具栏删除", [/删除|Delete/i], { expectConfirm: true });
  await clickAction(page, pageName, "工具栏导出", [/导出|Download/i], { expectDialog: false });
  await clickAction(page, pageName, "首行编辑", [/编辑|修改/i], { expectDialog: true });
  await clickAction(page, pageName, "首行数据权限", [/数据权限/i], { expectDialog: true });

  const assignBtn = page.locator("button.operation-text-btn", { hasText: "分配用户" }).first();
  if (await assignBtn.count()) {
    try {
      await Promise.all([
        page.waitForURL(/\/system\/role-auth\/user\//, { timeout: 6000 }),
        assignBtn.click()
      ]);
      pushResult(pageName, "首行分配用户", "PASS");
      await checkRoleAuthUserPage(page);
      await page.goto(`${BASE_URL}/system/role`);
      await waitReady(page);
    } catch (err) {
      pushResult(pageName, "首行分配用户", "FAIL", String(err.message || err));
    }
  } else {
    pushResult(pageName, "首行分配用户", "SKIP", "按钮不存在或无权限");
  }
}

async function checkMenuPage(page) {
  const pageName = "菜单管理";
  await page.goto(`${BASE_URL}/system/menu`);
  await waitReady(page);
  await checkSearchReset(page, pageName);
  await clickAction(page, pageName, "新建菜单", [/New Menu|新增/i], { expectDialog: true });
  await clickAction(page, pageName, "首行编辑", [/修改|Edit/i], { expectDialog: true });
}

async function checkDictPage(page) {
  const pageName = "字典管理";
  await page.goto(`${BASE_URL}/system/dict`);
  await waitReady(page);
  await checkSearchReset(page, pageName);
  await clickAction(page, pageName, "新建字典", [/New Dictionary|新增/i], { expectDialog: true });
  await selectFirstRow(page);
  await clickAction(page, pageName, "工具栏修改", [/修改|Edit/i], { expectDialog: true });
  await clickAction(page, pageName, "工具栏删除", [/删除|Delete/i], { expectConfirm: true });
  await clickAction(page, pageName, "工具栏导出", [/导出|Download/i], { expectDialog: false });
  await clickAction(page, pageName, "首行编辑", [/修改|Edit/i], { expectDialog: true });
}

async function checkConfigPage(page) {
  const pageName = "参数管理";
  await page.goto(`${BASE_URL}/system/config`);
  await waitReady(page);
  await checkSearchReset(page, pageName);
  await clickAction(page, pageName, "新建参数", [/New Config|新增/i], { expectDialog: true });
  await selectFirstRow(page);
  await clickAction(page, pageName, "工具栏修改", [/修改|Edit/i], { expectDialog: true });
  await clickAction(page, pageName, "工具栏删除", [/删除|Delete/i], { expectConfirm: true });
  await clickAction(page, pageName, "工具栏导出", [/导出|Download/i], { expectDialog: false });
  await clickAction(page, pageName, "刷新缓存", [/刷新缓存|Refresh/i], { expectDialog: false });
  await clickAction(page, pageName, "首行编辑", [/修改|Edit/i], { expectDialog: true });
}

async function checkNoticePage(page) {
  const pageName = "通知公告";
  await page.goto(`${BASE_URL}/system/notice`);
  await waitReady(page);
  await checkSearchReset(page, pageName);
  await clickAction(page, pageName, "新建公告", [/New Notice|新增/i], { expectDialog: true });
  await selectFirstRow(page);
  await clickAction(page, pageName, "工具栏修改", [/修改|Edit/i], { expectDialog: true });
  await clickAction(page, pageName, "工具栏删除", [/删除|Delete/i], { expectConfirm: true });
  await clickAction(page, pageName, "工具栏导出", [/导出|Download/i], { expectDialog: false });
  await clickAction(page, pageName, "首行编辑", [/修改|Edit/i], { expectDialog: true });
}

async function ensureLogin(page) {
  await page.goto(`${BASE_URL}/login`);
  await waitReady(page);
  if (!page.url().includes("/login")) return;
  const inputs = page.locator("input:visible");
  if ((await inputs.count()) < 2) {
    throw new Error("登录页输入框数量不足，无法自动登录");
  }
  await inputs.nth(0).fill(USERNAME);
  await inputs.nth(1).fill(PASSWORD);
  const loginBtn = await findVisibleButton(page, [/登录|Login/i]);
  if (!loginBtn) throw new Error("未找到登录按钮");
  await loginBtn.click();
  await page.waitForURL((url) => !url.pathname.includes("/login"), { timeout: 15000 });
}

async function main() {
  const edgePath = pickEdgePath();
  const browser = await chromium.launch({
    executablePath: edgePath,
    headless: true
  });
  const context = await browser.newContext();
  const page = await context.newPage();

  page.on("pageerror", (err) => runtimeErrors.push(`pageerror: ${err.message}`));
  page.on("console", (msg) => {
    if (msg.type() === "error") runtimeErrors.push(`console: ${msg.text()}`);
  });

  const safeRun = async (name, fn) => {
    try {
      await fn();
    } catch (err) {
      pushResult(name, "页面检查", "FAIL", String(err.message || err));
    }
  };

  try {
    await ensureLogin(page);
    await safeRun("用户管理", () => checkUserPage(page));
    await safeRun("角色管理", () => checkRolePage(page));
    await safeRun("菜单管理", () => checkMenuPage(page));
    await safeRun("字典管理", () => checkDictPage(page));
    await safeRun("参数管理", () => checkConfigPage(page));
    await safeRun("通知公告", () => checkNoticePage(page));
  } finally {
    await browser.close();
  }

  const outPath = path.resolve("scripts/ui-page-acceptance-report.json");
  fs.writeFileSync(outPath, JSON.stringify({
    generatedAt: new Date().toISOString(),
    runtimeErrors,
    report
  }, null, 2), "utf8");

  const summary = report.reduce((acc, r) => {
    acc[r.status] = (acc[r.status] || 0) + 1;
    return acc;
  }, {});
  console.log(`REPORT_FILE=${outPath}`);
  console.log(`SUMMARY=${JSON.stringify(summary)}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
