import { spawn } from "node:child_process";
import fsSync from "node:fs";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";

const root = fileURLToPath(new URL("..", import.meta.url));
const args = process.argv.slice(2);

function readArg(name, fallback = "") {
  const index = args.indexOf(name);
  if (index >= 0 && args[index + 1]) return args[index + 1];
  const prefix = `${name}=`;
  const inline = args.find((arg) => arg.startsWith(prefix));
  return inline ? inline.slice(prefix.length) : fallback;
}

const options = {
  screenshots: args.includes("--screenshots") || args.includes("--screenshot"),
  headed: args.includes("--headed"),
  keepServer: args.includes("--keep-server"),
  host: readArg("--host", "127.0.0.1"),
  port: Number(readArg("--port", process.env.PLAYWRIGHT_PORT || "8765")),
};

const baseUrl = `http://${options.host}:${options.port}`;
const outputDir = path.join(root, "_output", "playwright");
const quoteDataRoot = path.join(root, "_tmp", "playwright-quote-data");

function pythonCommand() {
  if (process.env.PYTHON) return process.env.PYTHON;
  if (process.platform !== "win32") return "python3";
  const bundled = path.join(os.homedir(), ".cache", "codex-runtimes", "codex-primary-runtime", "dependencies", "python", "python.exe");
  if (fsSync.existsSync(bundled)) return bundled;
  return "python";
}

async function healthOk() {
  try {
    const response = await fetch(`${baseUrl}/api/health`, { signal: AbortSignal.timeout(1200) });
    return response.ok;
  } catch {
    return false;
  }
}

async function waitForHealth(timeoutMs = 15000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    if (await healthOk()) return true;
    await new Promise((resolve) => setTimeout(resolve, 300));
  }
  return false;
}

function startServer() {
  const server = spawn(
    pythonCommand(),
    ["webapp/server.py", "--host", options.host, "--port", String(options.port)],
    {
      cwd: root,
      env: { ...process.env, APP_MODE: "local", QUOTE_DATA_ROOT: process.env.QUOTE_DATA_ROOT || quoteDataRoot },
      stdio: ["ignore", "pipe", "pipe"],
      windowsHide: true,
    },
  );
  const output = [];
  const collect = (chunk) => {
    output.push(String(chunk));
    if (output.join("").length > 8000) output.shift();
  };
  server.stdout.on("data", collect);
  server.stderr.on("data", collect);
  return { server, output };
}

async function stopServer(serverInfo) {
  if (!serverInfo || options.keepServer) return;
  if (serverInfo.server.killed) return;
  serverInfo.server.kill();
  await new Promise((resolve) => serverInfo.server.once("exit", resolve));
}

async function screenshot(page, name) {
  if (!options.screenshots) return "";
  await fs.mkdir(outputDir, { recursive: true });
  const filePath = path.join(outputDir, name);
  await page.screenshot({ path: filePath, fullPage: false });
  return filePath;
}

async function expectTopbarPrimaryAction(page, expectedAction) {
  const dashboardVisible = await page.locator("#backToDashboardButton").isVisible();
  const newQuoteVisible = await page.locator("#newQuoteButton").isVisible();
  const settingsVisible = await page.locator("#settingsButton").isVisible();
  if (!settingsVisible) {
    throw new Error("Pricing Reference topbar action should stay visible.");
  }
  if (expectedAction === "dashboard" && (!dashboardVisible || newQuoteVisible)) {
    throw new Error(`Expected Dashboard-only topbar action, found ${JSON.stringify({ dashboardVisible, newQuoteVisible })}.`);
  }
  if (expectedAction === "new-quote" && (!newQuoteVisible || dashboardVisible)) {
    throw new Error(`Expected New Quote-only topbar action, found ${JSON.stringify({ dashboardVisible, newQuoteVisible })}.`);
  }
}

async function dashboardPanelActionMetrics(page, label) {
  await page.locator("#dashboardSelectedSessionPanel").scrollIntoViewIfNeeded();
  const panelBox = await page.locator("#dashboardSelectedSessionPanel").boundingBox();
  const actionBox = await page.locator("#dashboardSelectedSessionPanel .dashboard-selected-actions").boundingBox();
  const firstActionBox = await page.locator("#dashboardSelectedSessionPanel .dashboard-selected-actions button, #dashboardSelectedSessionPanel .dashboard-selected-actions a").first().boundingBox();
  const actionPositions = await page.locator("#dashboardSelectedSessionPanel").evaluate((panel) => {
    const panelRect = panel.getBoundingClientRect();
    const actions = panel.querySelector(".dashboard-selected-actions");
    const actionRect = actions?.getBoundingClientRect();
    const modifyButton = actions?.querySelector('[data-dashboard-panel-action="modify-session"]');
    const deleteButton = actions?.querySelector('[data-dashboard-panel-action="delete-session"], [data-dashboard-panel-action="delete-selected"]');
    const clearButton = actions?.querySelector('[data-dashboard-panel-action="clear-selection"]');
    const modifyRect = modifyButton?.getBoundingClientRect();
    const deleteRect = deleteButton?.getBoundingClientRect();
    const clearRect = clearButton?.getBoundingClientRect();
    return {
      actionX: actionRect ? Math.round(actionRect.x - panelRect.x) : null,
      actionWidth: actionRect ? Math.round(actionRect.width) : null,
      actionTop: actionRect ? Math.round(actionRect.y - panelRect.y) : null,
      modifyTop: modifyRect ? Math.round(modifyRect.y - panelRect.y) : null,
      deleteTop: deleteRect ? Math.round(deleteRect.y - panelRect.y) : null,
      clearTop: clearRect ? Math.round(clearRect.y - panelRect.y) : null,
    };
  });
  if (!panelBox || !actionBox || !firstActionBox) {
    throw new Error(`Dashboard ${label} action footer was not measurable.`);
  }
  if (actionPositions.deleteTop === null || actionPositions.clearTop === null) {
    throw new Error(`Dashboard ${label} shared delete/clear actions were not measurable.`);
  }
  const deleteClearGap = actionPositions.clearTop - actionPositions.deleteTop;
  if (deleteClearGap < 38 || deleteClearGap > 54) {
    throw new Error(`Dashboard ${label} delete/clear spacing is unexpected: ${deleteClearGap}px.`);
  }
  const modifyDeleteGap = actionPositions.modifyTop === null ? null : actionPositions.deleteTop - actionPositions.modifyTop;
  if (modifyDeleteGap !== null && Math.abs(modifyDeleteGap - deleteClearGap) > 4) {
    throw new Error(`Dashboard ${label} modify/delete spacing differs from delete/clear spacing: ${modifyDeleteGap}px vs ${deleteClearGap}px.`);
  }
  const bottomGap = Math.round((panelBox.y + panelBox.height) - (actionBox.y + actionBox.height));
  if (bottomGap < 8 || bottomGap > 36) {
    throw new Error(`Dashboard ${label} action footer is not bottom anchored: ${bottomGap}px gap.`);
  }
  return {
    bottomGap,
    actionViewportTop: Math.round(actionBox.y),
    firstActionTop: Math.round(firstActionBox.y),
    deleteClearGap,
    modifyDeleteGap,
    ...actionPositions,
  };
}

async function installMockProfiles(page) {
  await page.route("**/api/settings/pricing-references/synthetic-exhibition-fixture-pricing**", async (route) => {
    await route.fulfill({
      status: 200,
      contentType: "application/json",
      body: JSON.stringify({
        pricing_reference: {
          id: "synthetic-exhibition-fixture-pricing",
          label: "Synthetic Exhibition Fixture Pricing",
          description: "Test-only pricing reference for the Playwright smoke.",
          source: "local",
          schema_version: 1,
          currency: "SGD",
          tax: { label: "GST", rate: 0.09 },
          item_count: 1,
          items: [{
            id: "synthetic-floor-needle-punch-carpet",
            section: "Floor Design",
            description: "Needle punch carpet in colour",
            unit_hint: "sqm",
            internal_cost: 10,
            markup_multiplier: 1.5,
            remarks: "Synthetic smoke fixture row",
          }],
        },
      }),
    });
  });
  await page.route("**/api/profiles", async (route) => {
    await route.fulfill({
      status: 200,
      contentType: "application/json",
      body: JSON.stringify({
        profiles: [{
          id: "synthetic-exhibition-fixture-template",
          label: "Synthetic Exhibition Fixture Template",
          description: "Test-only profile for the Playwright smoke.",
          default_pricing_reference: "synthetic-exhibition-fixture-pricing",
          default_quote_detail_preset: "synthetic-fixture-default",
          quote_detail_presets: [{
            id: "synthetic-fixture-default",
            name: "Synthetic Fallback Quote Company",
            details: {
              company: {
                name: "Synthetic Fallback Quote Company Pte Ltd",
                header_details: "Synthetic Fallback Quote Company Pte Ltd\n1 Synthetic Way\nSingapore 000001",
                logo_data_url: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
              },
              quote_text: {
                payment_terms: ["70% synthetic deposit upon confirmation."],
                cheque_payee: "Synthetic Fallback Quote Company Pte Ltd",
              },
              signature: {
                company_signatory: "Synthetic Signatory",
                company_title: "Synthetic Title",
                company_date_label: "Date:",
              },
            },
          }],
        }],
        pricing_references: [{
          id: "synthetic-exhibition-fixture-pricing",
          label: "Synthetic Exhibition Fixture Pricing",
          source: "local",
          currency: "SGD",
          tax: { label: "GST", rate: 0.09 },
          item_count: 1,
        }],
        default_profile_id: "synthetic-exhibition-fixture-template",
        default_pricing_reference_id: "synthetic-exhibition-fixture-pricing",
        company_id: "default",
        workspace: {
          company: { id: "default", slug: "default", display_name: "Quote Generator Workspace" },
          workspace: { id: "default", slug: "default", display_name: "Quote Generator Workspace" },
          runtime_dependencies: {},
        },
      }),
    });
  });
}

async function currentQuoteSessionId(page) {
  return page.evaluate(() => {
    try {
      const saved = JSON.parse(window.localStorage.getItem("swooshz_quote_session_v1") || "{}");
      return String(saved.quoteSessionId || "");
    } catch {
      return "";
    }
  });
}

async function dashboardQuoteSessionDetail(page, sessionId) {
  return page.evaluate(async (safeSessionId) => {
    const response = await fetch(`/api/quote-sessions/${encodeURIComponent(safeSessionId)}`);
    return response.json();
  }, sessionId);
}

async function createDashboardSmokeSession(page, suffix, options = {}) {
  return page.evaluate(async ({ suffix, sessionIdPrefix, customerName, projectName }) => {
    const sessionResponse = await fetch("/api/session");
    if (!sessionResponse.ok) throw new Error(`Session bootstrap failed: ${sessionResponse.status}`);
    const session = await sessionResponse.json();
    const headers = { "Content-Type": "application/json" };
    if (session.csrf_token) headers[session.csrf_header || "X-CSRF-Token"] = session.csrf_token;
    const safeSuffix = String(suffix || "session").replace(/[^A-Za-z0-9_-]/g, "-");
    const safePrefix = String(sessionIdPrefix || `playwright-bulk-${safeSuffix}`).replace(/[^A-Za-z0-9_-]/g, "-");
    const customer = customerName === undefined ? "Marina Bay Product Launch" : String(customerName);
    const project = projectName === undefined ? "Orchard Road Pop-up Booth" : String(projectName);
    const response = await fetch("/api/quote-sessions", {
      method: "POST",
      headers,
      body: JSON.stringify({
        session_id: `${safePrefix}-${Date.now()}`,
        customer_summary: {
          customer_name: customer,
          project_name: project,
        },
        quote_company_profile: {
          id: "synthetic-fixture-default",
          display_name: "Demo Quote Company",
        },
        pricing_reference: {
          id: "synthetic-exhibition-fixture-pricing",
          display_name: "Synthetic Exhibition Fixture Pricing",
        },
        commercials: {
          currency: "SGD",
          tax_label: "GST",
          tax_rate: 0.09,
          subtotal: 1200,
          tax_amount: 108,
          grand_total: 1308,
        },
        status: {
          quote_generated: true,
        },
      }),
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(`Quote session fixture failed: ${response.status}`);
    return data.quote_session?.session_id || "";
  }, {
    suffix,
    sessionIdPrefix: options.sessionIdPrefix || "",
    customerName: Object.prototype.hasOwnProperty.call(options, "customerName") ? options.customerName : undefined,
    projectName: Object.prototype.hasOwnProperty.call(options, "projectName") ? options.projectName : undefined,
  });
}

async function verifyConcurrentInitialDraftSaveUsesSingleSession(page) {
  await page.goto(baseUrl, { waitUntil: "domcontentloaded" });
  await page.locator("#quoteDashboardPanel").waitFor({ state: "visible", timeout: 15000 });
  await page.waitForFunction(() => {
    const countText = document.querySelector("#dashboardSessionCount")?.textContent || "";
    const emptyState = document.querySelector("#dashboardEmptyState");
    const sessionList = document.querySelector("#dashboardSessionsList");
    return !/Loading sessions/i.test(countText)
      && ((emptyState && !emptyState.hidden) || (sessionList && !sessionList.hidden));
  }, null, { timeout: 15000 });
  const emptyNewQuoteButton = page.locator("#dashboardEmptyNewQuoteButton:not([disabled])");
  if (await emptyNewQuoteButton.isVisible()) {
    await emptyNewQuoteButton.click();
  } else {
    await page.locator("#newQuoteButton:not([disabled])").click();
  }
  await page.locator("#sampleDetailsButton:not([disabled])").waitFor({ timeout: 15000 });
  await page.locator("#sampleDetailsButton").click();
  await page.locator("#fileList .file-item").first().waitFor({ timeout: 15000 });
  const raceResult = await page.evaluate(async () => {
    setSidePanel("customer", { force: true });
    state.quoteSessionId = "";
    state.quoteSessionDraftSaveStarted = true;
    saveSessionState();
    const beforeData = await fetch("/api/quote-sessions").then((response) => response.json());
    const beforeIds = new Set((beforeData.quote_sessions || []).map((session) => session.session_id));
    const responses = await Promise.all([
      saveQuoteSessionDraftState({ quoteGenerated: false }),
      saveQuoteSessionDraftState({ quoteGenerated: false }),
    ]);
    const afterData = await fetch("/api/quote-sessions").then((response) => response.json());
    const afterIds = (afterData.quote_sessions || []).map((session) => session.session_id);
    const createdIds = afterIds.filter((sessionId) => !beforeIds.has(sessionId));
    await Promise.all(createdIds.map((sessionId) => deleteQuoteSessionRecord(sessionId)));
    window.localStorage.removeItem("swooshz_quote_session_v1");
    return {
      createdIds,
      responseIds: responses.map((session) => session?.session_id || ""),
      stateSessionId: state.quoteSessionId,
    };
  });
  const uniqueCreatedIds = new Set(raceResult.createdIds.filter(Boolean));
  const uniqueResponseIds = new Set(raceResult.responseIds.filter(Boolean));
  if (raceResult.createdIds.length !== 1 || uniqueCreatedIds.size !== 1 || uniqueResponseIds.size !== 1) {
    throw new Error(`Concurrent initial draft saves should create one session, found ${JSON.stringify(raceResult)}.`);
  }
  await page.goto(baseUrl, { waitUntil: "domcontentloaded" });
  await page.locator("#quoteDashboardPanel").waitFor({ state: "visible", timeout: 15000 });
  await page.waitForFunction(() => {
    const countText = document.querySelector("#dashboardSessionCount")?.textContent || "";
    const emptyState = document.querySelector("#dashboardEmptyState");
    const sessionList = document.querySelector("#dashboardSessionsList");
    return !/Loading sessions/i.test(countText)
      && ((emptyState && !emptyState.hidden) || (sessionList && !sessionList.hidden));
  }, null, { timeout: 15000 });
}

async function main() {
  let serverInfo = null;
  const hasExistingServer = await healthOk();
  if (!hasExistingServer) {
    await fs.rm(quoteDataRoot, { recursive: true, force: true });
    serverInfo = startServer();
    if (!(await waitForHealth())) {
      const serverOutput = serverInfo.output.join("").trim();
      await stopServer(serverInfo);
      throw new Error(`Could not start webapp at ${baseUrl}.${serverOutput ? `\n\n${serverOutput}` : ""}`);
    }
  }

  const browser = await chromium.launch({ headless: !options.headed });
  const page = await browser.newPage({ viewport: { width: 1365, height: 768 } });
  const consoleProblems = [];
  const networkProblems = [];
  page.on("console", (message) => {
    if (["error", "warning"].includes(message.type())) consoleProblems.push(`${message.type()}: ${message.text()}`);
  });
  page.on("pageerror", (error) => consoleProblems.push(`pageerror: ${error.message}`));
  page.on("response", (response) => {
    if (response.status() >= 400) networkProblems.push(`${response.status()} ${response.url()}`);
  });

  try {
    await installMockProfiles(page);
    await verifyConcurrentInitialDraftSaveUsesSingleSession(page);
    await page.goto(baseUrl, { waitUntil: "domcontentloaded" });
    await page.getByRole("heading", { name: "Swooshz Quote Generator" }).waitFor();
    await page.locator("#quoteDashboardPanel").waitFor({ state: "visible" });
    await page.getByRole("heading", { name: "Dashboard" }).waitFor();
    const statusFilterLabels = await page.locator("#dashboardStatusFilter option").evaluateAll((options) => (
      options.map((option) => option.textContent?.trim())
    ));
    if (JSON.stringify(statusFilterLabels) !== JSON.stringify(["All statuses", "Draft", "Generated"])) {
      throw new Error(`Unexpected dashboard status filters: ${JSON.stringify(statusFilterLabels)}`);
    }
    await expectTopbarPrimaryAction(page, "new-quote");
    await page.waitForFunction(() => {
      const countText = document.querySelector("#dashboardSessionCount")?.textContent || "";
      const emptyState = document.querySelector("#dashboardEmptyState");
      const sessionList = document.querySelector("#dashboardSessionsList");
      return !/Loading sessions/i.test(countText)
        && ((emptyState && !emptyState.hidden) || (sessionList && !sessionList.hidden));
    }, null, { timeout: 15000 });
    const dashboardShot = await screenshot(page, "dashboard.png");
    const emptyNewQuoteButton = page.locator("#dashboardEmptyNewQuoteButton:not([disabled])");
    if (await emptyNewQuoteButton.isVisible()) {
      await emptyNewQuoteButton.click();
    } else {
      await page.locator("#newQuoteButton:not([disabled])").click();
    }
    await page.locator("#imageIntake").waitFor({ state: "visible" });
    await expectTopbarPrimaryAction(page, "dashboard");
    const homeShot = await screenshot(page, "home.png");

    await page.locator("#topbarBrandButton", { hasText: "Swooshz Quote Generator" }).click();
    await page.locator("#quoteDashboardPanel").waitFor({ state: "visible", timeout: 15000 });
    await page.locator("#dashboardEmptyState").waitFor({ state: "visible", timeout: 15000 });
    const blankDraftRows = await page.locator(".dashboard-session-card").count();
    if (blankDraftRows !== 0) {
      throw new Error(`Blank quote draft should be discarded on dashboard return, found ${blankDraftRows} dashboard rows.`);
    }
    await page.locator("#newQuoteButton:not([disabled])").click();
    await page.locator("#imageIntake").waitFor({ state: "visible" });
    await expectTopbarPrimaryAction(page, "dashboard");

    await page.locator("#sampleDetailsButton:not([disabled])").waitFor({ timeout: 15000 });
    await page.locator("#sampleDetailsButton").click();
    await page.locator("#fileList .file-item").first().waitFor({ timeout: 15000 });
    const preCustomerQuoteSessionId = await currentQuoteSessionId(page);
    if (preCustomerQuoteSessionId) {
      throw new Error(`Expected dashboard draft saving to wait until Next: Customer, found ${preCustomerQuoteSessionId}.`);
    }
    await page.setViewportSize({ width: 520, height: 720 });
    await page.evaluate(() => window.scrollTo(0, document.documentElement.scrollHeight));
    await page.waitForFunction(() => window.scrollY > 40, null, { timeout: 15000 });
    await page.locator("#sideNextButton", { hasText: "Next: Customer" }).click();
    await page.locator("#customerDetailsPanel").waitFor({ state: "visible", timeout: 15000 });
    await page.waitForTimeout(50);
    const postNextScrollY = await page.evaluate(() => window.scrollY);
    if (postNextScrollY > 2) {
      throw new Error(`Next: Customer should reset the page scroll to the top, found scrollY=${postNextScrollY}.`);
    }
    await page.setViewportSize({ width: 1365, height: 768 });
    await page.waitForFunction(() => {
      try {
        const saved = JSON.parse(window.localStorage.getItem("swooshz_quote_session_v1") || "{}");
        return Boolean(saved.quoteSessionId);
      } catch {
        return false;
      }
    }, null, { timeout: 15000 });
    await page.locator('.rail-button[data-side-panel="quote_company"]:not([disabled])').waitFor({ timeout: 15000 });
    await page.locator('.rail-button[data-side-panel="quote_company"]').click();
    await page.locator("#quoteCompanyPanel").waitFor({ state: "visible" });
    const samplePresetValue = await page.locator("#presetSelect").inputValue();
    if (samplePresetValue !== "profile:synthetic-fixture-default") {
      throw new Error(`Expected sample to select the synthetic fixture preset, found ${samplePresetValue}.`);
    }
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.getByRole("heading", { name: "Swooshz Quote Generator" }).waitFor();
    await page.locator("#panel-analysis.is-active").waitFor({ state: "visible", timeout: 15000 });
    await page.locator("#quoteCompanyPanel.is-active").waitFor({ state: "visible", timeout: 15000 });
    await expectTopbarPrimaryAction(page, "dashboard");
    const refreshedActiveRailTexts = await page.locator(".rail-button.is-active").evaluateAll((buttons) => (
      buttons.map((button) => button.textContent?.trim() || "")
    ));
    if (refreshedActiveRailTexts.length !== 1 || refreshedActiveRailTexts[0] !== "Quote Company") {
      throw new Error(`Expected refresh to restore the last quote menu, found ${JSON.stringify(refreshedActiveRailTexts)}.`);
    }
    const restoredQuoteSessionId = await currentQuoteSessionId(page);
    if (!restoredQuoteSessionId) {
      throw new Error("Expected refresh recovery to keep the current quote session id.");
    }
    await page.locator("#backToDashboardButton", { hasText: "Dashboard" }).click();
    await page.locator("#quoteDashboardPanel").waitFor({ state: "visible", timeout: 15000 });
    await expectTopbarPrimaryAction(page, "new-quote");
    await page.evaluate((storageKey) => {
      const saved = JSON.parse(window.localStorage.getItem(storageKey) || "{}");
      saved.quoteSessionId = "quote-unrelated-regression";
      saved.activeAppView = "dashboard";
      window.localStorage.setItem(storageKey, JSON.stringify(saved));
    }, "swooshz_quote_session_v1");
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.getByRole("heading", { name: "Dashboard" }).waitFor();
    const unrelatedLocalQuoteSessionId = await currentQuoteSessionId(page);
    if (unrelatedLocalQuoteSessionId === restoredQuoteSessionId) {
      throw new Error(`Expected Modify quote regression to use a non-current browser draft, found ${unrelatedLocalQuoteSessionId}.`);
    }
    await page.locator(`.dashboard-session-card[data-quote-session-id="${restoredQuoteSessionId}"]`).click();
    await page.locator('[data-dashboard-panel-action="modify-session"]', { hasText: "Modify quote" }).waitFor({ timeout: 15000 });
    const restoredDetailBeforeModify = await dashboardQuoteSessionDetail(page, restoredQuoteSessionId);
    const restoredUpdatedAtBeforeModify = restoredDetailBeforeModify.quote_session?.updated_at || "";
    if (!restoredUpdatedAtBeforeModify) {
      throw new Error("Expected dashboard session detail to expose updated_at before Modify quote.");
    }
    await page.locator('[data-dashboard-panel-action="modify-session"]', { hasText: "Modify quote" }).click();
    await page.locator("#panel-analysis.is-active").waitFor({ state: "visible", timeout: 15000 });
    await expectTopbarPrimaryAction(page, "dashboard");
    const modifiedQuoteSessionId = await currentQuoteSessionId(page);
    if (modifiedQuoteSessionId !== restoredQuoteSessionId) {
      throw new Error(`Expected Modify quote to restore saved dashboard session ${restoredQuoteSessionId}, found ${modifiedQuoteSessionId}.`);
    }
    const restoredActiveRailTexts = await page.locator(".rail-button.is-active").evaluateAll((buttons) => (
      buttons.map((button) => button.textContent?.trim() || "")
    ));
    if (restoredActiveRailTexts.length !== 1 || !["Upload", "Customer", "Quote Company"].includes(restoredActiveRailTexts[0])) {
      throw new Error(`Expected Modify quote to restore a usable quote panel, found ${JSON.stringify(restoredActiveRailTexts)}.`);
    }
    const restoredFiles = await page.locator("#fileList .file-item").evaluateAll((items) => (
      items.map((item) => item.textContent?.trim() || "")
    ));
    if (restoredFiles.length !== 1 || !restoredFiles[0].includes("kent-group.pdf")) {
      throw new Error(`Expected refresh to preserve the sample PDF reference, found ${JSON.stringify(restoredFiles)}.`);
    }
    await page.locator("#backToDashboardButton", { hasText: "Dashboard" }).click();
    await page.locator("#quoteDashboardPanel").waitFor({ state: "visible", timeout: 15000 });
    await expectTopbarPrimaryAction(page, "new-quote");
    const restoredCardAfterCleanReturn = page.locator(`.dashboard-session-card[data-quote-session-id="${restoredQuoteSessionId}"]`);
    await restoredCardAfterCleanReturn.waitFor({ state: "visible", timeout: 15000 });
    const restoredDetailAfterCleanReturn = await dashboardQuoteSessionDetail(page, restoredQuoteSessionId);
    const restoredUpdatedAtAfterCleanReturn = restoredDetailAfterCleanReturn.quote_session?.updated_at || "";
    if (restoredUpdatedAtAfterCleanReturn !== restoredUpdatedAtBeforeModify) {
      throw new Error(`Modify -> Dashboard without edits should not rewrite or reorder the session: ${restoredUpdatedAtBeforeModify} -> ${restoredUpdatedAtAfterCleanReturn}.`);
    }
    await restoredCardAfterCleanReturn.click();
    await page.locator('[data-dashboard-panel-action="modify-session"]', { hasText: "Modify quote" }).waitFor({ timeout: 15000 });
    await page.locator('[data-dashboard-panel-action="modify-session"]', { hasText: "Modify quote" }).click();
    await page.locator("#panel-analysis.is-active").waitFor({ state: "visible", timeout: 15000 });
    await expectTopbarPrimaryAction(page, "dashboard");
    await page.locator('.rail-button[data-side-panel="quote_company"]:not([disabled])').waitFor({ timeout: 15000 });
    await page.locator('.rail-button[data-side-panel="quote_company"]').click();
    await page.locator("#quoteCompanyPanel").waitFor({ state: "visible", timeout: 15000 });
    const restoredPresetValue = await page.locator("#presetSelect").inputValue();
    if (restoredPresetValue !== "profile:synthetic-fixture-default") {
      throw new Error(`Expected refresh to preserve company preset, found ${restoredPresetValue}.`);
    }
    const presetSelectBox = await page.locator("#presetSelect").boundingBox();
    if (!presetSelectBox || presetSelectBox.width < 200) {
      throw new Error("Company preset dropdown is unexpectedly narrow.");
    }
    await page.locator('.rail-button[data-side-panel="customer"]:not([disabled])').waitFor({ timeout: 15000 });
    await page.locator('.rail-button[data-side-panel="customer"]').click();
    await page.mouse.move(4, 4);
    await page.locator("#customerDetailsPanel").waitFor({ state: "visible" });
    await page.locator(".workspace-pane-scroll").evaluate((element) => {
      element.scrollTop = 0;
    });
    const pricingSummary = await page.locator("#selectedPricingReferenceSummary").innerText();
    if (!pricingSummary.includes("Managed in Settings")) {
      throw new Error(`Unexpected pricing reference summary: ${pricingSummary}`);
    }
    const pricingTaxText = await page.locator("#selectedPricingReferenceTax").innerText();
    if (!/GST|VAT/i.test(pricingTaxText)) {
      throw new Error(`Unexpected pricing reference tax badge: ${pricingTaxText}`);
    }
    const pricingCurrencyBox = await page.locator("#selectedPricingReferenceCurrency").boundingBox();
    const pricingTaxBox = await page.locator("#selectedPricingReferenceTax").boundingBox();
    if (!pricingCurrencyBox || !pricingTaxBox || pricingCurrencyBox.width < 70 || pricingTaxBox.width < 70) {
      throw new Error(`Pricing reference pills are unexpectedly narrow: ${JSON.stringify({ pricingCurrencyBox, pricingTaxBox })}`);
    }
    const selectedPricingValue = await page.locator("#profileSelect").inputValue();
    if (!selectedPricingValue) {
      throw new Error("Pricing reference select did not have a selected value.");
    }
    const profileSelectBox = await page.locator("#profileSelect").boundingBox();
    if (!profileSelectBox || profileSelectBox.width < 200) {
      throw new Error("Pricing reference dropdown is unexpectedly narrow.");
    }
    if (presetSelectBox.width > profileSelectBox.width) {
      throw new Error(`Company preset dropdown width ${presetSelectBox.width} should stay no wider than pricing reference dropdown width ${profileSelectBox.width}.`);
    }
    const customerPricingShot = await screenshot(page, "customer-pricing.png");
    await page.locator("#settingsButton").click();
    await page.locator("#pricingReferenceModal").waitFor({ state: "visible" });
    await page.getByRole("heading", { name: "Pricing Reference Settings" }).waitFor();
    await page.keyboard.press("Escape");
    await page.locator("#pricingReferenceModal").waitFor({ state: "hidden" });
    await page.locator("#quoteDate").waitFor({ state: "visible" });
    const dateBoldButton = page.locator('[data-date-format-command="bold"]');
    await page.locator("#quoteDate").focus();
    await dateBoldButton.waitFor({ state: "visible" });
    await dateBoldButton.click();
    if ((await dateBoldButton.getAttribute("aria-pressed")) !== "true") {
      throw new Error("Quote date bold formatting did not toggle on.");
    }
    await page.mouse.move(4, 4);
    const activeRailTexts = await page.locator(".rail-button.is-active").evaluateAll((buttons) => (
      buttons.map((button) => button.textContent?.trim() || "")
    ));
    if (activeRailTexts.length !== 1 || activeRailTexts[0] !== "Customer") {
      throw new Error(`Expected only Customer rail item to be active, found ${JSON.stringify(activeRailTexts)}.`);
    }
    const customerShot = await screenshot(page, "customer.png");

    const footerBox = await page.locator(".workspace-pane-footer").boundingBox();
    const viewport = page.viewportSize();
    if (!footerBox || !viewport || footerBox.y + footerBox.height > viewport.height + 1) {
      throw new Error("Workspace footer is not inside the current viewport.");
    }

    await page.setViewportSize({ width: 520, height: 720 });
    await page.evaluate(() => window.scrollTo(0, 0));
    const mobileScrollExtent = await page.evaluate(() => document.documentElement.scrollHeight - window.innerHeight);
    if (mobileScrollExtent < 80) {
      throw new Error(`Mobile layout did not create enough page scroll distance: ${mobileScrollExtent}.`);
    }
    const scrollPaneBox = await page.locator(".workspace-pane-scroll").boundingBox();
    if (!scrollPaneBox) {
      throw new Error("Workspace scroll pane was not visible in mobile layout.");
    }
    const mobileScrollBefore = await page.evaluate(() => window.scrollY);
    const paneScrollBefore = await page.locator(".workspace-pane-scroll").evaluate((element) => element.scrollTop);
    await page.mouse.move(scrollPaneBox.x + scrollPaneBox.width / 2, scrollPaneBox.y + Math.min(scrollPaneBox.height / 2, 260));
    await page.mouse.wheel(0, 520);
    await page.waitForTimeout(100);
    let mobileScrollAfter = await page.evaluate(() => window.scrollY);
    let paneScrollAfter = await page.locator(".workspace-pane-scroll").evaluate((element) => element.scrollTop);
    if (mobileScrollAfter <= mobileScrollBefore + 10 && paneScrollAfter <= paneScrollBefore + 10) {
      await page.keyboard.press("PageDown");
      await page.waitForTimeout(100);
      mobileScrollAfter = await page.evaluate(() => window.scrollY);
      paneScrollAfter = await page.locator(".workspace-pane-scroll").evaluate((element) => element.scrollTop);
    }
    if (mobileScrollAfter <= mobileScrollBefore + 10 && paneScrollAfter <= paneScrollBefore + 10) {
      throw new Error(`Mobile layout did not scroll from wheel or keyboard input: page ${mobileScrollBefore} -> ${mobileScrollAfter}, pane ${paneScrollBefore} -> ${paneScrollAfter}.`);
    }

    const bodyText = await page.locator("body").innerText();
    if (!bodyText.includes("Upload") || !bodyText.includes("Customer") || !bodyText.includes("Quote date")) {
      throw new Error("Rendered page did not include the expected workspace text.");
    }

    await page.setViewportSize({ width: 1365, height: 768 });
    await page.locator("#backToDashboardButton", { hasText: "Dashboard" }).waitFor({ state: "visible", timeout: 15000 });
    const currentDashboardSessionId = await currentQuoteSessionId(page);
    if (!currentDashboardSessionId) {
      throw new Error("Expected the current quote session id before returning to the dashboard.");
    }
    await page.locator("#backToDashboardButton").click();
    await page.locator("#quoteDashboardPanel").waitFor({ state: "visible", timeout: 15000 });
    await expectTopbarPrimaryAction(page, "new-quote");
    await page.locator("#quoteDashboardPanel").evaluate((element) => {
      element.scrollTop = 0;
    });
    const currentDashboardCard = page.locator(`.dashboard-session-card[data-quote-session-id="${currentDashboardSessionId}"]`);
    await currentDashboardCard.waitFor({ state: "visible", timeout: 15000 });
    const currentDashboardDetail = await dashboardQuoteSessionDetail(page, currentDashboardSessionId);
    const currentDraftState = currentDashboardDetail.quote_session?.draft_state || {};
    if (!Array.isArray(currentDraftState.images) || !currentDraftState.images.some((image) => /kent-group\.pdf/i.test(image.name || ""))) {
      throw new Error(`Expected dashboard session draft state to include the reference PDF metadata, found ${JSON.stringify(currentDraftState.images || [])}.`);
    }
    if (JSON.stringify(currentDraftState).includes("data:application/pdf")) {
      throw new Error("Dashboard session draft state should not store raw PDF data URLs.");
    }
    await currentDashboardCard.click();
    await page.locator("#dashboardSelectedSessionPanel").waitFor({ state: "visible", timeout: 15000 });
    const normalModeCheckboxVisible = await currentDashboardCard.locator(".dashboard-session-select-control").isVisible();
    if (normalModeCheckboxVisible) {
      throw new Error("Row checkbox should stay hidden until Select mode is enabled.");
    }
    const singlePanelActionMetrics = await dashboardPanelActionMetrics(page, "single selection");
    const dashboardSingleSelectedShot = await screenshot(page, "dashboard-single-selected.png");
    await createDashboardSmokeSession(page, "reference-c", {
      sessionIdPrefix: "quote-4f-search-row",
      customerName: "Kent Group Exhibition Booth",
      projectName: "Marina Bay Product Launch",
    });
    await page.getByRole("button", { name: "Clear selected session", exact: true }).click();
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.getByRole("heading", { name: "Dashboard" }).waitFor();
    await page.locator("#dashboardSearchInput").fill("4f");
    await page.locator(".dashboard-session-card").first().waitFor({ state: "visible", timeout: 15000 });
    const characterSearchRows = await page.locator(".dashboard-session-card").count();
    if (characterSearchRows !== 1) {
      throw new Error(`Expected search for 4f to match only the REF QUOTE-4F row, found ${characterSearchRows}.`);
    }
    const characterSearchText = await page.locator(".dashboard-session-card").first().innerText();
    if (!characterSearchText.includes("REF QUOTE-4F")) {
      throw new Error(`Search for 4f did not return the visible quote reference row: ${characterSearchText}`);
    }
    await page.locator("#dashboardSearchInput").fill("");
    await createDashboardSmokeSession(page, "alpha", { sessionIdPrefix: "quote-7a-playwright-alpha" });
    await createDashboardSmokeSession(page, "beta", { sessionIdPrefix: "quote-2c-hidden-3" });
    await createDashboardSmokeSession(page, "gamma", { sessionIdPrefix: "quote-bulk-extra-1" });
    await createDashboardSmokeSession(page, "delta", { sessionIdPrefix: "quote-bulk-extra-2" });
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.getByRole("heading", { name: "Dashboard" }).waitFor();
    await page.locator("#dashboardSearchInput").fill("7a");
    await page.locator(".dashboard-session-card").first().waitFor({ state: "visible", timeout: 15000 });
    const referenceSearchRows = await page.locator(".dashboard-session-card").count();
    if (referenceSearchRows !== 1) {
      throw new Error(`Expected search for 7a to match exactly the REF QUOTE-7A row, found ${referenceSearchRows}.`);
    }
    const referenceSearchText = await page.locator(".dashboard-session-card").first().innerText();
    if (!referenceSearchText.includes("REF QUOTE-7A")) {
      throw new Error(`Search for 7a did not return the visible quote reference row: ${referenceSearchText}`);
    }
    await page.locator("#dashboardSearchInput").fill("7a");
    await page.locator(".dashboard-session-card").first().waitFor({ state: "visible", timeout: 15000 });
    const digitSearchRows = await page.locator(".dashboard-session-card").count();
    if (digitSearchRows !== 1) {
      throw new Error(`Expected repeated search for 7a to match only the REF QUOTE-7A row, found ${digitSearchRows}.`);
    }
    const digitSearchText = await page.locator(".dashboard-session-card").first().innerText();
    if (!digitSearchText.includes("REF QUOTE-7A")) {
      throw new Error(`Repeated search for 7a did not return the visible quote reference row: ${digitSearchText}`);
    }
    await page.locator("#dashboardSearchInput").fill("");
    await createDashboardSmokeSession(page, "untitled-visible", {
      sessionIdPrefix: "quote-2c-untitled-visible",
      customerName: "",
      projectName: "",
    });
    await page.reload({ waitUntil: "domcontentloaded" });
    await page.getByRole("heading", { name: "Dashboard" }).waitFor();
    await page.locator("#dashboardSearchInput").fill("untitled customer");
    await page.locator(".dashboard-session-card").first().waitFor({ state: "visible", timeout: 15000 });
    const untitledCustomerTexts = await page.locator(".dashboard-session-card").evaluateAll((cards) => (
      cards.map((card) => card.innerText || "")
    ));
    if (!untitledCustomerTexts.some((text) => text.includes("Untitled customer") && text.includes("REF QUOTE-2C"))) {
      throw new Error(`Search for Untitled customer did not return the visible REF QUOTE-2C fallback row: ${JSON.stringify(untitledCustomerTexts)}`);
    }
    await page.locator("#dashboardSearchInput").fill("untitled quote");
    await page.locator(".dashboard-session-card").first().waitFor({ state: "visible", timeout: 15000 });
    const untitledProjectTexts = await page.locator(".dashboard-session-card").evaluateAll((cards) => (
      cards.map((card) => card.innerText || "")
    ));
    if (!untitledProjectTexts.some((text) => text.includes("Untitled quote") && text.includes("REF QUOTE-2C"))) {
      throw new Error(`Search for Untitled quote did not return the visible REF QUOTE-2C fallback row: ${JSON.stringify(untitledProjectTexts)}`);
    }
    await page.locator("#dashboardSearchInput").fill("Marina Bay Product Launch");
    await page.locator(".dashboard-session-card").first().waitFor({ state: "visible", timeout: 15000 });
    const visibleBulkRows = await page.locator(".dashboard-session-card").count();
    if (visibleBulkRows < 2) {
      throw new Error(`Expected at least two bulk smoke rows, found ${visibleBulkRows}.`);
    }
    const filteredFirstCard = page.locator(".dashboard-session-card").first();
    const selectModeButton = page.locator("#dashboardSelectModeButton", { hasText: "Select" });
    const selectModeBox = await selectModeButton.boundingBox();
    if (!selectModeBox || selectModeBox.height > 40 || selectModeBox.width > 100) {
      throw new Error(`Select mode control is too large: ${JSON.stringify(selectModeBox)}.`);
    }
    await selectModeButton.click();
    await page.locator("#dashboardSelectModeButton", { hasText: "Select all visible" }).waitFor({ state: "visible", timeout: 15000 });
    await filteredFirstCard.locator(".dashboard-session-select-control").waitFor({ state: "visible", timeout: 15000 });
    const firstCardTopBeforeBulk = await filteredFirstCard.evaluate((element) => element.getBoundingClientRect().top);
    await filteredFirstCard.click();
    const firstCardTopAfterBulk = await filteredFirstCard.evaluate((element) => element.getBoundingClientRect().top);
    if (Math.abs(firstCardTopAfterBulk - firstCardTopBeforeBulk) > 12) {
      throw new Error(`Bulk panel selection shifted the session list: ${firstCardTopBeforeBulk} -> ${firstCardTopAfterBulk}.`);
    }
    const rowCheckboxBox = await filteredFirstCard.locator("[data-dashboard-select]").boundingBox();
    if (!rowCheckboxBox || rowCheckboxBox.width > 18 || rowCheckboxBox.height > 18) {
      throw new Error(`Row checkbox is too large: ${JSON.stringify(rowCheckboxBox)}.`);
    }
    await page.locator(".dashboard-session-card.is-selected").first().waitFor({ state: "visible", timeout: 15000 });
    await page.locator("#dashboardSelectedSessionPanel", { hasText: "SELECTED SESSION" }).waitFor({ state: "visible", timeout: 15000 });
    const selectAllButton = page.locator("#dashboardSelectModeButton", { hasText: "Select all visible" });
    const selectAllBox = await selectAllButton.boundingBox();
    if (!selectAllBox || selectAllBox.height > 40 || selectAllBox.width > 180) {
      throw new Error(`Select all control is too large: ${JSON.stringify(selectAllBox)}.`);
    }
    await selectAllButton.click();
    await page.locator(".dashboard-bulk-selection-summary", { hasText: `${visibleBulkRows} quote sessions selected` }).waitFor({ state: "visible", timeout: 15000 });
    await page.locator("#dashboardSelectedSessionPanel", { hasText: `${visibleBulkRows} quote sessions selected` }).waitFor({ state: "visible", timeout: 15000 });
    await page.locator("#dashboardSelectedSessionPanel", { hasText: "Bulk selection" }).waitFor({ state: "visible", timeout: 15000 });
    const bulkPanelActionMetrics = await dashboardPanelActionMetrics(page, "bulk selection");
    if (Math.abs(bulkPanelActionMetrics.bottomGap - singlePanelActionMetrics.bottomGap) > 8) {
      throw new Error(`Dashboard action footer bottom moved between single and bulk states: ${singlePanelActionMetrics.bottomGap}px -> ${bulkPanelActionMetrics.bottomGap}px.`);
    }
    if (Math.abs(bulkPanelActionMetrics.actionX - singlePanelActionMetrics.actionX) > 4) {
      throw new Error(`Dashboard action footer x offset moved between single and bulk states: ${singlePanelActionMetrics.actionX}px -> ${bulkPanelActionMetrics.actionX}px.`);
    }
    if (Math.abs(bulkPanelActionMetrics.actionWidth - singlePanelActionMetrics.actionWidth) > 4) {
      throw new Error(`Dashboard action footer width changed between single and bulk states: ${singlePanelActionMetrics.actionWidth}px -> ${bulkPanelActionMetrics.actionWidth}px.`);
    }
    if (Math.abs(bulkPanelActionMetrics.deleteClearGap - singlePanelActionMetrics.deleteClearGap) > 4) {
      throw new Error(`Dashboard delete/clear spacing changed between single and bulk states: ${singlePanelActionMetrics.deleteClearGap}px -> ${bulkPanelActionMetrics.deleteClearGap}px.`);
    }
    await page.setViewportSize({ width: 520, height: 720 });
    const mobileBulkPanelActionMetrics = await dashboardPanelActionMetrics(page, "mobile bulk selection");
    if (mobileBulkPanelActionMetrics.bottomGap < 8 || mobileBulkPanelActionMetrics.bottomGap > 36) {
      throw new Error(`Mobile bulk action footer is not bottom anchored: ${mobileBulkPanelActionMetrics.bottomGap}px gap.`);
    }
    const dashboardSelectedMobileShot = await screenshot(page, "dashboard-selected-mobile.png");
    await page.setViewportSize({ width: 1365, height: 768 });
    await page.locator("#dashboardSelectedSessionPanel", { hasText: "Bulk selection" }).waitFor({ state: "visible", timeout: 15000 });
    const bulkExtraBlocks = await page.locator(".dashboard-bulk-breakdown, .dashboard-bulk-value-card").count();
    if (bulkExtraBlocks !== 0) {
      throw new Error("Bulk panel should not show status breakdown or combined value blocks.");
    }
    const bulkPanelText = await page.locator("#dashboardSelectedSessionPanel").innerText();
    if (bulkPanelText.includes("Status breakdown") || bulkPanelText.includes("Combined Value")) {
      throw new Error("Bulk panel copy still includes removed status or combined value sections.");
    }
    const selectAllChecked = await page.locator("#dashboardSelectModeButton").evaluate((button) => button.classList.contains("is-all-selected"));
    if (!selectAllChecked) {
      throw new Error("Select all control did not enter the checked visual state after selecting all visible rows.");
    }
    await selectAllButton.click();
    const selectedAfterDeselectAll = await page.locator(".dashboard-session-card.is-selected").count();
    if (selectedAfterDeselectAll !== 0) {
      throw new Error(`Select all toggle did not deselect visible rows: ${selectedAfterDeselectAll} remain selected.`);
    }
    await page.locator("#dashboardSelectedSessionPanel").waitFor({ state: "hidden", timeout: 15000 });
    await page.locator("#dashboardSelectModeButton", { hasText: "Select" }).waitFor({ state: "visible", timeout: 15000 });
    const returnedToNoneMode = await page.locator("#dashboardSelectModeButton").evaluate((button) => (
      !button.classList.contains("is-all-selected")
        && !button.classList.contains("is-selecting")
        && button.getAttribute("aria-pressed") === "false"
    ));
    if (!returnedToNoneMode) {
      throw new Error("Select all toggle did not return the control to the original Select mode.");
    }
    const rowCheckboxVisibleAfterDeselectAll = await filteredFirstCard.locator(".dashboard-session-select-control").isVisible();
    if (rowCheckboxVisibleAfterDeselectAll) {
      throw new Error("Row checkboxes stayed visible after Select all visible was toggled off.");
    }
    await page.locator("#dashboardSelectModeButton", { hasText: "Select" }).click();
    await page.locator("#dashboardSelectModeButton", { hasText: "Select all visible" }).waitFor({ state: "visible", timeout: 15000 });
    await page.locator("#dashboardSelectModeButton", { hasText: "Select all visible" }).click();
    await page.locator(".dashboard-bulk-selection-summary", { hasText: `${visibleBulkRows} quote sessions selected` }).waitFor({ state: "visible", timeout: 15000 });
    await page.locator("#quoteDashboardPanel").evaluate((panel) => {
      panel.scrollTop = 0;
    });
    const dashboardSelectedShot = await screenshot(page, "dashboard-selected.png");
    await page.locator('[data-dashboard-panel-action="delete-selected"]', { hasText: "Delete selected" }).click();
    await page.locator("#quoteSessionDeleteModal").waitFor({ state: "visible", timeout: 15000 });
    const dashboardDeleteModalShot = await screenshot(page, "dashboard-delete-modal.png");
    const bulkDeleteTitle = await page.locator("#quoteSessionDeleteTitle").innerText();
    if (bulkDeleteTitle !== "Delete selected quote sessions?") {
      throw new Error(`Unexpected bulk delete confirmation title: ${bulkDeleteTitle}`);
    }
    const bulkDeleteCopy = await page.locator("#quoteSessionDeleteText").innerText();
    if (bulkDeleteCopy !== "This removes the selected local dashboard records and any saved local exports for those quote sessions. This cannot be undone.") {
      throw new Error(`Unexpected bulk delete confirmation copy: ${bulkDeleteCopy}`);
    }
    await page.locator("#cancelQuoteSessionDeleteButton").click();
    await page.locator("#quoteSessionDeleteModal").waitFor({ state: "hidden", timeout: 15000 });
    await page.locator('[data-dashboard-panel-action="delete-selected"]', { hasText: "Delete selected" }).click();
    await page.locator("#quoteSessionDeleteModal").waitFor({ state: "visible", timeout: 15000 });
    await page.locator("#confirmQuoteSessionDeleteButton").click();
    await page.waitForFunction(() => document.querySelectorAll(".dashboard-session-card").length === 0, null, { timeout: 15000 });
    await page.locator("#dashboardSearchInput").fill("");
    await currentDashboardCard.waitFor({ state: "visible", timeout: 15000 });
    await currentDashboardCard.click();
    await page.locator('[data-dashboard-panel-action="delete-session"]', { hasText: "Delete session" }).click();
    await page.locator("#quoteSessionDeleteModal").waitFor({ state: "visible", timeout: 15000 });
    const singleDeleteCopy = await page.locator("#quoteSessionDeleteText").innerText();
    if (singleDeleteCopy !== "This removes the local dashboard record and any saved local exports for this quote session. This cannot be undone.") {
      throw new Error(`Unexpected single delete confirmation copy: ${singleDeleteCopy}`);
    }
    await page.locator("#cancelQuoteSessionDeleteButton").focus();
    await page.keyboard.press("Enter");
    await currentDashboardCard.waitFor({ state: "detached", timeout: 15000 });

    console.log(JSON.stringify({
      status: "ok",
      url: page.url(),
      screenshots: [
        dashboardShot,
        homeShot,
        customerPricingShot,
        customerShot,
        dashboardSingleSelectedShot,
        dashboardSelectedShot,
        dashboardSelectedMobileShot,
        dashboardDeleteModalShot,
      ].filter(Boolean),
      consoleProblems,
      networkProblems,
    }, null, 2));
  } finally {
    await browser.close();
    await stopServer(serverInfo);
  }
}

main().catch((error) => {
  console.error(error.message || error);
  process.exitCode = 1;
});
