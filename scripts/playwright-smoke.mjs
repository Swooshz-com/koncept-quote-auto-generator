import { spawn } from "node:child_process";
import fs from "node:fs/promises";
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

function pythonCommand() {
  return process.env.PYTHON || (process.platform === "win32" ? "python" : "python3");
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
      env: { ...process.env, APP_MODE: "local" },
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

async function main() {
  let serverInfo = null;
  if (!(await healthOk())) {
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
  page.on("console", (message) => {
    if (["error", "warning"].includes(message.type())) consoleProblems.push(`${message.type()}: ${message.text()}`);
  });
  page.on("pageerror", (error) => consoleProblems.push(`pageerror: ${error.message}`));

  try {
    await page.goto(baseUrl, { waitUntil: "networkidle" });
    await page.getByRole("heading", { name: "Swooshz Quote Generator" }).waitFor();
    await page.locator("#imageIntake").waitFor({ state: "visible" });
    const homeShot = await screenshot(page, "home.png");

    await page.locator("#sampleDetailsButton:not([disabled])").waitFor({ timeout: 15000 });
    await page.locator("#sampleDetailsButton").click();
    await page.locator("#fileList .file-item").first().waitFor({ timeout: 15000 });
    await page.locator('[data-side-panel="quote_company"]:not([disabled])').waitFor({ timeout: 15000 });
    await page.locator('[data-side-panel="quote_company"]').click();
    await page.locator("#quoteCompanyPanel").waitFor({ state: "visible" });
    const samplePresetValue = await page.locator("#presetSelect").inputValue();
    if (samplePresetValue !== "profile:koncept-image-default") {
      throw new Error(`Expected sample to select Koncept Images preset, found ${samplePresetValue}.`);
    }
    await page.reload({ waitUntil: "networkidle" });
    await page.getByRole("heading", { name: "Swooshz Quote Generator" }).waitFor();
    await page.locator('[data-side-panel="quote_company"]:not([disabled])').waitFor({ timeout: 15000 });
    await page.locator('[data-side-panel="quote_company"]').click();
    await page.locator("#quoteCompanyPanel").waitFor({ state: "visible" });
    const restoredPresetValue = await page.locator("#presetSelect").inputValue();
    if (restoredPresetValue !== "profile:koncept-image-default") {
      throw new Error(`Expected refresh to preserve company preset, found ${restoredPresetValue}.`);
    }
    await page.locator('[data-side-panel="customer"]:not([disabled])').waitFor({ timeout: 15000 });
    await page.locator('[data-side-panel="customer"]').click();
    await page.mouse.move(4, 4);
    await page.locator("#customerDetailsPanel").waitFor({ state: "visible" });
    await page.locator(".workspace-pane-scroll").evaluate((element) => {
      element.scrollTop = 0;
    });
    const pricingButtonText = await page.locator(".pricing-reference-panel .settings-button-row").innerText();
    if (!pricingButtonText.includes("New") || !pricingButtonText.includes("Delete") || pricingButtonText.includes("New Pricing Reference")) {
      throw new Error(`Unexpected pricing reference button labels: ${pricingButtonText}`);
    }
    const profileSelectBox = await page.locator("#profileSelect").boundingBox();
    const pricingButtonsBox = await page.locator(".pricing-reference-panel .settings-button-row").boundingBox();
    if (!profileSelectBox || !pricingButtonsBox || Math.abs(profileSelectBox.width - pricingButtonsBox.width) > 2) {
      throw new Error("Pricing reference buttons are not aligned to the dropdown width.");
    }
    const customerPricingShot = await screenshot(page, "customer-pricing.png");
    await page.locator("#quoteDate").waitFor({ state: "visible" });
    const dateBoldButton = page.locator('[data-date-format-command="bold"]');
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

    const bodyText = await page.locator("body").innerText();
    if (!bodyText.includes("Images") || !bodyText.includes("Customer") || !bodyText.includes("Quote date")) {
      throw new Error("Rendered page did not include the expected workspace text.");
    }

    console.log(JSON.stringify({
      status: "ok",
      url: page.url(),
      screenshots: [homeShot, customerPricingShot, customerShot].filter(Boolean),
      consoleProblems,
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
