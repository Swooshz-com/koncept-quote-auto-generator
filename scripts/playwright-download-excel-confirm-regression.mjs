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
  headed: args.includes("--headed"),
  screenshots: args.includes("--screenshots") || args.includes("--screenshot"),
  keepServer: args.includes("--keep-server"),
  host: readArg("--host", "127.0.0.1"),
  port: Number(readArg("--port", process.env.PLAYWRIGHT_PORT || "8767")),
};

const baseUrl = `http://${options.host}:${options.port}`;
const outputDir = path.join(root, "_output", "playwright");
const sessionStorageKey = "swooshz_quote_session_v1";
const tinyPngDataUrl =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

const partitionKeyword = "booth-structure-double-side-partition-wall-at-height-2-4m-wooden-construct-in-painted-finished-as-per-design-proposal";
const partitionDescription = "m length double side partition wall at height 2.4m; wooden construct in painted finished as per design proposal";

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
      env: { ...process.env, APP_MODE: "local", USER_TYPE: "operator" },
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

function quoteDetails() {
  return {
    quote_date: "2026-06-17",
    project_number: "KI-PLAYWRIGHT-001",
    client: {
      name: "Playwright Regression Client",
      attention: "Alex Tan",
      title: "Project Manager",
      address: "10 Sample Street\nSingapore 000010",
    },
    project: {
      title: "Accepted AI Confirm Export Regression",
      booth_width: "6",
      booth_depth: "6",
      booth_size: "6m x 6m",
      dimension_source: "analysis",
    },
    company: {
      name: "Koncept Image",
      header_details: "Koncept Image Pte Ltd\nSingapore",
      logo_data_url: tinyPngDataUrl,
      logo_name: "logo.png",
      logo_type: "image/png",
    },
    tax: { label: "GST", rate: 0.09 },
    quote_text: {
      terms_heading: "Terms & Conditions:",
      payment_terms: ["50% deposit upon confirmation", "Balance before handover"],
      notes_heading: "Note:",
      standard_notes: ["All prices are subject to final confirmation."],
      acceptance_text: "We accept the quotation amount and the terms",
      person_label: "Person in charge",
      stamp_label: "Company name & stamp",
      date_label: "Date:",
    },
    signature: {
      company_signatory: "Sales Team",
      company_title: "Project Sales",
      company_date_label: "Date:",
    },
    rich_text: {},
  };
}

function acceptedConfirmSessionSnapshot() {
  const basisLineId = "basis-double-side-partition";
  const rows = [
    {
      section: "Booth Structure",
      description: partitionDescription,
      quantity: 1,
      unit: "m length",
      price_mode: "Priced",
      unit_price_override: "",
      catalog_unit_price: 540,
      amount: 540,
      pricing_keyword: partitionKeyword,
      catalog_description: partitionDescription,
      pricing_reference_description: partitionDescription,
      source_basis_line_id: basisLineId,
    },
    {
      section: "AV Equipment Rental Items",
      description: 'nos. 55" LED TV Monitor (With Speaker - Full HD)',
      quantity: 1,
      unit: "nos",
      price_mode: "Priced",
      unit_price_override: "",
      catalog_unit_price: 675,
      amount: 675,
      pricing_keyword: "av-equipment-rental-items-nos-55-led-tv-monitor-with-speaker-full-hd",
      catalog_description: 'nos. 55" LED TV Monitor (With Speaker - Full HD)',
      pricing_reference_description: 'nos. 55" LED TV Monitor (With Speaker - Full HD)',
    },
    {
      section: "AV Equipment Rental Items",
      description: 'nos. 85" LED TV Monitor (With Speaker - Full HD)',
      quantity: 1,
      unit: "nos",
      price_mode: "Priced",
      unit_price_override: "",
      catalog_unit_price: 2250,
      amount: 2250,
      pricing_keyword: "av-equipment-rental-items-nos-85-led-tv-monitor-with-speaker-full-hd",
      catalog_description: 'nos. 85" LED TV Monitor (With Speaker - Full HD)',
      pricing_reference_description: 'nos. 85" LED TV Monitor (With Speaker - Full HD)',
    },
  ];
  return {
    version: 4,
    savedAt: new Date().toISOString(),
    profileId: "koncept",
    pricingReferenceId: "koncept-exhibition-quotation",
    pricingReferenceSource: "bundled",
    selectedPresetValue: "profile:koncept-image-default",
    images: [{
      name: "booth-render.png",
      type: "image/png",
      size: 68,
      data_url: tinyPngDataUrl,
    }],
    quoteDetails: quoteDetails(),
    workflowStage: "completed",
    quoteBasis: {},
    quoteBasisSections: [{
      id: "booth-structure",
      title: "Booth Structure",
      lines: [{
        id: basisLineId,
        tag: "Custom",
        custom_pricing: true,
        custom_confirmed: true,
        text: `[ ${partitionDescription} ]`,
        quantity: 1,
        unit: "m length",
      }],
    }, {
      id: "av-equipment-rental-items",
      title: "AV Equipment Rental Items",
      lines: [{
        id: "basis-55-tv",
        tag: "Include",
        text: '[ nos. 55" LED TV Monitor (With Speaker - Full HD) ]',
        quantity: 1,
        unit: "nos",
        pricing_keyword: "av-equipment-rental-items-nos-55-led-tv-monitor-with-speaker-full-hd",
        catalog_description: 'nos. 55" LED TV Monitor (With Speaker - Full HD)',
        pricing_reference_description: 'nos. 55" LED TV Monitor (With Speaker - Full HD)',
      }, {
        id: "basis-85-tv",
        tag: "Include",
        text: '[ nos. 85" LED TV Monitor (With Speaker - Full HD) ]',
        quantity: 1,
        unit: "nos",
        pricing_keyword: "av-equipment-rental-items-nos-85-led-tv-monitor-with-speaker-full-hd",
        catalog_description: 'nos. 85" LED TV Monitor (With Speaker - Full HD)',
        pricing_reference_description: 'nos. 85" LED TV Monitor (With Speaker - Full HD)',
      }],
    }],
    lineItems: rows.map((row) => ({
      section: row.section,
      description: row.description,
      quantity: row.quantity,
      unit: row.unit,
      pricing_keyword: row.pricing_keyword,
      price_mode: row.price_mode,
      catalog_unit_price: row.catalog_unit_price,
      catalog_description: row.catalog_description,
      pricing_reference_description: row.pricing_reference_description,
      source_basis_line_id: row.source_basis_line_id || "",
    })),
    outputRows: rows,
    originalOutputRows: rows,
    outputErrors: [],
    outputSortMode: "pricing_reference",
    analysisFindings: [],
    blockingClarificationQuestions: [],
    boothDimensions: {
      booth_width: "6",
      booth_depth: "6",
      booth_size: "6m x 6m",
      dimension_source: "analysis",
    },
    originalAnalysisSnapshot: null,
    basisConfirmed: true,
    aiFailed: false,
    draftSource: "playwright-regression",
    lastAnalysisMode: "standard",
    activeSidePanel: "output",
    downloadFile: null,
    pricingMatches: [],
    pricingIssues: [],
    activeJob: null,
  };
}

function basisReviewSessionSnapshot() {
  const snapshot = acceptedConfirmSessionSnapshot();
  return {
    ...snapshot,
    workflowStage: "basis_review",
    outputRows: [],
    originalOutputRows: [],
    outputErrors: [],
    basisConfirmed: false,
    activeSidePanel: "basis",
    downloadFile: null,
  };
}

async function outputTableText(page) {
  return (await page.locator("#pricingTableWrap").innerText()).replace(/\s+/g, " ").trim();
}

async function assertPartitionPriced(page, label) {
  const tableText = await outputTableText(page);
  if (!tableText.includes(partitionDescription)) {
    throw new Error(`${label}: partition row was not rendered. Table: ${tableText}`);
  }
  if (!tableText.includes("540.00")) {
    throw new Error(`${label}: partition price was not visible as 540.00. Table: ${tableText}`);
  }
  if (tableText.includes("???")) {
    throw new Error(`${label}: output table still contains ???. Table: ${tableText}`);
  }
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

  await fs.mkdir(outputDir, { recursive: true });
  const browser = await chromium.launch({ headless: !options.headed });
  const context = await browser.newContext({
    acceptDownloads: true,
    viewport: { width: 1365, height: 768 },
  });
  await context.addInitScript(({ key, snapshot }) => {
    window.localStorage.setItem(key, JSON.stringify(snapshot));
  }, { key: sessionStorageKey, snapshot: basisReviewSessionSnapshot() });

  const page = await context.newPage();
  await page.route("**/api/line-items/normalize", async (route) => {
    await new Promise((resolve) => setTimeout(resolve, 700));
    await route.continue();
  });
  const consoleProblems = [];
  page.on("console", (message) => {
    if (["error", "warning"].includes(message.type())) consoleProblems.push(`${message.type()}: ${message.text()}`);
  });
  page.on("pageerror", (error) => consoleProblems.push(`pageerror: ${error.message}`));

  try {
    await page.goto(baseUrl, { waitUntil: "networkidle" });
    await page.getByRole("heading", { name: "Swooshz Quote Generator" }).waitFor();
    await page.locator("#sideDrawerTitle", { hasText: "Quote Basis" }).waitFor();
    await page.locator("#sideNextButton", { hasText: "Confirm Quotation Basis" }).click();
    await page.locator("#excelGeneratingModal").waitFor({ state: "visible", timeout: 5000 });
    await page.locator("#excelGeneratingTitle", { hasText: "Preparing Output" }).waitFor({ timeout: 5000 });
    const preparingOutputShot = await screenshot(page, "preparing-output-modal.png");
    await page.locator("#sideDrawerTitle", { hasText: "Output" }).waitFor();
    await page.locator("#excelGeneratingModal").waitFor({ state: "hidden", timeout: 5000 });
    await page.locator("#pricingMatchesBody tr").first().waitFor({ timeout: 15000 });
    await assertPartitionPriced(page, "after confirm basis");

    const downloadPromise = page.waitForEvent("download", { timeout: 30000 });
    await page.getByRole("link", { name: "Download Excel" }).click();
    await page.locator("#excelGeneratingModal").waitFor({ state: "visible", timeout: 5000 });
    const generatingShot = await screenshot(page, "generating-excel-modal.png");
    await page.locator("#resultStatus", { hasText: "Completed" }).waitFor({ timeout: 30000 });
    const download = await downloadPromise;
    const downloadPath = path.join(outputDir, "accepted-ai-confirm-quotation.xlsx");
    await download.saveAs(downloadPath);
    await page.locator("#excelGeneratingModal").waitFor({ state: "hidden", timeout: 5000 });
    const completedShot = await screenshot(page, "download-complete-output.png");

    await assertPartitionPriced(page, "after download click");
    const downloaded = await fs.stat(downloadPath);
    if (!downloaded.size) {
      throw new Error(`Downloaded file is empty: ${downloadPath}`);
    }

    console.log(JSON.stringify({
      status: "ok",
      url: page.url(),
      downloadPath,
      downloadBytes: downloaded.size,
      screenshots: [preparingOutputShot, generatingShot, completedShot].filter(Boolean),
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
