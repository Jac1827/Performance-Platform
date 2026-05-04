#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const browserClientPath =
  "/Users/jacheflin/.codex/plugins/cache/openai-bundled/browser-use/0.1.0-alpha1/scripts/browser-client.mjs";
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..");

const publishedPages = [
  {
    label: "performance platform",
    relativePath: "docs/index.html",
    markers: [
      "<title>RISE Performance Platform</title>",
      "Performance Platform",
      "Save To Cloud",
      "Pull From Cloud",
      "Copy Bundle",
    ],
  },
  {
    label: "operations dashboard",
    relativePath: "docs/portfolio-operations-dashboard/index.html",
    markers: [
      "ATLAS Portfolio Tracker Reports for RISE logo",
      "Portfolio Home",
      "Community Setup",
      "Data Import",
    ],
  },
  {
    label: "financial accountability",
    relativePath: "docs/portfolio-operations-dashboard/financial-accountability.html",
    markers: [
      "<title>RISE Residential Financial Platform</title>",
      '<div id="app"></div>',
      "RISE Residential Financial Platform",
      "Investor Report",
      "Upload",
    ],
  },
];

const sourceParityPairs = [
  ["index.html", "docs/index.html"],
  ["rise_leasing_dashboard-5-12.html", "docs/portfolio-operations-dashboard/index.html"],
  ["financial-accountability.html", "docs/portfolio-operations-dashboard/financial-accountability.html"],
];

function fail(message) {
  throw new Error(message);
}

function readUtf8(relativePath) {
  return fs.readFileSync(path.join(repoRoot, relativePath), "utf8");
}

function requireIncludes(haystack, needle, label) {
  if (!haystack.includes(needle)) {
    fail(`${label} is missing expected marker: ${needle}`);
  }
}

function validateRelativeRefs(relativePath, html) {
  const attrRe = /(?:href|src)=(["'])(.*?)\1/g;
  const pageDir = path.dirname(path.join(repoRoot, relativePath));
  for (const match of html.matchAll(attrRe)) {
    const ref = match[2];
    if (
      !ref ||
      ref.includes("${") ||
      ref.startsWith("http:") ||
      ref.startsWith("https:") ||
      ref.startsWith("data:") ||
      ref.startsWith("mailto:") ||
      ref.startsWith("#") ||
      ref.startsWith("javascript:")
    ) {
      continue;
    }
    const target = path.resolve(pageDir, ref.split("?")[0].split("#")[0]);
    if (!fs.existsSync(target)) {
      fail(`${relativePath} points to a missing asset: ${ref}`);
    }
  }
}

function runStaticChecks() {
  for (const page of publishedPages) {
    const fullPath = path.join(repoRoot, page.relativePath);
    if (!fs.existsSync(fullPath)) {
      fail(`Missing published page: ${page.relativePath}`);
    }
    const html = readUtf8(page.relativePath);
    page.markers.forEach((marker) => requireIncludes(html, marker, page.relativePath));
    validateRelativeRefs(page.relativePath, html);
  }

  for (const [sourcePath, publishedPath] of sourceParityPairs) {
    const source = fs.readFileSync(path.join(repoRoot, sourcePath));
    const published = fs.readFileSync(path.join(repoRoot, publishedPath));
    if (!source.equals(published)) {
      fail(`Source/published drift detected between ${sourcePath} and ${publishedPath}`);
    }
  }
}

async function tryRunBrowserChecks() {
  try {
    const { setupAtlasRuntime } = await import(browserClientPath);
    await setupAtlasRuntime({ globals: globalThis, backend: "iab" });
    await agent.browser.nameSession("🔎 daily glitch review smoke");
  } catch (error) {
    return {
      mode: "static-only",
      reason: error instanceof Error ? error.message : String(error),
    };
  }

  const pageUrls = {
    performance: pathToFileURL(path.join(repoRoot, "docs/index.html")).href,
    operations: pathToFileURL(path.join(repoRoot, "docs/portfolio-operations-dashboard/index.html")).href,
    financial: pathToFileURL(path.join(repoRoot, "docs/portfolio-operations-dashboard/financial-accountability.html")).href,
  };

  async function requireVisible(locator, label) {
    if (!(await locator.isVisible())) {
      fail(`${label} was not visible`);
    }
  }

  async function requireCountAtLeast(locator, label, minimum) {
    const count = await locator.count();
    if (count < minimum) {
      fail(`${label} count was ${count}, expected at least ${minimum}`);
    }
  }

  async function requireNoConsoleErrors(tab, label) {
    const errors = await tab.dev.logs({ levels: ["error"], limit: 50 });
    if (errors.length > 0) {
      const lines = errors.map((entry) => `[${entry.level}] ${entry.message}`).join("\n");
      fail(`${label} emitted console errors:\n${lines}`);
    }
  }

  async function verifyPerformancePage(tab) {
    await tab.goto(pageUrls.performance);
    await tab.playwright.waitForLoadState({ state: "load", timeoutMs: 15000 });
    await requireNoConsoleErrors(tab, "docs/index.html");
    await requireVisible(tab.playwright.getByText("Performance Platform", { exact: false }), "Performance Platform title");
    await requireVisible(tab.playwright.getByText("Secure coaching, PIP, and roster management", { exact: false }), "performance subtitle");
  }

  async function verifyOperationsPage(tab) {
    await tab.goto(pageUrls.operations);
    await tab.playwright.waitForLoadState({ state: "load", timeoutMs: 15000 });
    await requireNoConsoleErrors(tab, "docs/portfolio-operations-dashboard/index.html");
    await requireVisible(tab.playwright.locator('img[src*="atlas-rise-logo"]'), "ATLAS logo");
    await requireVisible(tab.playwright.locator('img[src*="atlas-rise-tagline"]'), "tagline image");
    await requireVisible(tab.playwright.getByRole("button", { name: "Save Data" }), "Save Data button");
    await requireVisible(tab.playwright.getByRole("button", { name: "Portfolio Home" }), "Portfolio Home tab");
    await requireVisible(tab.playwright.getByRole("button", { name: "Community Setup" }), "Community Setup tab");
    await requireVisible(tab.playwright.getByRole("button", { name: "Data Import" }), "Data Import tab");
    await requireCountAtLeast(tab.playwright.locator("button.tab"), "operations workspace tabs", 8);
  }

  async function verifyFinancialPage(tab) {
    await tab.goto(pageUrls.financial);
    await tab.playwright.waitForLoadState({ state: "load", timeoutMs: 15000 });
    await requireNoConsoleErrors(tab, "docs/portfolio-operations-dashboard/financial-accountability.html");
    await requireVisible(tab.playwright.locator('img[src*="rise-wordmark-blue"]'), "financial page logo");
    await requireVisible(tab.playwright.getByRole("button", { name: "Import Financial Files" }), "Import Financial Files button");
    await requireVisible(tab.playwright.getByRole("button", { name: "Export Scorecards CSV" }), "Export Scorecards CSV button");
    await requireVisible(tab.playwright.getByRole("button", { name: "Export Board Summary" }), "Export Board Summary button");
    await requireVisible(tab.playwright.getByRole("button", { name: "Export Ownership Report" }), "Export Ownership Report button");
    await requireVisible(tab.playwright.getByRole("button", { name: "Export Board Deck" }), "Export Board Deck button");
    await requireCountAtLeast(tab.playwright.locator("button.tab-btn"), "financial layer tabs", 3);
  }

  const performanceTab = await agent.browser.tabs.new();
  await verifyPerformancePage(performanceTab);

  const operationsTab = await agent.browser.tabs.new();
  await verifyOperationsPage(operationsTab);

  const financialTab = await agent.browser.tabs.new();
  await verifyFinancialPage(financialTab);

  return { mode: "browser+static" };
}

async function main() {
  runStaticChecks();
  const browserResult = await tryRunBrowserChecks();
  if (browserResult.mode === "static-only") {
    console.log(
      `Daily Glitch Review smoke passed static coverage for docs/index.html, docs/portfolio-operations-dashboard/index.html, and docs/portfolio-operations-dashboard/financial-accountability.html. Browser checks were skipped: ${browserResult.reason}`,
    );
    return;
  }
  console.log(
    "Daily Glitch Review smoke passed static and browser coverage for docs/index.html, docs/portfolio-operations-dashboard/index.html, and docs/portfolio-operations-dashboard/financial-accountability.html.",
  );
}

main().catch((error) => {
  console.error(error instanceof Error ? error.stack || error.message : String(error));
  process.exitCode = 1;
});
