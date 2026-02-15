import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, "..");

const sourceConfigPath = path.join(rootDir, "config.js");
const exampleConfigPath = path.join(rootDir, "config.example.js");
const publicDir = path.join(rootDir, "public");
const targetConfigPath = path.join(publicDir, "config.js");

async function run() {
  await fs.mkdir(publicDir, { recursive: true });

  if (await fileExists(sourceConfigPath)) {
    await fs.copyFile(sourceConfigPath, targetConfigPath);
    return;
  }

  if (hasDashEnv()) {
    await fs.writeFile(targetConfigPath, buildConfigFromEnv(), "utf8");
    return;
  }

  if (await fileExists(exampleConfigPath)) {
    await fs.copyFile(exampleConfigPath, targetConfigPath);
  }
}

async function fileExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch (_error) {
    return false;
  }
}

function hasDashEnv() {
  return Object.keys(process.env).some((key) => key.startsWith("DASH_"));
}

function buildConfigFromEnv() {
  const appsScriptUrl = process.env.DASH_APPS_SCRIPT_URL || "";
  const backendApiUrl = process.env.DASH_BACKEND_API_URL || "";
  const sheetName = process.env.DASH_SHEET_NAME || "Triple Whale Hourly";
  const refreshIntervalMs = toInt(process.env.DASH_REFRESH_INTERVAL_MS, 15 * 60 * 1000);
  const cycleIntervalMs = toInt(process.env.DASH_CYCLE_INTERVAL_MS, 15 * 1000);
  const passcode = process.env.DASH_PASSCODE || "";
  const salesTargetMultiplier = toNumber(process.env.DASH_SALES_TARGET_MULTIPLIER, 1.15);
  const salesTargetValue = process.env.DASH_SALES_TARGET_VALUE ? toNumber(process.env.DASH_SALES_TARGET_VALUE, null) : null;
  const currencySymbol = process.env.DASH_CURRENCY_SYMBOL || "EUR";

  return `window.ELAVE_DASH_CONFIG = {
  appsScriptUrl: ${JSON.stringify(appsScriptUrl)},
  backendApiUrl: ${JSON.stringify(backendApiUrl)},
  sheetName: ${JSON.stringify(sheetName)},
  refreshIntervalMs: ${refreshIntervalMs},
  cycleIntervalMs: ${cycleIntervalMs},
  passcode: ${JSON.stringify(passcode)},
  salesTargetMultiplier: ${salesTargetMultiplier},
  salesTargetValue: ${salesTargetValue === null ? "null" : salesTargetValue},
  salesTargetsByMonth: {},
  currencySymbol: ${JSON.stringify(currencySymbol)},
};
`;
}

function toInt(value, fallback) {
  const n = Number.parseInt(String(value || ""), 10);
  return Number.isFinite(n) ? n : fallback;
}

function toNumber(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

run().catch((error) => {
  console.error("Failed to sync config.js into public/", error);
  process.exitCode = 1;
});
