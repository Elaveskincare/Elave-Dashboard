window.ELAVE_DASH_CONFIG = {
  // Apps Script web app URL (from Deploy > Manage deployments).
  appsScriptUrl: "https://script.google.com/macros/s/REPLACE_ME/exec",

  // Optional backend API URL (Supabase/Shopify merged data).
  backendApiUrl: "",

  // Sheet tab name in your Google Sheet.
  sheetName: "Triple Whale Hourly",

  // Refresh every 15 minutes.
  refreshIntervalMs: 15 * 60 * 1000,

  // Rotate selected metric in presentation mode.
  cycleIntervalMs: 15 * 1000,

  // Simple browser lock.
  passcode: "4321",

  // Sales target logic:
  // 1 = match/beat previous month, 1.1 = 10% above previous month.
  salesTargetMultiplier: 1,

  // Optional hard target override. Set null to use previous month logic.
  salesTargetValue: null,

  // Optional month-specific targets that override multiplier/global value.
  // Format: "YYYY-MM": number
  salesTargetsByMonth: {
    // "2026-02": 120000,
  },

  // Dashboard is fixed to EUR (config value kept for compatibility).
  currencySymbol: "EUR",
};
