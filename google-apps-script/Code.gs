const CONFIG = {
  RAW_SHEET: "Triple Whale Hourly",
  CLEAN_SHEET: "Clean_Data",
  ERROR_SHEET: "Error_Log",
  SUMMARY_SHEET: "Dashboard_Summary",
  EXPECTED_ROWS_PER_DAY: 24,
  ALERT_EMAIL: "",
  LOCAL_TZ: Session.getScriptTimeZone() || "Etc/UTC",
};

const CLEAN_HEADERS = [
  "row_key",
  "logged_at_utc",
  "logged_at_local",
  "sales_amount",
  "orders",
  "ad_spend",
  "roas",
  "sales_method",
  "source_sheet",
  "raw_logged_at",
  "ingested_at_utc",
];

const ERROR_HEADERS = [
  "ingested_at_utc",
  "reason",
  "row_key",
  "raw_logged_at",
  "raw_sales",
  "raw_ad_spend",
  "raw_orders",
  "source_sheet",
  "raw_row_json",
];

const SHOPIFY_DEFAULT_API_VERSION = "2024-10";
const SHOPIFY_ORDERS_QUERY = `
query Orders($cursor: String, $query: String!) {
  orders(first: 250, after: $cursor, query: $query, sortKey: CREATED_AT) {
    pageInfo {
      hasNextPage
      endCursor
    }
    edges {
      node {
        createdAt
        currentTotalPriceSet {
          shopMoney {
            amount
          }
        }
      }
    }
  }
}`;

function doGet(e) {
  ensureSheets_();
  const params = (e && e.parameter) || {};
  const mode = String(params.mode || "rows").toLowerCase();

  if (mode === "run") {
    return jsonOutput(runHourlyPipeline());
  }

  if (mode === "summary") {
    return jsonOutput(getSummaryPayload_());
  }

  if (mode === "clean") {
    return getSheetRowsJson_(CONFIG.CLEAN_SHEET);
  }

  const sheetName = params.sheet || CONFIG.RAW_SHEET;
  return getSheetRowsJson_(sheetName);
}

function doPost(e) {
  ensureSheets_();
  const payload = parsePostPayload_(e);
  const rows = Array.isArray(payload.rows) ? payload.rows : [payload];
  const appendStats = appendRawRows_(rows);
  const pipelineStats = runHourlyPipeline();

  return jsonOutput({
    ok: true,
    appendedToRaw: appendStats,
    pipeline: pipelineStats,
  });
}

function runHourlyPipeline() {
  ensureSheets_();
  const runAt = new Date();
  const ingestStats = ingestRawRows_(runAt);
  const shopifyStats = syncShopifySalesIntoClean_(runAt);
  const summary = updateSummary_(runAt, ingestStats, shopifyStats);
  maybeSendAlert_(summary.alerts);

  return {
    ok: true,
    ranAtUtc: runAt.toISOString(),
    ingest: ingestStats,
    shopify: shopifyStats,
    summary,
  };
}

function installHourlyTrigger() {
  const handler = "runHourlyPipeline";
  const existing = ScriptApp.getProjectTriggers().some((t) => t.getHandlerFunction() === handler);
  if (!existing) {
    ScriptApp.newTrigger(handler).timeBased().everyHours(1).create();
  }
}

function resetPipelineSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.CLEAN_SHEET, CONFIG.ERROR_SHEET, CONFIG.SUMMARY_SHEET].forEach((name) => {
    const s = ss.getSheetByName(name);
    if (s) {
      s.clear();
    }
  });
  ensureSheets_();
}

function configureShopifyConnection() {
  PropertiesService.getScriptProperties().setProperties(
    {
      SHOPIFY_STORE_DOMAIN: "your-store.myshopify.com",
      SHOPIFY_ACCESS_TOKEN: "replace_with_shpat_token",
      SHOPIFY_API_VERSION: SHOPIFY_DEFAULT_API_VERSION,
      SHOPIFY_API_KEY: "",
      SHOPIFY_API_SECRET: "",
    },
    false
  );
}

function testShopifyConnection() {
  const cfg = getShopifyConfig_();
  if (!cfg.isConfigured) {
    return { ok: false, error: "Missing SHOPIFY_STORE_DOMAIN or SHOPIFY_ACCESS_TOKEN in Script Properties." };
  }
  const runAt = new Date();
  const startUtc = new Date(Date.UTC(runAt.getUTCFullYear(), runAt.getUTCMonth(), 1, 0, 0, 0));
  const pull = pullShopifyHourlySales_(cfg, startUtc, runAt);
  return {
    ok: true,
    store: cfg.storeDomain,
    apiVersion: cfg.apiVersion,
    hoursReturned: Object.keys(pull.hourly).length,
    ordersFetched: pull.ordersFetched,
    pagesFetched: pull.pagesFetched,
  };
}

function ingestRawRows_(runAt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName(CONFIG.RAW_SHEET);
  const clean = ss.getSheetByName(CONFIG.CLEAN_SHEET);
  const err = ss.getSheetByName(CONFIG.ERROR_SHEET);
  const runAtIso = runAt.toISOString();

  if (!raw) {
    throw new Error(`Missing sheet: ${CONFIG.RAW_SHEET}`);
  }

  const values = raw.getDataRange().getValues();
  if (values.length <= 1) {
    return {
      scannedRows: 0,
      insertedRows: 0,
      invalidParseCount: 0,
      duplicateKeyCount: 0,
    };
  }

  const headers = values[0].map((h, i) => cleanText_(h || `Column ${i + 1}`));
  const body = values.slice(1).filter((row) => row.some((v) => cleanText_(v) !== ""));
  const existingKeys = getExistingKeys_(clean);

  const cleanRows = [];
  const errorRows = [];
  let invalidParseCount = 0;
  let duplicateKeyCount = 0;

  body.forEach((row) => {
    const obj = rowToObject_(headers, row);
    const normalized = normalizeRawRow_(obj, runAtIso);

    if (!normalized.valid) {
      invalidParseCount += 1;
      errorRows.push(toErrorRow_(normalized, obj, runAtIso, "INVALID_PARSE"));
      return;
    }

    if (existingKeys[normalized.rowKey]) {
      duplicateKeyCount += 1;
      errorRows.push(toErrorRow_(normalized, obj, runAtIso, "DUPLICATE_KEY"));
      return;
    }

    existingKeys[normalized.rowKey] = true;
    cleanRows.push([
      normalized.rowKey,
      normalized.loggedAtUtcIso,
      normalized.loggedAtLocal,
      normalized.salesAmount,
      normalized.orders,
      normalized.adSpend,
      normalized.roas,
      normalized.salesMethod,
      CONFIG.RAW_SHEET,
      normalized.rawLoggedAt,
      runAtIso,
    ]);
  });

  if (cleanRows.length) {
    clean.getRange(clean.getLastRow() + 1, 1, cleanRows.length, CLEAN_HEADERS.length).setValues(cleanRows);
  }

  if (errorRows.length) {
    err.getRange(err.getLastRow() + 1, 1, errorRows.length, ERROR_HEADERS.length).setValues(errorRows);
  }

  return {
    scannedRows: body.length,
    insertedRows: cleanRows.length,
    invalidParseCount,
    duplicateKeyCount,
  };
}

function updateSummary_(runAt, ingestStats, shopifyStats) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clean = ss.getSheetByName(CONFIG.CLEAN_SHEET);
  const summary = ss.getSheetByName(CONFIG.SUMMARY_SHEET);

  const values = clean.getDataRange().getValues();
  const headers = values[0] || [];
  const idx = indexMap_(headers);

  const now = runAt;
  const monthKeyNow = Utilities.formatDate(now, CONFIG.LOCAL_TZ, "yyyy-MM");
  const todayKey = Utilities.formatDate(now, CONFIG.LOCAL_TZ, "yyyy-MM-dd");
  const dayCounts = {};

  let mtdSales = 0;
  let mtdOrders = 0;
  let mtdAdSpend = 0;

  for (let i = 1; i < values.length; i += 1) {
    const row = values[i];
    if (!row || !row.length || cleanText_(row[idx.row_key]) === "") {
      continue;
    }
    const dt = new Date(row[idx.logged_at_utc]);
    if (Number.isNaN(dt.getTime())) {
      continue;
    }

    const monthKey = Utilities.formatDate(dt, CONFIG.LOCAL_TZ, "yyyy-MM");
    const dayKey = Utilities.formatDate(dt, CONFIG.LOCAL_TZ, "yyyy-MM-dd");
    dayCounts[dayKey] = (dayCounts[dayKey] || 0) + 1;

    if (monthKey !== monthKeyNow) {
      continue;
    }

    mtdSales += asNumber_(row[idx.sales_amount]) || 0;
    mtdOrders += asNumber_(row[idx.orders]) || 0;
    mtdAdSpend += asNumber_(row[idx.ad_spend]) || 0;
  }

  const rowsToday = dayCounts[todayKey] || 0;
  const rowsPerDayNear24 = Math.abs(rowsToday - CONFIG.EXPECTED_ROWS_PER_DAY) <= 2;
  const mtdRoas = mtdAdSpend > 0 ? mtdSales / mtdAdSpend : null;

  const alerts = [];
  if (ingestStats.invalidParseCount > 0) {
    alerts.push(`Invalid parse rows: ${ingestStats.invalidParseCount}`);
  }
  if (ingestStats.duplicateKeyCount > 0) {
    alerts.push(`Duplicate row keys: ${ingestStats.duplicateKeyCount}`);
  }
  if (!rowsPerDayNear24) {
    alerts.push(`Rows today ${rowsToday} (expected near ${CONFIG.EXPECTED_ROWS_PER_DAY})`);
  }
  if (shopifyStats && shopifyStats.status === "error") {
    alerts.push(`Shopify sync error: ${shopifyStats.error}`);
  }

  const rows = [
    ["metric", "value", "updated_at_utc"],
    ["last_run_utc", runAt.toISOString(), runAt.toISOString()],
    ["rows_scanned", ingestStats.scannedRows, runAt.toISOString()],
    ["rows_inserted", ingestStats.insertedRows, runAt.toISOString()],
    ["invalid_parse_count", ingestStats.invalidParseCount, runAt.toISOString()],
    ["duplicate_key_count", ingestStats.duplicateKeyCount, runAt.toISOString()],
    ["rows_today", rowsToday, runAt.toISOString()],
    ["rows_per_day_near_24", rowsPerDayNear24 ? "OK" : "ALERT", runAt.toISOString()],
    ["mtd_sales", round2_(mtdSales), runAt.toISOString()],
    ["mtd_orders", round2_(mtdOrders), runAt.toISOString()],
    ["mtd_ad_spend", round2_(mtdAdSpend), runAt.toISOString()],
    ["mtd_roas", mtdRoas === null ? "" : round4_(mtdRoas), runAt.toISOString()],
    ["shopify_sync_status", shopifyStats ? shopifyStats.status : "not_run", runAt.toISOString()],
    ["shopify_hours_upserted", shopifyStats && shopifyStats.hoursUpserted ? shopifyStats.hoursUpserted : 0, runAt.toISOString()],
    ["shopify_rows_inserted", shopifyStats && shopifyStats.insertedRows ? shopifyStats.insertedRows : 0, runAt.toISOString()],
    ["shopify_rows_updated", shopifyStats && shopifyStats.updatedRows ? shopifyStats.updatedRows : 0, runAt.toISOString()],
    ["shopify_orders_fetched", shopifyStats && shopifyStats.ordersFetched ? shopifyStats.ordersFetched : 0, runAt.toISOString()],
    ["alerts", alerts.join(" | "), runAt.toISOString()],
  ];

  summary.clear();
  summary.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  return {
    rowsToday,
    rowsPerDayNear24,
    mtdSales: round2_(mtdSales),
    mtdOrders: round2_(mtdOrders),
    mtdAdSpend: round2_(mtdAdSpend),
    mtdRoas: mtdRoas === null ? null : round4_(mtdRoas),
    shopify: shopifyStats || null,
    alerts,
  };
}

function syncShopifySalesIntoClean_(runAt) {
  const cfg = getShopifyConfig_();
  if (!cfg.isConfigured) {
    return {
      status: "skipped_missing_config",
      message: "Set SHOPIFY_STORE_DOMAIN and SHOPIFY_ACCESS_TOKEN in Script Properties.",
      insertedRows: 0,
      updatedRows: 0,
      hoursUpserted: 0,
      ordersFetched: 0,
    };
  }

  try {
    const startUtc = new Date(Date.UTC(runAt.getUTCFullYear(), runAt.getUTCMonth(), 1, 0, 0, 0));
    const pulled = pullShopifyHourlySales_(cfg, startUtc, runAt);
    const merged = mergeShopifyHourlyIntoClean_(pulled.hourly, runAt);

    return {
      status: "ok",
      store: cfg.storeDomain,
      apiVersion: cfg.apiVersion,
      startUtc: startUtc.toISOString(),
      endUtc: runAt.toISOString(),
      hoursUpserted: Object.keys(pulled.hourly).length,
      ordersFetched: pulled.ordersFetched,
      pagesFetched: pulled.pagesFetched,
      insertedRows: merged.insertedRows,
      updatedRows: merged.updatedRows,
    };
  } catch (error) {
    return {
      status: "error",
      error: String((error && error.message) || error || "Unknown Shopify sync error"),
      insertedRows: 0,
      updatedRows: 0,
      hoursUpserted: 0,
      ordersFetched: 0,
    };
  }
}

function pullShopifyHourlySales_(cfg, startUtc, endUtc) {
  const query = `created_at:>=${startUtc.toISOString()} created_at:<=${endUtc.toISOString()}`;
  const hourly = {};
  let cursor = null;
  let hasNextPage = true;
  let pagesFetched = 0;
  let ordersFetched = 0;
  const maxPages = 120;

  while (hasNextPage && pagesFetched < maxPages) {
    const page = fetchShopifyOrdersPage_(cfg, query, cursor);
    pagesFetched += 1;
    hasNextPage = page.hasNextPage;
    cursor = page.endCursor;

    page.orders.forEach((order) => {
      const dt = new Date(order.createdAt);
      if (Number.isNaN(dt.getTime())) {
        return;
      }
      const amount = asNumber_(order.amount) || 0;
      const hourKey = Utilities.formatDate(dt, "UTC", "yyyy-MM-dd-HH");
      if (!hourly[hourKey]) {
        hourly[hourKey] = { sales: 0, orders: 0 };
      }
      hourly[hourKey].sales += amount;
      hourly[hourKey].orders += 1;
      ordersFetched += 1;
    });
  }

  if (pagesFetched >= maxPages) {
    throw new Error(`Shopify pagination exceeded ${maxPages} pages; narrow the time window.`);
  }

  return { hourly, pagesFetched, ordersFetched };
}

function fetchShopifyOrdersPage_(cfg, query, cursor) {
  const url = `https://${cfg.storeDomain}/admin/api/${cfg.apiVersion}/graphql.json`;
  const payload = {
    query: SHOPIFY_ORDERS_QUERY,
    variables: {
      query,
      cursor: cursor || null,
    },
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      "X-Shopify-Access-Token": cfg.accessToken,
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const status = response.getResponseCode();
  const body = response.getContentText();
  if (status >= 400) {
    throw new Error(`Shopify API ${status}: ${body.slice(0, 300)}`);
  }

  const parsed = JSON.parse(body);
  if (parsed.errors && parsed.errors.length) {
    throw new Error(`Shopify GraphQL errors: ${JSON.stringify(parsed.errors)}`);
  }

  const ordersBlock = parsed && parsed.data && parsed.data.orders;
  if (!ordersBlock) {
    return { orders: [], hasNextPage: false, endCursor: null };
  }

  const orders = (ordersBlock.edges || []).map((edge) => {
    const node = edge && edge.node ? edge.node : {};
    const amount =
      node &&
      node.currentTotalPriceSet &&
      node.currentTotalPriceSet.shopMoney &&
      node.currentTotalPriceSet.shopMoney.amount;
    return {
      createdAt: node.createdAt,
      amount,
    };
  });

  return {
    orders,
    hasNextPage: Boolean(ordersBlock.pageInfo && ordersBlock.pageInfo.hasNextPage),
    endCursor: ordersBlock.pageInfo ? ordersBlock.pageInfo.endCursor : null,
  };
}

function mergeShopifyHourlyIntoClean_(hourlyMap, runAt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clean = ss.getSheetByName(CONFIG.CLEAN_SHEET);
  const values = clean.getDataRange().getValues();
  const headers = values[0] || CLEAN_HEADERS;
  const idx = indexMap_(headers);
  const rows = values.slice(1).filter((row) => row.some((v) => cleanText_(v) !== ""));

  const rowIndexByKey = {};
  rows.forEach((row, i) => {
    const key = cleanText_(row[idx.row_key]);
    if (key) {
      rowIndexByKey[key] = i;
    }
  });

  const runAtIso = runAt.toISOString();
  let insertedRows = 0;
  let updatedRows = 0;

  Object.keys(hourlyMap)
    .sort()
    .forEach((hourKey) => {
      const agg = hourlyMap[hourKey];
      const rowKey = `${hourKey}-tw`;
      const salesAmount = round2_(agg.sales || 0);
      const orders = round2_(agg.orders || 0);
      const utcDate = hourKeyToUtcDate_(hourKey);
      if (!utcDate) {
        return;
      }
      const loggedAtUtcIso = Utilities.formatDate(utcDate, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'");
      const loggedAtLocal = Utilities.formatDate(utcDate, CONFIG.LOCAL_TZ, "yyyy-MM-dd HH:mm:ss");

      if (Object.prototype.hasOwnProperty.call(rowIndexByKey, rowKey)) {
        const row = rows[rowIndexByKey[rowKey]];
        row[idx.sales_amount] = salesAmount;
        row[idx.orders] = orders;
        row[idx.sales_method] = "shopify_orders";
        row[idx.ingested_at_utc] = runAtIso;
        updatedRows += 1;
      } else {
        const newRow = [
          rowKey,
          loggedAtUtcIso,
          loggedAtLocal,
          salesAmount,
          orders,
          "",
          "",
          "shopify_orders",
          "shopify",
          "",
          runAtIso,
        ];
        rows.push(newRow);
        rowIndexByKey[rowKey] = rows.length - 1;
        insertedRows += 1;
      }
    });

  rows.sort((a, b) => {
    const ta = new Date(a[idx.logged_at_utc]).getTime();
    const tb = new Date(b[idx.logged_at_utc]).getTime();
    if (Number.isNaN(ta) || Number.isNaN(tb)) {
      return cleanText_(a[idx.row_key]).localeCompare(cleanText_(b[idx.row_key]));
    }
    return ta - tb;
  });

  clean.clearContents();
  clean.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    clean.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  return { insertedRows, updatedRows };
}

function getSummaryPayload_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(CONFIG.SUMMARY_SHEET);
  if (!s) {
    return { ok: false, error: "Summary sheet missing" };
  }
  const values = s.getDataRange().getValues();
  if (values.length <= 1) {
    return { ok: true, summary: {} };
  }
  const out = {};
  for (let i = 1; i < values.length; i += 1) {
    out[String(values[i][0])] = values[i][1];
  }
  return { ok: true, updatedAt: new Date().toISOString(), summary: out };
}

function appendRawRows_(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName(CONFIG.RAW_SHEET);
  if (!raw) {
    throw new Error(`Missing sheet: ${CONFIG.RAW_SHEET}`);
  }
  if (!rows.length) {
    return { rowsReceived: 0, rowsAppended: 0 };
  }

  const headerRange = raw.getRange(1, 1, 1, raw.getLastColumn() || 1);
  const headers = headerRange.getValues()[0].map((h, i) => cleanText_(h || `Column ${i + 1}`));
  const rowsToAppend = rows.map((obj) => headers.map((h) => (obj[h] == null ? "" : obj[h])));

  raw.getRange(raw.getLastRow() + 1, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
  return { rowsReceived: rows.length, rowsAppended: rowsToAppend.length };
}

function getSheetRowsJson_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return jsonOutput({ error: "Sheet not found", sheet: sheetName });
  }

  const values = sheet.getDataRange().getValues();
  if (!values.length) {
    return jsonOutput({
      updatedAt: new Date().toISOString(),
      sheet: sheetName,
      rowCount: 0,
      data: [],
    });
  }

  const headers = values[0].map((header, i) => cleanText_(header || `Column ${i + 1}`));
  const body = values.slice(1);
  const data = body
    .map((row) => rowToObject_(headers, row))
    .filter((row) => Object.keys(row).some((k) => row[k] !== null && row[k] !== ""));

  return jsonOutput({
    updatedAt: new Date().toISOString(),
    sheet: sheetName,
    rowCount: data.length,
    data,
  });
}

function normalizeRawRow_(raw, runAtIso) {
  const loggedAtRaw = cleanText_(raw["Logged At"]);
  const salesRaw = raw["Order Revenue (Current)"] != null ? raw["Order Revenue (Current)"] : raw["Blended Sales (Current)"];
  const adSpendRaw = raw["Blended Ad Spend (Current)"];
  const ordersRaw = raw["Orders (Current)"];

  const loggedAt = parseLoggedAt_(loggedAtRaw);
  if (!loggedAt) {
    return {
      valid: false,
      reason: "invalid_logged_at",
      rowKey: "",
      rawLoggedAt: loggedAtRaw,
      rawSales: cleanText_(salesRaw),
      rawAdSpend: cleanText_(adSpendRaw),
      rawOrders: cleanText_(ordersRaw),
      runAtIso,
    };
  }

  const roas = parseNumber_(raw["Blended ROAS (Current)"]);
  const adSpend = parseNumber_(adSpendRaw);
  const orders = parseNumber_(ordersRaw);
  const salesDirect = parseNumber_(salesRaw);

  let salesAmount = salesDirect;
  let salesMethod = "direct_sales_column";

  if (salesAmount == null && roas != null && adSpend != null) {
    salesAmount = roas * adSpend;
    salesMethod = "roas_x_ad_spend";
  }

  if (
    salesAmount != null &&
    roas != null &&
    adSpend != null &&
    raw["Blended Sales (Current)"] != null &&
    Math.abs(salesAmount - roas) < 0.000001
  ) {
    salesAmount = roas * adSpend;
    salesMethod = "roas_x_ad_spend_fallback";
  }

  if (salesAmount == null) {
    return {
      valid: false,
      reason: "invalid_sales",
      rowKey: "",
      rawLoggedAt: loggedAtRaw,
      rawSales: cleanText_(salesRaw),
      rawAdSpend: cleanText_(adSpendRaw),
      rawOrders: cleanText_(ordersRaw),
      runAtIso,
    };
  }

  const hourKey = Utilities.formatDate(loggedAt, "UTC", "yyyy-MM-dd-HH");
  const rowKey = `${hourKey}-tw`;

  return {
    valid: true,
    reason: "",
    rowKey,
    loggedAtUtcIso: Utilities.formatDate(loggedAt, "UTC", "yyyy-MM-dd'T'HH:mm:ss'Z'"),
    loggedAtLocal: Utilities.formatDate(loggedAt, CONFIG.LOCAL_TZ, "yyyy-MM-dd HH:mm:ss"),
    salesAmount: round2_(salesAmount),
    orders: orders == null ? "" : round2_(orders),
    adSpend: adSpend == null ? "" : round2_(adSpend),
    roas: roas == null ? "" : round4_(roas),
    salesMethod,
    rawLoggedAt: loggedAtRaw,
    rawSales: cleanText_(salesRaw),
    rawAdSpend: cleanText_(adSpendRaw),
    rawOrders: cleanText_(ordersRaw),
    runAtIso,
  };
}

function toErrorRow_(normalized, rawObj, runAtIso, reasonOverride) {
  return [
    runAtIso,
    reasonOverride || normalized.reason || "invalid_row",
    normalized.rowKey || "",
    normalized.rawLoggedAt || "",
    normalized.rawSales || "",
    normalized.rawAdSpend || "",
    normalized.rawOrders || "",
    CONFIG.RAW_SHEET,
    JSON.stringify(rawObj),
  ];
}

function ensureSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheetWithHeaders_(ss, CONFIG.CLEAN_SHEET, CLEAN_HEADERS);
  ensureSheetWithHeaders_(ss, CONFIG.ERROR_SHEET, ERROR_HEADERS);
  ensureSheetWithHeaders_(ss, CONFIG.SUMMARY_SHEET, ["metric", "value", "updated_at_utc"]);
}

function ensureSheetWithHeaders_(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
  }
  if (s.getLastRow() === 0) {
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const current = s.getRange(1, 1, 1, headers.length).getValues()[0];
  const mismatch = headers.some((h, i) => String(current[i] || "").trim() !== h);
  if (mismatch) {
    s.clear();
    s.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
}

function getExistingKeys_(cleanSheet) {
  const keys = {};
  const lastRow = cleanSheet.getLastRow();
  if (lastRow <= 1) {
    return keys;
  }
  const values = cleanSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  values.forEach((r) => {
    const key = cleanText_(r[0]);
    if (key) {
      keys[key] = true;
    }
  });
  return keys;
}

function parsePostPayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("Missing POST body");
  }
  const payload = JSON.parse(e.postData.contents);
  if (payload == null || typeof payload !== "object") {
    throw new Error("Invalid JSON payload");
  }
  return payload;
}

function getShopifyConfig_() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const storeDomain = cleanText_(props.SHOPIFY_STORE_DOMAIN);
  const accessToken = cleanText_(props.SHOPIFY_ACCESS_TOKEN);
  const apiVersion = cleanText_(props.SHOPIFY_API_VERSION) || SHOPIFY_DEFAULT_API_VERSION;

  return {
    storeDomain,
    accessToken,
    apiVersion,
    apiKey: cleanText_(props.SHOPIFY_API_KEY),
    apiSecret: cleanText_(props.SHOPIFY_API_SECRET),
    isConfigured: Boolean(storeDomain && accessToken),
  };
}

function parseLoggedAt_(raw) {
  const text = cleanText_(raw);
  if (!text) {
    return null;
  }
  const normalized = text.replace(" GMT", "").replace(" at ", " ");
  const parsed = new Date(`${normalized} UTC`);
  if (Number.isNaN(parsed.getTime())) {
    return null;
  }
  return parsed;
}

function parseNumber_(raw) {
  if (raw === null || typeof raw === "undefined" || raw === "") {
    return null;
  }
  if (typeof raw === "number") {
    return Number.isFinite(raw) ? raw : null;
  }

  const cleaned = cleanText_(raw)
    .replace(/[$,%]/g, "")
    .replace(/,/g, "")
    .replace(/\s+/g, "");

  if (!cleaned) {
    return null;
  }
  const num = Number(cleaned);
  return Number.isFinite(num) ? num : null;
}

function rowToObject_(headers, row) {
  const out = {};
  headers.forEach((h, i) => {
    out[h] = normalizeValue_(row[i]);
  });
  return out;
}

function hourKeyToUtcDate_(hourKey) {
  const match = String(hourKey || "").match(/^(\d{4})-(\d{2})-(\d{2})-(\d{2})$/);
  if (!match) {
    return null;
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  return new Date(Date.UTC(year, month - 1, day, hour, 0, 0));
}

function indexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[String(h)] = i;
  });
  return map;
}

function maybeSendAlert_(alerts) {
  if (!alerts || !alerts.length) {
    return;
  }
  Logger.log(`Pipeline alerts: ${alerts.join(" | ")}`);
  if (!CONFIG.ALERT_EMAIL) {
    return;
  }
  MailApp.sendEmail({
    to: CONFIG.ALERT_EMAIL,
    subject: "Elave Dashboard data quality alert",
    body: alerts.join("\n"),
  });
}

function normalizeValue_(value) {
  if (value === "" || typeof value === "undefined") {
    return null;
  }
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return value.toISOString();
  }
  return value;
}

function cleanText_(value) {
  return String(value == null ? "" : value)
    .replace(/\u200B/g, "")
    .replace(/\u00A0/g, " ")
    .trim();
}

function asNumber_(value) {
  if (value == null || value === "") {
    return null;
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}

function round4_(n) {
  return Math.round(n * 10000) / 10000;
}

function jsonOutput(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}
