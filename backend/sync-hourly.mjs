import { loadEnvFile, requireEnv } from "./lib/env.mjs";
import { cleanNumber, hourKeyFromDate, round, utcFromHourKey } from "./lib/metrics-utils.mjs";
import { getSupabaseConfig, supabaseSelectAll, supabaseUpsert } from "./lib/supabase-rest.mjs";

loadEnvFile();

const SHOPIFY_API_VERSION = process.env.SHOPIFY_API_VERSION || "2024-10";
const SYNC_DAYS = Math.max(7, Math.min(365, Number(process.env.SYNC_DAYS || "90")));

async function main() {
  const appsScriptUrl = requireEnv("APPS_SCRIPT_URL");
  const shopifyDomain = requireEnv("SHOPIFY_STORE_DOMAIN");
  const shopifyToken = requireEnv("SHOPIFY_ACCESS_TOKEN");
  const supabase = getSupabaseConfig();

  const now = new Date();
  const start = new Date(now.getTime() - SYNC_DAYS * 24 * 60 * 60 * 1000);
  const runAtIso = now.toISOString();

  const [marketingByHour, shopifyDetail, existingRows] = await Promise.all([
    fetchMarketingFromAppsScript(appsScriptUrl),
    fetchShopifyOrdersDetailed({
      domain: shopifyDomain,
      token: shopifyToken,
      apiVersion: SHOPIFY_API_VERSION,
      start,
      end: now,
    }),
    fetchExistingRows({ supabase, start }),
  ]);

  const shopifyStructured = buildShopifyStructures(shopifyDetail.orders, runAtIso);
  const mergedHourlyRows = buildMergedHourlyRows({
    marketingByHour,
    salesByHour: shopifyStructured.hourly,
    existingRows,
    runAtIso,
  });

  const hourlyAffected = await upsertChunks({
    config: supabase,
    table: supabase.table,
    onConflict: "row_key",
    rows: mergedHourlyRows,
  });

  const ordersAffected = await upsertChunks({
    config: supabase,
    table: supabase.ordersTable,
    onConflict: "order_id",
    rows: shopifyStructured.orderRows,
  });

  const linesAffected = await upsertChunks({
    config: supabase,
    table: supabase.orderLinesTable,
    onConflict: "order_line_key",
    rows: shopifyStructured.lineRows,
  });

  const summary = {
    ok: true,
    sync_window: {
      days: SYNC_DAYS,
      start_utc: start.toISOString(),
      end_utc: now.toISOString(),
    },
    apps_script: {
      marketing_hours: Object.keys(marketingByHour).length,
    },
    shopify: {
      orders_fetched: shopifyDetail.orders.length,
      pages_fetched: shopifyDetail.pagesFetched,
      hourly_points: Object.keys(shopifyStructured.hourly).length,
      order_rows_upserted: shopifyStructured.orderRows.length,
      line_rows_upserted: shopifyStructured.lineRows.length,
    },
    supabase: {
      hourly_rows_input: mergedHourlyRows.length,
      hourly_rows_returned: hourlyAffected,
      order_rows_returned: ordersAffected,
      line_rows_returned: linesAffected,
    },
  };

  console.log(JSON.stringify(summary, null, 2));
}

async function upsertChunks({ config, table, onConflict, rows, chunkSize = 500 }) {
  if (!rows.length) {
    return 0;
  }
  let affected = 0;
  for (let i = 0; i < rows.length; i += chunkSize) {
    const chunk = rows.slice(i, i + chunkSize);
    const result = await supabaseUpsert({ config, table, onConflict, rows: chunk });
    affected += result.length;
  }
  return affected;
}

function extractRows(payload) {
  if (Array.isArray(payload)) {
    return payload;
  }
  if (payload && Array.isArray(payload.data)) {
    return payload.data;
  }
  if (payload && Array.isArray(payload.rows)) {
    return payload.rows;
  }
  return [];
}

function normalizeHourKey(row) {
  const rowKey = String(row.row_key || "").trim();
  const rowMatch = rowKey.match(/^(\d{4}-\d{2}-\d{2}-\d{2})/);
  if (rowMatch) {
    return rowMatch[1];
  }

  const logged = row.logged_at_utc || row.logged_at_local || row["Logged At"];
  const dt = new Date(logged);
  if (!Number.isNaN(dt.getTime())) {
    return hourKeyFromDate(dt);
  }
  return "";
}

async function fetchMarketingFromAppsScript(appsScriptUrl) {
  const separator = appsScriptUrl.includes("?") ? "&" : "?";
  const url = `${appsScriptUrl}${separator}mode=clean&_=${Date.now()}`;
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Apps Script mode=clean failed (${res.status}): ${await res.text()}`);
  }

  const payload = await res.json();
  const rows = extractRows(payload);
  const byHour = {};

  rows.forEach((row) => {
    const hourKey = normalizeHourKey(row);
    if (!hourKey) {
      return;
    }
    const adSpend = cleanNumber(row.ad_spend);
    const roas = cleanNumber(row.roas);
    if (!Number.isFinite(adSpend) && !Number.isFinite(roas)) {
      return;
    }

    byHour[hourKey] = {
      ad_spend: Number.isFinite(adSpend) ? round(adSpend, 2) : null,
      roas: Number.isFinite(roas) ? round(roas, 4) : null,
      logged_at_local: row.logged_at_local || null,
    };
  });

  return byHour;
}

async function fetchShopifyOrdersDetailed({ domain, token, apiVersion, start, end }) {
  let nextUrl =
    `https://${domain}/admin/api/${apiVersion}/orders.json` +
    `?status=any&limit=250&order=created_at%20asc` +
    `&created_at_min=${encodeURIComponent(start.toISOString())}` +
    `&created_at_max=${encodeURIComponent(end.toISOString())}`;

  const orders = [];
  let pagesFetched = 0;
  const maxPages = 120;

  while (nextUrl && pagesFetched < maxPages) {
    const res = await fetch(nextUrl, {
      headers: {
        "X-Shopify-Access-Token": token,
      },
    });

    if (!res.ok) {
      throw new Error(`Shopify Orders API failed (${res.status}): ${(await res.text()).slice(0, 500)}`);
    }

    const payload = await res.json();
    const pageOrders = Array.isArray(payload.orders) ? payload.orders : [];
    orders.push(...pageOrders);

    const linkHeader = res.headers.get("link");
    nextUrl = parseShopifyNextLink(linkHeader);
    pagesFetched += 1;
  }

  if (pagesFetched >= maxPages) {
    throw new Error(`Shopify pagination exceeded ${maxPages} pages. Reduce SYNC_DAYS.`);
  }

  return { orders, pagesFetched };
}

function parseShopifyNextLink(linkHeader) {
  if (!linkHeader) {
    return "";
  }
  const parts = String(linkHeader).split(",");
  for (const part of parts) {
    const section = part.trim();
    if (!/rel="next"/.test(section)) {
      continue;
    }
    const match = section.match(/<([^>]+)>/);
    if (match) {
      return match[1];
    }
  }
  return "";
}

function buildShopifyStructures(orders, runAtIso) {
  const hourly = {};
  const orderRows = [];
  const lineRows = [];

  orders.forEach((order) => {
    const createdAt = new Date(order.created_at);
    if (Number.isNaN(createdAt.getTime())) {
      return;
    }
    if (!isReportableShopifyOrder(order)) {
      return;
    }

    const orderId = String(order.id);
    const hourKey = hourKeyFromDate(createdAt);
    const gross = deriveGrossSales(order);
    const net = cleanNumber(order.current_subtotal_price) ?? cleanNumber(order.subtotal_price);
    const totalSales = cleanNumber(order.current_total_price) ?? cleanNumber(order.total_price);
    const discounts = cleanNumber(order.current_total_discounts) ?? cleanNumber(order.total_discounts);
    const returnsAmount = deriveOrderRefundAmount(order);
    const sourceName = order.source_name || "unknown";
    const currency = order.currency || null;
    const customerId = order?.customer?.id ? String(order.customer.id) : null;
    const customerOrdersCount = Number(order?.customer?.orders_count || 0);
    const customerType = customerId ? (customerOrdersCount <= 1 ? "new" : "returning") : "unknown";

    if (!hourly[hourKey]) {
      hourly[hourKey] = { sales_amount: 0, orders: 0 };
    }
    if (Number.isFinite(totalSales)) {
      hourly[hourKey].sales_amount += totalSales;
    }
    hourly[hourKey].orders += 1;

    const lineItems = Array.isArray(order.line_items) ? order.line_items : [];
    const refunds = Array.isArray(order.refunds) ? order.refunds : [];
    const refundByLineId = buildRefundLineMap(refunds);
    const refundsCount = refunds.length;
    const itemsCount = lineItems.reduce((sum, item) => sum + (cleanNumber(item.quantity) || 0), 0);

    orderRows.push({
      order_id: orderId,
      order_name: order.name || null,
      created_at_utc: createdAt.toISOString(),
      processed_at_utc: order.processed_at ? new Date(order.processed_at).toISOString() : createdAt.toISOString(),
      currency,
      source_name: sourceName,
      customer_id: customerId,
      customer_type: customerType,
      gross_sales: Number.isFinite(gross) ? round(gross, 2) : null,
      net_sales: Number.isFinite(net) ? round(net, 2) : null,
      total_sales: Number.isFinite(totalSales) ? round(totalSales, 2) : null,
      discounts: Number.isFinite(discounts) ? round(discounts, 2) : null,
      returns_amount: Number.isFinite(returnsAmount) ? round(returnsAmount, 2) : null,
      refunds_count: refundsCount,
      line_items_count: round(itemsCount, 2),
      financial_status: order.financial_status || null,
      fulfillment_status: order.fulfillment_status || null,
      ingested_at_utc: runAtIso,
    });

    lineItems.forEach((line) => {
      const lineId = String(line.id);
      const quantity = cleanNumber(line.quantity) || 0;
      const unitPrice = cleanNumber(line.price) || 0;
      const grossRevenue = unitPrice * quantity;
      const discountAmount = cleanNumber(line.total_discount) || 0;
      const netRevenue = grossRevenue - discountAmount;

      const refunded = refundByLineId[lineId] || { returned_quantity: 0, returned_revenue: 0 };
      const netQuantity = Math.max(0, quantity - refunded.returned_quantity);
      const netRevenueAfterReturns = netRevenue - refunded.returned_revenue;

      lineRows.push({
        order_line_key: `${orderId}:${lineId}`,
        order_id: orderId,
        line_item_id: lineId,
        created_at_utc: createdAt.toISOString(),
        product_id: line.product_id ? String(line.product_id) : null,
        variant_id: line.variant_id ? String(line.variant_id) : null,
        sku: line.sku || null,
        product_title: line.title || null,
        variant_title: line.variant_title || null,
        vendor: line.vendor || null,
        source_name: sourceName,
        customer_type: customerType,
        quantity: round(quantity, 2),
        gross_revenue: round(grossRevenue, 2),
        discount_amount: round(discountAmount, 2),
        net_revenue: round(netRevenue, 2),
        returned_quantity: round(refunded.returned_quantity, 2),
        returned_revenue: round(refunded.returned_revenue, 2),
        net_quantity: round(netQuantity, 2),
        net_revenue_after_returns: round(netRevenueAfterReturns, 2),
        ingested_at_utc: runAtIso,
      });
    });
  });

  Object.keys(hourly).forEach((hourKey) => {
    hourly[hourKey].sales_amount = round(hourly[hourKey].sales_amount, 2);
    hourly[hourKey].orders = round(hourly[hourKey].orders, 2);
  });

  return { hourly, orderRows, lineRows };
}

function isReportableShopifyOrder(order) {
  const financialStatus = String(order?.financial_status || "").toLowerCase();
  if (financialStatus === "voided") {
    return false;
  }
  if (order?.cancelled_at) {
    return false;
  }
  if (order?.test === true) {
    return false;
  }
  return true;
}

function deriveGrossSales(order) {
  const subtotal = cleanNumber(order.current_subtotal_price) ?? cleanNumber(order.subtotal_price);
  const discounts = cleanNumber(order.current_total_discounts) ?? cleanNumber(order.total_discounts);
  if (Number.isFinite(subtotal) && Number.isFinite(discounts)) {
    return subtotal + discounts;
  }
  return cleanNumber(order.total_line_items_price);
}

function deriveOrderRefundAmount(order) {
  const refunds = Array.isArray(order.refunds) ? order.refunds : [];
  let total = 0;
  refunds.forEach((refund) => {
    const txs = Array.isArray(refund.transactions) ? refund.transactions : [];
    txs.forEach((tx) => {
      const amount = cleanNumber(tx.amount);
      if (Number.isFinite(amount)) {
        total += amount;
      }
    });
  });

  if (total > 0) {
    return total;
  }

  let fallback = 0;
  refunds.forEach((refund) => {
    const lines = Array.isArray(refund.refund_line_items) ? refund.refund_line_items : [];
    lines.forEach((rli) => {
      const subtotal = cleanNumber(rli.subtotal);
      if (Number.isFinite(subtotal)) {
        fallback += subtotal;
      }
    });
  });
  return fallback;
}

function buildRefundLineMap(refunds) {
  const map = {};
  refunds.forEach((refund) => {
    const lines = Array.isArray(refund.refund_line_items) ? refund.refund_line_items : [];
    lines.forEach((rli) => {
      const lineId = rli.line_item_id ? String(rli.line_item_id) : rli?.line_item?.id ? String(rli.line_item.id) : "";
      if (!lineId) {
        return;
      }
      const quantity = cleanNumber(rli.quantity) || 0;
      const subtotal = cleanNumber(rli.subtotal);
      let returnedRevenue = subtotal;
      if (!Number.isFinite(returnedRevenue)) {
        const linePrice = cleanNumber(rli?.line_item?.price) || 0;
        returnedRevenue = linePrice * quantity;
      }

      if (!map[lineId]) {
        map[lineId] = { returned_quantity: 0, returned_revenue: 0 };
      }
      map[lineId].returned_quantity += quantity;
      map[lineId].returned_revenue += returnedRevenue;
    });
  });
  return map;
}

async function fetchExistingRows({ supabase, start }) {
  const filter = `logged_at_utc=gte.${encodeURIComponent(start.toISOString())}`;
  return supabaseSelectAll({
    config: supabase,
    table: supabase.table,
    select: "row_key,logged_at_utc,logged_at_local,sales_amount,orders,ad_spend,roas,source_sales,source_marketing",
    orderBy: "logged_at_utc.asc",
    filter,
  });
}

function buildMergedHourlyRows({ marketingByHour, salesByHour, existingRows, runAtIso }) {
  const existingByHour = {};
  existingRows.forEach((row) => {
    const match = String(row.row_key || "").match(/^(\d{4}-\d{2}-\d{2}-\d{2})/);
    if (!match) {
      return;
    }
    existingByHour[match[1]] = row;
  });

  const allHours = new Set([
    ...Object.keys(existingByHour),
    ...Object.keys(marketingByHour),
    ...Object.keys(salesByHour),
  ]);

  const out = [];
  Array.from(allHours)
    .sort()
    .forEach((hourKey) => {
      const utcDate = utcFromHourKey(hourKey);
      if (!utcDate) {
        return;
      }

      const existing = existingByHour[hourKey] || {};
      const marketing = marketingByHour[hourKey] || {};
      const sales = salesByHour[hourKey] || {};
      const loggedUtcIso = utcDate.toISOString();
      const salesAmount = Number.isFinite(sales.sales_amount) ? sales.sales_amount : cleanNumber(existing.sales_amount);
      const orders = Number.isFinite(sales.orders) ? sales.orders : cleanNumber(existing.orders);
      const adSpend = Number.isFinite(marketing.ad_spend) ? marketing.ad_spend : cleanNumber(existing.ad_spend);
      const roas = Number.isFinite(marketing.roas) ? marketing.roas : cleanNumber(existing.roas);

      out.push({
        row_key: `${hourKey}-tw`,
        logged_at_utc: loggedUtcIso,
        logged_at_local: marketing.logged_at_local || existing.logged_at_local || loggedUtcIso,
        sales_amount: Number.isFinite(salesAmount) ? round(salesAmount, 2) : null,
        orders: Number.isFinite(orders) ? round(orders, 2) : null,
        ad_spend: Number.isFinite(adSpend) ? round(adSpend, 2) : null,
        roas: Number.isFinite(roas) ? round(roas, 4) : null,
        source_sales: Number.isFinite(salesAmount) ? "shopify" : existing.source_sales || null,
        source_marketing: Number.isFinite(adSpend) || Number.isFinite(roas) ? "apps_script" : existing.source_marketing || null,
        ingested_at_utc: runAtIso,
      });
    });

  return out;
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
