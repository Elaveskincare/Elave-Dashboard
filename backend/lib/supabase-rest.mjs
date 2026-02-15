import { requireEnv } from "./env.mjs";

export function getSupabaseConfig() {
  return {
    url: requireEnv("SUPABASE_URL").replace(/\/+$/, ""),
    serviceRoleKey: requireEnv("SUPABASE_SERVICE_ROLE_KEY"),
    table: process.env.SUPABASE_TABLE || "hourly_metrics",
    ordersTable: process.env.SUPABASE_ORDERS_TABLE || "shopify_orders",
    orderLinesTable: process.env.SUPABASE_ORDER_LINES_TABLE || "shopify_order_lines",
  };
}

export async function supabaseSelect({
  config,
  table,
  select = "*",
  orderBy = "logged_at_utc.asc",
  filter = "",
  offset = 0,
  limit = 1000,
}) {
  const tableName = table || config.table;
  const base = `${config.url}/rest/v1/${tableName}`;
  const query = [`select=${encodeURIComponent(select)}`, `order=${encodeURIComponent(orderBy)}`];
  if (filter) {
    query.push(filter);
  }

  const res = await fetch(`${base}?${query.join("&")}`, {
    headers: {
      apikey: config.serviceRoleKey,
      Authorization: `Bearer ${config.serviceRoleKey}`,
      "Range-Unit": "items",
      Range: `${Math.max(0, offset)}-${Math.max(0, offset + Math.max(1, limit) - 1)}`,
    },
  });

  if (!res.ok) {
    throw new Error(`Supabase select failed (${res.status}): ${await res.text()}`);
  }

  return res.json();
}

export async function supabaseSelectAll({
  config,
  table,
  select = "*",
  orderBy = "logged_at_utc.asc",
  filter = "",
  pageSize = 1000,
  maxPages = 200,
}) {
  const size = Math.max(1, Math.min(5000, Number(pageSize) || 1000));
  const limitPages = Math.max(1, Math.min(5000, Number(maxPages) || 200));
  const out = [];

  for (let page = 0; page < limitPages; page += 1) {
    const offset = page * size;
    const chunk = await supabaseSelect({
      config,
      table,
      select,
      orderBy,
      filter,
      offset,
      limit: size,
    });
    out.push(...chunk);
    if (!Array.isArray(chunk) || chunk.length < size) {
      return out;
    }
  }

  throw new Error(`Supabase select pagination exceeded maxPages=${limitPages} for table ${table || config.table}`);
}

export async function supabaseUpsert({ config, table, rows, onConflict = "row_key" }) {
  if (!rows.length) {
    return [];
  }
  const tableName = table || config.table;

  const res = await fetch(`${config.url}/rest/v1/${tableName}?on_conflict=${encodeURIComponent(onConflict)}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      apikey: config.serviceRoleKey,
      Authorization: `Bearer ${config.serviceRoleKey}`,
      Prefer: "resolution=merge-duplicates,return=representation",
    },
    body: JSON.stringify(rows),
  });

  if (!res.ok) {
    throw new Error(`Supabase upsert failed (${res.status}): ${await res.text()}`);
  }

  return res.json();
}
