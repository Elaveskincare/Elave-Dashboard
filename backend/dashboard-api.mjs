import http from "node:http";
import path from "node:path";
import { URL, fileURLToPath } from "node:url";
import { loadEnvFile } from "./lib/env.mjs";
import { getSupabaseConfig, supabaseSelectAll } from "./lib/supabase-rest.mjs";

loadEnvFile();

const PORT = Number(process.env.PORT || 8787);
const supabase = getSupabaseConfig();
const SHOPIFY_DOMAIN = String(process.env.SHOPIFY_STORE_DOMAIN || "").trim();
const SHOPIFY_ACCESS_TOKEN = String(process.env.SHOPIFY_ACCESS_TOKEN || "").trim();
const SHOPIFY_ORDERS_API_VERSION = process.env.SHOPIFY_API_VERSION || "2024-10";
const SHOPIFY_ANALYTICS_API_VERSION = process.env.SHOPIFY_ANALYTICS_API_VERSION || "2025-10";
const REPORTING_TIMEZONE_RAW = String(process.env.REPORTING_TIMEZONE || "UTC").trim() || "UTC";
const REPORTING_TIMEZONE = resolveReportingTimezone(REPORTING_TIMEZONE_RAW);
const SHOPIFYQL_CACHE_TTL_MS = Math.max(5000, Number(process.env.SHOPIFYQL_CACHE_TTL_MS || "60000"));
const SHOPIFY_ACCESS_SCOPES_CACHE_TTL_MS = Math.max(
  60_000,
  Number(process.env.SHOPIFY_ACCESS_SCOPES_CACHE_TTL_MS || "300000")
);
const shopifyqlCache = new Map();
let shopifyAccessScopesCache = { fetchedAt: 0, scopes: null };
const timeFormatterCache = new Map();
const GOOGLE_OAUTH_CLIENT_ID = String(
  process.env.GOOGLE_OAUTH_CLIENT_ID || process.env.GOOGLE_CLIENT_ID || ""
).trim();
const GOOGLE_OAUTH_CLIENT_SECRET = String(
  process.env.GOOGLE_OAUTH_CLIENT_SECRET || process.env.GOOGLE_CLIENT_SECRET || ""
).trim();
const GOOGLE_OAUTH_REDIRECT_URI = String(process.env.GOOGLE_OAUTH_REDIRECT_URI || "").trim();
const GOOGLE_OAUTH_SCOPES = String(
  process.env.GOOGLE_OAUTH_SCOPES || "https://www.googleapis.com/auth/calendar.readonly"
)
  .trim()
  .split(/\s+/)
  .filter(Boolean)
  .join(" ");
const GOOGLE_CALENDAR_ID = String(process.env.GOOGLE_CALENDAR_ID || "primary").trim() || "primary";
let runtimeGoogleRefreshToken = String(process.env.GOOGLE_CALENDAR_REFRESH_TOKEN || "").trim();
let runtimeGoogleAccessToken = "";
let runtimeGoogleAccessTokenExpiresAt = 0;
const GOOGLE_REFRESH_TOKEN_COOKIE_NAME = "elave_gcal_rt";
const GOOGLE_REFRESH_TOKEN_COOKIE_MAX_AGE_SECONDS = 60 * 60 * 24 * 180;

const HOURLY_SELECT =
  "row_key,logged_at_utc,logged_at_local,sales_amount,orders,ad_spend,roas,source_sales,source_marketing,ingested_at_utc";
const ORDERS_SELECT =
  "order_id,order_name,created_at_utc,processed_at_utc,currency,source_name,customer_id,customer_type,gross_sales,net_sales,total_sales,discounts,returns_amount,refunds_count,line_items_count,financial_status,fulfillment_status,ingested_at_utc";
const LINES_SELECT =
  "order_line_key,order_id,line_item_id,created_at_utc,product_id,variant_id,sku,product_title,variant_title,vendor,source_name,customer_type,quantity,gross_revenue,discount_amount,net_revenue,returned_quantity,returned_revenue,net_quantity,net_revenue_after_returns,ingested_at_utc";

const ENDPOINTS = [
  "/api/health",
  "/api/endpoints",
  "/api/clean?days=120",
  "/api/latest",
  "/api/summary",
  "/api/kpis",
  "/api/ytd",
  "/api/cells",
  "/api/trend/hourly?days=30",
  "/api/trend/daily?days=90",
  "/api/quality?days=30",
  "/api/sources?days=30",
  "/api/products/top-units?limit=10",
  "/api/products/top-revenue?limit=10",
  "/api/products/momentum?metric=revenue&limit=10",
  "/api/pace",
  "/api/projection",
  "/api/finance/gross-net-returns",
  "/api/aov",
  "/api/sessions/mtd",
  "/api/customers/new-vs-returning",
  "/api/channels",
  "/api/discount-impact",
  "/api/heatmap/today",
  "/api/refund-watchlist?limit=10",
  "/api/google/oauth/start",
  "/api/google/oauth/callback",
  "/api/google/calendar/upcoming?max=4",
];

export async function handleDashboardApiRequest(req, res) {
  try {
    const url = new URL(req.url || "/", getRequestBaseUrl(req));

    if (req.method === "OPTIONS") {
      writeJson(res, 204, {});
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/health") {
      writeJson(res, 200, {
        ok: true,
        service: "dashboard-api",
        time: new Date().toISOString(),
      });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/endpoints") {
      writeJson(res, 200, { endpoints: ENDPOINTS });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/google/oauth/start") {
      if (!hasGoogleOAuthClientConfig()) {
        writeJson(res, 503, {
          status: "google_oauth_not_configured",
          error: "Google OAuth client id/secret not configured on backend",
        });
        return;
      }

      const authUrl = buildGoogleOAuthAuthorizeUrl(url);
      res.writeHead(302, {
        Location: authUrl,
        "Cache-Control": "no-store",
      });
      res.end();
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/google/oauth/callback") {
      if (!hasGoogleOAuthClientConfig()) {
        writeHtml(
          res,
          503,
          "<h1>Google Calendar OAuth not configured</h1><p>Set GOOGLE_OAUTH_CLIENT_ID and GOOGLE_OAUTH_CLIENT_SECRET on the backend.</p>"
        );
        return;
      }

      const oauthError = String(url.searchParams.get("error") || "").trim();
      if (oauthError) {
        writeHtml(res, 400, `<h1>Google authorization failed</h1><p>${escapeHtml(oauthError)}</p>`);
        return;
      }

      const code = String(url.searchParams.get("code") || "").trim();
      if (!code) {
        writeHtml(res, 400, "<h1>Missing OAuth code</h1><p>Google did not return an authorization code.</p>");
        return;
      }

      try {
        const tokenPayload = await exchangeGoogleCodeForTokens({
          code,
          redirectUri: getGoogleRedirectUri(url),
        });
        const accessToken = String(tokenPayload.access_token || "").trim();
        const expiresIn = Number(tokenPayload.expires_in || 0);
        const refreshToken = String(tokenPayload.refresh_token || "").trim();

        if (!accessToken) {
          throw new Error("Google token response did not include an access token");
        }

        if (refreshToken) {
          runtimeGoogleRefreshToken = refreshToken;
        }

        runtimeGoogleAccessToken = accessToken;
        runtimeGoogleAccessTokenExpiresAt = Date.now() + Math.max(30, expiresIn) * 1000;

        const persistedHint = refreshToken
          ? `<p>Connected.</p><p>This browser now has a secure cookie for Calendar access. For all devices and cold starts, set <code>GOOGLE_CALENDAR_REFRESH_TOKEN</code> in Vercel using this value, then redeploy:</p><textarea readonly style="width:100%;min-height:88px;font-family:monospace;padding:10px;">${escapeHtml(
              refreshToken
            )}</textarea>`
          : runtimeGoogleRefreshToken
            ? "<p>Connected using an existing refresh token.</p><p>For reliable access on cold starts, set <code>GOOGLE_CALENDAR_REFRESH_TOKEN</code> in Vercel env.</p>"
            : "<p>Connected for this runtime only. Re-auth may be needed after restart.</p>";

        const responseHeaders = refreshToken
          ? {
              "Set-Cookie": buildGoogleRefreshTokenCookie(refreshToken),
            }
          : {};

        writeHtml(
          res,
          200,
          `<h1>Google Calendar connected</h1>${persistedHint}<p>You can close this tab and refresh the dashboard.</p>`,
          responseHeaders
        );
      } catch (error) {
        writeHtml(
          res,
          500,
          `<h1>Google OAuth token exchange failed</h1><p>${escapeHtml(
            String(error && error.message ? error.message : error)
          )}</p>`
        );
      }
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/google/calendar/upcoming") {
      if (!runtimeGoogleRefreshToken) {
        const requestRefreshToken = readGoogleRefreshTokenFromRequest(req);
        if (requestRefreshToken) {
          runtimeGoogleRefreshToken = requestRefreshToken;
        }
      }

      const maxResults = parseLimit(url, 4, 10);
      const calendarId = sanitizeCalendarId(url.searchParams.get("calendarId") || GOOGLE_CALENDAR_ID);
      try {
        const payload = await getGoogleCalendarUpcoming({ calendarId, maxResults });
        writeJson(res, 200, {
          updatedAt: new Date().toISOString(),
          calendar_id: calendarId,
          time_zone: payload.timeZone || null,
          events: payload.events,
        });
      } catch (error) {
        const errCode = String(error && error.code ? error.code : "");
        const message = String(error && error.message ? error.message : error || "Google Calendar request failed");
        if (errCode === "google_oauth_not_configured") {
          writeJson(res, 503, {
            status: errCode,
            error: message,
          });
          return;
        }
        if (errCode === "google_auth_required") {
          writeJson(res, 401, {
            status: errCode,
            error: message,
            auth_url: hasGoogleOAuthClientConfig() ? buildGoogleOAuthAuthorizeUrl(url) : null,
          });
          return;
        }
        writeJson(res, 502, {
          status: "google_calendar_error",
          error: message,
        });
      }
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/clean") {
      const days = parseDays(url, 120, 365);
      const rows = await fetchHourlyRowsSinceDays(days);
      writeJson(res, 200, {
        updatedAt: new Date().toISOString(),
        days,
        rowCount: rows.length,
        rows,
        data: rows,
      });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/latest") {
      const rows = await fetchHourlyRowsSinceDays(7);
      const latest = rows.length ? rows[rows.length - 1] : null;
      writeJson(res, 200, {
        updatedAt: new Date().toISOString(),
        latest,
      });
      return;
    }

    if (req.method === "GET" && (url.pathname === "/api/summary" || url.pathname === "/api/kpis")) {
      const payload = await buildSummaryPayload();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/ytd") {
      const payload = await getYtdComparisonPayload();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/cells") {
      const payload = await buildCellsPayload();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/trend/hourly") {
      const days = parseDays(url, 30, 365);
      const rows = await fetchHourlyRowsSinceDays(days);
      const series = rows.map((row) => ({
        logged_at_utc: row.logged_at_utc,
        logged_at_local: row.logged_at_local,
        sales_amount: row.sales_amount,
        orders: row.orders,
        ad_spend: row.ad_spend,
        roas: row.roas,
        aov: row.orders && row.orders !== 0 ? round(row.sales_amount / row.orders, 2) : null,
      }));
      writeJson(res, 200, {
        updatedAt: new Date().toISOString(),
        days,
        points: series.length,
        series,
      });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/trend/daily") {
      const days = parseDays(url, 90, 730);
      const rows = await fetchHourlyRowsSinceDays(days);
      const byDay = {};
      rows.forEach((row) => {
        const day = String(row.logged_at_utc || "").slice(0, 10);
        if (!day) {
          return;
        }
        if (!byDay[day]) {
          byDay[day] = { day, sales_amount: 0, orders: 0, ad_spend: 0, rows: 0 };
        }
        byDay[day].sales_amount += row.sales_amount || 0;
        byDay[day].orders += row.orders || 0;
        byDay[day].ad_spend += row.ad_spend || 0;
        byDay[day].rows += 1;
      });

      const series = Object.keys(byDay)
        .sort()
        .map((day) => {
          const item = byDay[day];
          const aov = item.orders > 0 ? item.sales_amount / item.orders : null;
          const roas = item.ad_spend > 0 ? item.sales_amount / item.ad_spend : null;
          return {
            day,
            sales_amount: round(item.sales_amount, 2),
            orders: round(item.orders, 2),
            ad_spend: round(item.ad_spend, 2),
            aov: round(aov, 2),
            roas: round(roas, 4),
            rows: item.rows,
          };
        });

      writeJson(res, 200, {
        updatedAt: new Date().toISOString(),
        days,
        points: series.length,
        series,
      });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/quality") {
      const days = parseDays(url, 30, 365);
      const rows = await fetchHourlyRowsSinceDays(days);
      writeJson(res, 200, {
        updatedAt: new Date().toISOString(),
        days,
        quality: buildQuality(rows),
      });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/sources") {
      const days = parseDays(url, 30, 365);
      const rows = await fetchHourlyRowsSinceDays(days);
      writeJson(res, 200, {
        updatedAt: new Date().toISOString(),
        days,
        sources: buildSourceCoverage(rows),
      });
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/products/top-units") {
      const limit = parseLimit(url, 10, 50);
      const payload = await getTopProductsByUnits(limit);
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/products/top-revenue") {
      const limit = parseLimit(url, 10, 50);
      const payload = await getTopProductsByRevenue(limit);
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/products/momentum") {
      const limit = parseLimit(url, 10, 50);
      const metric = String(url.searchParams.get("metric") || "revenue").toLowerCase() === "units" ? "units" : "revenue";
      const payload = await getProductMomentum({ limit, metric });
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/pace") {
      const payload = await getDailySalesPace();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/projection") {
      const payload = await getMtdProjection();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/finance/gross-net-returns") {
      const payload = await getGrossNetReturns();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/aov") {
      const payload = await getAovPayload();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/sessions/mtd") {
      const payload = await getWebsiteSessionsMtdPayload();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/customers/new-vs-returning") {
      const payload = await getNewVsReturningPayload();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/channels") {
      const payload = await getChannelSplitPayload();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/discount-impact") {
      const payload = await getDiscountImpactPayload();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/heatmap/today") {
      const payload = await getHourlyHeatmapToday();
      writeJson(res, 200, payload);
      return;
    }

    if (req.method === "GET" && url.pathname === "/api/refund-watchlist") {
      const limit = parseLimit(url, 10, 50);
      const payload = await getRefundWatchlist(limit);
      writeJson(res, 200, payload);
      return;
    }

    writeJson(res, 404, { error: "Not found" });
  } catch (error) {
    writeJson(res, 500, { error: String(error && error.message ? error.message : error) });
  }
}

function isDirectExecution() {
  if (!process.argv[1]) {
    return false;
  }
  try {
    return path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url));
  } catch (_error) {
    return false;
  }
}

if (isDirectExecution()) {
  const server = http.createServer(handleDashboardApiRequest);
  server.listen(PORT, () => {
    console.log(`Dashboard API listening on http://localhost:${PORT}`);
  });
}

async function buildCellsPayload() {
  const jobs = {
    summary: () => buildSummaryPayload(),
    ytdComparison: () => getYtdComparisonPayload(),
    topUnits: () => getTopProductsByUnits(10),
    topRevenue: () => getTopProductsByRevenue(10),
    momentum: () => getProductMomentum({ limit: 10, metric: "revenue" }),
    pace: () => getDailySalesPace(),
    projection: () => getMtdProjection(),
    grossNetReturns: () => getGrossNetReturns(),
    aov: () => getAovPayload(),
    websiteSessionsMtd: () => getWebsiteSessionsMtdPayload(),
    newVsReturning: () => getNewVsReturningPayload(),
    channels: () => getChannelSplitPayload(),
    discountImpact: () => getDiscountImpactPayload(),
    heatmap: () => getHourlyHeatmapToday(),
    refundWatchlist: () => getRefundWatchlist(10),
  };

  const entries = Object.entries(jobs);
  const settled = await Promise.allSettled(entries.map(([, fn]) => fn()));
  const resultMap = {};
  const errors = {};

  settled.forEach((result, idx) => {
    const key = entries[idx][0];
    if (result.status === "fulfilled") {
      resultMap[key] = result.value;
      return;
    }
    resultMap[key] = null;
    errors[key] = String(result.reason && result.reason.message ? result.reason.message : result.reason || "Unknown error");
    console.warn(`Cells payload section failed: ${key}`, errors[key]);
  });

  const summary = resultMap.summary;

  return {
    updatedAt: new Date().toISOString(),
    summary: summary && summary.summary ? summary.summary : null,
    kpis: summary && summary.kpis ? summary.kpis : null,
    ytd_comparison: resultMap.ytdComparison,
    top_products_units: resultMap.topUnits,
    top_products_revenue: resultMap.topRevenue,
    product_momentum: resultMap.momentum,
    daily_sales_pace: resultMap.pace,
    mtd_projection: resultMap.projection,
    gross_net_returns: resultMap.grossNetReturns,
    aov: resultMap.aov,
    website_sessions_mtd: resultMap.websiteSessionsMtd,
    new_vs_returning: resultMap.newVsReturning,
    channel_split: resultMap.channels,
    discount_impact: resultMap.discountImpact,
    hourly_heatmap_today: resultMap.heatmap,
    refund_watchlist: resultMap.refundWatchlist,
    errors,
  };
}

function writeJson(res, status, payload) {
  res.writeHead(status, {
    "Content-Type": "application/json",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type,Authorization",
  });
  res.end(JSON.stringify(payload));
}

function writeHtml(res, status, html, extraHeaders = {}) {
  res.writeHead(status, {
    "Content-Type": "text/html; charset=utf-8",
    "Cache-Control": "no-store",
    ...extraHeaders,
  });
  res.end(`<!doctype html><html><head><meta charset="utf-8" /><title>Google Calendar OAuth</title></head><body style="font-family:Arial,sans-serif;padding:20px;line-height:1.45;">${html}</body></html>`);
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getRequestBaseUrl(req) {
  const forwardedProto = getForwardedHeaderValue(req?.headers?.["x-forwarded-proto"]);
  const forwardedHost = getForwardedHeaderValue(req?.headers?.["x-forwarded-host"]);
  const host = forwardedHost || getForwardedHeaderValue(req?.headers?.host) || "localhost";
  const protocol = normalizeRequestProtocol(forwardedProto, host);
  return `${protocol}://${host}`;
}

function getForwardedHeaderValue(value) {
  const raw = Array.isArray(value) ? value[0] : value;
  const text = String(raw || "").trim();
  if (!text) {
    return "";
  }
  return text.split(",")[0].trim();
}

function normalizeRequestProtocol(candidate, host) {
  const protocol = String(candidate || "").trim().toLowerCase();
  if (protocol === "http" || protocol === "https") {
    return protocol;
  }
  return isLocalHost(host) ? "http" : "https";
}

function isLocalHost(host) {
  const hostname = normalizeHostname(host);
  return hostname === "localhost" || hostname === "127.0.0.1" || hostname === "[::1]" || hostname === "::1";
}

function normalizeHostname(host) {
  const raw = String(host || "").trim();
  if (!raw) {
    return "";
  }
  try {
    return new URL(`http://${raw}`).hostname.toLowerCase();
  } catch (_error) {
    return raw.replace(/:\d+$/, "").toLowerCase();
  }
}

function buildGoogleRefreshTokenCookie(refreshToken) {
  const token = encodeURIComponent(String(refreshToken || "").trim());
  return `${GOOGLE_REFRESH_TOKEN_COOKIE_NAME}=${token}; Path=/; Max-Age=${GOOGLE_REFRESH_TOKEN_COOKIE_MAX_AGE_SECONDS}; HttpOnly; Secure; SameSite=Lax`;
}

function readGoogleRefreshTokenFromRequest(req) {
  const cookies = parseCookieHeader(req?.headers?.cookie);
  const token = String(cookies[GOOGLE_REFRESH_TOKEN_COOKIE_NAME] || "").trim();
  return token;
}

function parseCookieHeader(rawCookieHeader) {
  const header = String(rawCookieHeader || "");
  if (!header) {
    return {};
  }

  const cookies = {};
  header.split(";").forEach((part) => {
    const entry = String(part || "").trim();
    if (!entry) {
      return;
    }
    const eq = entry.indexOf("=");
    if (eq <= 0) {
      return;
    }
    const key = entry.slice(0, eq).trim();
    const value = entry.slice(eq + 1).trim();
    if (!key) {
      return;
    }
    try {
      cookies[key] = decodeURIComponent(value);
    } catch (_error) {
      cookies[key] = value;
    }
  });

  return cookies;
}

function hasGoogleOAuthClientConfig() {
  return Boolean(GOOGLE_OAUTH_CLIENT_ID && GOOGLE_OAUTH_CLIENT_SECRET);
}

function getGoogleRedirectUri(requestUrl) {
  const fallbackRedirect = `${requestUrl.origin}/api/google/oauth/callback`;
  if (!GOOGLE_OAUTH_REDIRECT_URI) {
    return fallbackRedirect;
  }
  try {
    const configured = new URL(GOOGLE_OAUTH_REDIRECT_URI);
    if (!isLocalHost(requestUrl.hostname) && isLocalHost(configured.hostname)) {
      return fallbackRedirect;
    }
    return configured.toString();
  } catch (_error) {
    return fallbackRedirect;
  }
}

function buildGoogleOAuthAuthorizeUrl(requestUrl) {
  const url = new URL("https://accounts.google.com/o/oauth2/v2/auth");
  url.searchParams.set("client_id", GOOGLE_OAUTH_CLIENT_ID);
  url.searchParams.set("redirect_uri", getGoogleRedirectUri(requestUrl));
  url.searchParams.set("response_type", "code");
  url.searchParams.set("scope", GOOGLE_OAUTH_SCOPES);
  url.searchParams.set("access_type", "offline");
  url.searchParams.set("include_granted_scopes", "true");
  url.searchParams.set("prompt", "consent");
  return url.toString();
}

function sanitizeCalendarId(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "primary";
  }
  return text.slice(0, 200);
}

async function exchangeGoogleCodeForTokens({ code, redirectUri }) {
  const tokenUrl = "https://oauth2.googleapis.com/token";
  const body = new URLSearchParams({
    code,
    client_id: GOOGLE_OAUTH_CLIENT_ID,
    client_secret: GOOGLE_OAUTH_CLIENT_SECRET,
    redirect_uri: redirectUri,
    grant_type: "authorization_code",
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const payload = await safeParseJson(res);
  if (!res.ok) {
    const message = payload?.error_description || payload?.error || `Google OAuth token exchange failed (${res.status})`;
    throw new Error(message);
  }
  return payload || {};
}

async function refreshGoogleAccessToken() {
  if (!hasGoogleOAuthClientConfig()) {
    const error = new Error("Google OAuth client credentials are not configured");
    error.code = "google_oauth_not_configured";
    throw error;
  }
  if (!runtimeGoogleRefreshToken) {
    const error = new Error("Google Calendar requires authorization. Visit /api/google/oauth/start first.");
    error.code = "google_auth_required";
    throw error;
  }

  const tokenUrl = "https://oauth2.googleapis.com/token";
  const body = new URLSearchParams({
    client_id: GOOGLE_OAUTH_CLIENT_ID,
    client_secret: GOOGLE_OAUTH_CLIENT_SECRET,
    refresh_token: runtimeGoogleRefreshToken,
    grant_type: "refresh_token",
  });

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const payload = await safeParseJson(res);
  if (!res.ok) {
    const message = payload?.error_description || payload?.error || `Google OAuth refresh failed (${res.status})`;
    if (String(payload?.error || "").toLowerCase() === "invalid_grant") {
      runtimeGoogleAccessToken = "";
      runtimeGoogleAccessTokenExpiresAt = 0;
      runtimeGoogleRefreshToken = "";
      const authError = new Error("Google Calendar authorization expired. Reconnect via /api/google/oauth/start.");
      authError.code = "google_auth_required";
      throw authError;
    }
    throw new Error(message);
  }

  const accessToken = String(payload?.access_token || "").trim();
  const expiresIn = Number(payload?.expires_in || 0);
  if (!accessToken) {
    throw new Error("Google OAuth refresh response missing access token");
  }

  runtimeGoogleAccessToken = accessToken;
  runtimeGoogleAccessTokenExpiresAt = Date.now() + Math.max(30, expiresIn) * 1000;
  return runtimeGoogleAccessToken;
}

async function getGoogleAccessToken() {
  const nowMs = Date.now();
  if (runtimeGoogleAccessToken && nowMs + 30_000 < runtimeGoogleAccessTokenExpiresAt) {
    return runtimeGoogleAccessToken;
  }
  return refreshGoogleAccessToken();
}

async function safeParseJson(res) {
  const text = await res.text();
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text);
  } catch (_error) {
    return { raw: text };
  }
}

function normalizeGoogleCalendarEvent(item) {
  const title = String(item?.summary || "").trim() || "(Untitled)";
  const startDateTime = typeof item?.start?.dateTime === "string" ? item.start.dateTime : "";
  const endDateTime = typeof item?.end?.dateTime === "string" ? item.end.dateTime : "";
  const startDate = typeof item?.start?.date === "string" ? item.start.date : "";
  const endDate = typeof item?.end?.date === "string" ? item.end.date : "";
  const start = startDateTime || startDate || "";
  const end = endDateTime || endDate || "";
  if (!start) {
    return null;
  }
  return {
    id: String(item?.id || ""),
    title,
    start,
    end,
    is_all_day: Boolean(startDate && !startDateTime),
    meet_link: item?.hangoutLink || null,
    location: item?.location || null,
    html_link: item?.htmlLink || null,
  };
}

async function getGoogleCalendarUpcoming({ calendarId, maxResults }) {
  const token = await getGoogleAccessToken();
  const endpoint = new URL(
    `https://www.googleapis.com/calendar/v3/calendars/${encodeURIComponent(calendarId)}/events`
  );
  endpoint.searchParams.set("singleEvents", "true");
  endpoint.searchParams.set("orderBy", "startTime");
  endpoint.searchParams.set("timeMin", new Date().toISOString());
  endpoint.searchParams.set("maxResults", String(maxResults));
  endpoint.searchParams.set("fields", "timeZone,items(id,summary,status,start,end,hangoutLink,location,htmlLink)");

  const res = await fetch(endpoint, {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  });

  const payload = await safeParseJson(res);
  if (!res.ok) {
    if (res.status === 401 || res.status === 403) {
      runtimeGoogleAccessToken = "";
      runtimeGoogleAccessTokenExpiresAt = 0;
      const error = new Error("Google Calendar authorization required");
      error.code = "google_auth_required";
      throw error;
    }
    const message =
      payload?.error?.message ||
      payload?.error_description ||
      payload?.error ||
      `Google Calendar API request failed (${res.status})`;
    throw new Error(message);
  }

  const events = Array.isArray(payload?.items)
    ? payload.items
        .filter((item) => String(item?.status || "").toLowerCase() !== "cancelled")
        .map(normalizeGoogleCalendarEvent)
        .filter(Boolean)
    : [];

  return {
    timeZone: payload?.timeZone || null,
    events,
  };
}

function parseDays(url, fallback, max) {
  const raw = Number(url.searchParams.get("days") || String(fallback));
  if (!Number.isFinite(raw)) {
    return fallback;
  }
  return Math.max(1, Math.min(max, Math.floor(raw)));
}

function parseLimit(url, fallback, max) {
  const raw = Number(url.searchParams.get("limit") || String(fallback));
  if (!Number.isFinite(raw)) {
    return fallback;
  }
  return Math.max(1, Math.min(max, Math.floor(raw)));
}

function resolveReportingTimezone(candidate) {
  const value = String(candidate || "").trim() || "UTC";
  try {
    new Intl.DateTimeFormat("en-US", { timeZone: value }).format(new Date());
    return value;
  } catch (_error) {
    console.warn(`Invalid REPORTING_TIMEZONE="${value}", falling back to UTC.`);
    return "UTC";
  }
}

function nowUtc() {
  return new Date();
}

function getFormatter(timeZone) {
  const key = timeZone || "UTC";
  if (!timeFormatterCache.has(key)) {
    timeFormatterCache.set(
      key,
      new Intl.DateTimeFormat("en-CA", {
        timeZone: key,
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: false,
        hourCycle: "h23",
      })
    );
  }
  return timeFormatterCache.get(key);
}

function getZonedParts(date, timeZone = REPORTING_TIMEZONE) {
  const dtf = getFormatter(timeZone);
  const parts = dtf.formatToParts(date);
  const values = {};
  parts.forEach((part) => {
    if (part.type !== "literal") {
      values[part.type] = part.value;
    }
  });

  return {
    year: Number(values.year),
    month: Number(values.month),
    day: Number(values.day),
    hour: Number(values.hour),
    minute: Number(values.minute),
    second: Number(values.second),
  };
}

function getTimezoneOffsetMinutes(date, timeZone = REPORTING_TIMEZONE) {
  const parts = getZonedParts(date, timeZone);
  const asUtc = Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second);
  return (asUtc - date.getTime()) / (60 * 1000);
}

function zonedDateTimeToUtc({ year, month, day, hour = 0, minute = 0, second = 0 }, timeZone = REPORTING_TIMEZONE) {
  let utcMs = Date.UTC(year, month - 1, day, hour, minute, second, 0);
  for (let i = 0; i < 3; i += 1) {
    const offsetMinutes = getTimezoneOffsetMinutes(new Date(utcMs), timeZone);
    const adjusted = Date.UTC(year, month - 1, day, hour, minute, second, 0) - offsetMinutes * 60 * 1000;
    if (adjusted === utcMs) {
      break;
    }
    utcMs = adjusted;
  }
  return new Date(utcMs);
}

function startOfUtcMonth(date) {
  const parts = getZonedParts(date);
  return zonedDateTimeToUtc({
    year: parts.year,
    month: parts.month,
    day: 1,
  });
}

function startOfUtcYear(date) {
  const parts = getZonedParts(date);
  return zonedDateTimeToUtc({
    year: parts.year,
    month: 1,
    day: 1,
  });
}

function addUtcYears(date, delta) {
  const parts = getZonedParts(date);
  return zonedDateTimeToUtc({
    year: parts.year + delta,
    month: 1,
    day: 1,
  });
}

function addUtcMonths(date, delta) {
  const parts = getZonedParts(date);
  return zonedDateTimeToUtc({
    year: parts.year,
    month: parts.month + delta,
    day: 1,
  });
}

function startOfUtcDay(date) {
  const parts = getZonedParts(date);
  return zonedDateTimeToUtc({
    year: parts.year,
    month: parts.month,
    day: parts.day,
  });
}

function daysInUtcMonth(date) {
  const parts = getZonedParts(date);
  return new Date(Date.UTC(parts.year, parts.month, 0)).getUTCDate();
}

function addUtcDays(date, delta) {
  const parts = getZonedParts(date);
  return zonedDateTimeToUtc({
    year: parts.year,
    month: parts.month,
    day: parts.day + delta,
  });
}

function dayOfMonthInReportingZone(date) {
  return getZonedParts(date).day;
}

function previousMtdComparableEnd(now, currentMonthStart = startOfUtcMonth(now)) {
  const prevMonthStart = addUtcMonths(currentMonthStart, -1);
  const prevMonthDays = daysInUtcMonth(prevMonthStart);
  const comparableDay = Math.min(dayOfMonthInReportingZone(now), prevMonthDays);
  const prevParts = getZonedParts(prevMonthStart);
  const comparableDayStart = zonedDateTimeToUtc({
    year: prevParts.year,
    month: prevParts.month,
    day: comparableDay,
  });
  const nextDayStart = addUtcDays(comparableDayStart, 1);
  return new Date(nextDayStart.getTime() - 1);
}

function previousYtdComparableEnd(now) {
  const currentParts = getZonedParts(now);
  const previousYear = currentParts.year - 1;
  const previousMonthDays = daysInYearMonth(previousYear, currentParts.month);
  const comparableDay = Math.min(currentParts.day, previousMonthDays);

  return zonedDateTimeToUtc(
    {
      year: previousYear,
      month: currentParts.month,
      day: comparableDay,
      hour: currentParts.hour,
      minute: currentParts.minute,
      second: currentParts.second,
    },
    REPORTING_TIMEZONE
  );
}

function daysInYearMonth(year, month) {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

function toYmd(date) {
  const parts = getZonedParts(date);
  const y = parts.year;
  const m = String(parts.month).padStart(2, "0");
  const d = String(parts.day).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function pctChange(current, previous) {
  if (!Number.isFinite(current) || !Number.isFinite(previous) || previous === 0) {
    return null;
  }
  return round(((current - previous) / Math.abs(previous)) * 100, 2);
}

function round(value, digits) {
  if (!Number.isFinite(value)) {
    return null;
  }
  const p = 10 ** digits;
  return Math.round(value * p) / p;
}

function toNum(value) {
  if (value === null || typeof value === "undefined" || value === "") {
    return null;
  }
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function getShopifyqlCache(key) {
  const hit = shopifyqlCache.get(key);
  if (!hit) {
    return null;
  }
  if (Date.now() - hit.cachedAt > SHOPIFYQL_CACHE_TTL_MS) {
    shopifyqlCache.delete(key);
    return null;
  }
  return hit.value;
}

function setShopifyqlCache(key, value) {
  shopifyqlCache.set(key, { cachedAt: Date.now(), value });
}

async function fetchShopifyqlTable(queryText) {
  if (!SHOPIFY_DOMAIN || !SHOPIFY_ACCESS_TOKEN) {
    return null;
  }

  const cacheKey = `${SHOPIFY_ANALYTICS_API_VERSION}:${queryText}`;
  const cached = getShopifyqlCache(cacheKey);
  if (cached) {
    return cached;
  }

  const url = `https://${SHOPIFY_DOMAIN}/admin/api/${SHOPIFY_ANALYTICS_API_VERSION}/graphql.json`;
  const gql = `query RunShopifyQL($query: String!) {
    shopifyqlQuery(query: $query) {
      parseErrors
      tableData {
        columns {
          name
          displayName
          dataType
          subType
        }
        rows
      }
    }
  }`;

  const res = await fetch(url, {
    method: "POST",
    headers: {
      "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      query: gql,
      variables: { query: queryText },
    }),
  });

  if (!res.ok) {
    throw new Error(`ShopifyQL request failed (${res.status}): ${(await res.text()).slice(0, 500)}`);
  }

  const payload = await res.json();
  if (Array.isArray(payload.errors) && payload.errors.length) {
    throw new Error(`ShopifyQL GraphQL error: ${payload.errors.map((err) => err.message).join("; ")}`);
  }

  const response = payload?.data?.shopifyqlQuery;
  const parseErrors = Array.isArray(response?.parseErrors) ? response.parseErrors : [];
  if (parseErrors.length) {
    throw new Error(`ShopifyQL parse error: ${parseErrors.join("; ")}`);
  }

  const tableData = response?.tableData || null;
  const columns = Array.isArray(tableData?.columns) ? tableData.columns : [];
  const columnIndex = {};
  columns.forEach((col, idx) => {
    const name = String(col?.name || "").trim();
    if (name) {
      columnIndex[name] = idx;
    }
  });

  const out = {
    columns,
    rows: Array.isArray(tableData?.rows) ? tableData.rows : [],
    columnIndex,
  };
  setShopifyqlCache(cacheKey, out);
  return out;
}

function readShopifyqlCell(row, columnIndex, key) {
  if (!row) {
    return null;
  }
  if (typeof row === "object" && !Array.isArray(row) && Object.prototype.hasOwnProperty.call(row, key)) {
    return row[key];
  }
  if (Array.isArray(row)) {
    const idx = columnIndex[key];
    if (Number.isInteger(idx)) {
      return row[idx];
    }
  }
  return null;
}

function parseShopifyMetric(value) {
  const n = Number(value);
  return Number.isFinite(n) ? round(n, 2) : null;
}

function errorMessage(error, fallback = "Unknown error") {
  if (!error) {
    return fallback;
  }
  if (typeof error === "string") {
    return error || fallback;
  }
  if (typeof error.message === "string" && error.message.trim()) {
    return error.message.trim();
  }
  return fallback;
}

function pickShopifyqlMetricColumn(columnIndex, candidates) {
  if (!columnIndex || typeof columnIndex !== "object" || !Array.isArray(candidates)) {
    return "";
  }
  for (const key of candidates) {
    if (Number.isInteger(columnIndex[key])) {
      return key;
    }
  }
  return "";
}

async function getShopifySalesMonthSnapshot(now) {
  if (!SHOPIFY_DOMAIN || !SHOPIFY_ACCESS_TOKEN) {
    return null;
  }

  const monthStart = startOfUtcMonth(now);
  const prevMonthStart = addUtcMonths(monthStart, -1);
  const nextMonthStart = addUtcMonths(monthStart, 1);
  const queryText =
    `FROM sales SHOW gross_sales, net_sales, total_sales GROUP BY month ` +
    `SINCE ${toYmd(prevMonthStart)} UNTIL ${toYmd(nextMonthStart)}`;

  const table = await fetchShopifyqlTable(queryText);
  if (!table) {
    return null;
  }

  const byMonth = {};
  table.rows.forEach((row) => {
    const monthRaw = readShopifyqlCell(row, table.columnIndex, "month");
    const monthKey = String(monthRaw || "").slice(0, 10);
    if (!monthKey) {
      return;
    }
    byMonth[monthKey] = {
      gross_sales: parseShopifyMetric(readShopifyqlCell(row, table.columnIndex, "gross_sales")),
      net_sales: parseShopifyMetric(readShopifyqlCell(row, table.columnIndex, "net_sales")),
      total_sales: parseShopifyMetric(readShopifyqlCell(row, table.columnIndex, "total_sales")),
    };
  });

  const currentKey = toYmd(monthStart);
  const previousKey = toYmd(prevMonthStart);
  return {
    source: "shopifyql",
    current: byMonth[currentKey] || null,
    previous: byMonth[previousKey] || null,
    range: {
      since: toYmd(prevMonthStart),
      until: toYmd(nextMonthStart),
    },
  };
}

async function getShopifySalesComparableSnapshot(now) {
  if (!SHOPIFY_DOMAIN || !SHOPIFY_ACCESS_TOKEN) {
    return null;
  }

  const monthStart = startOfUtcMonth(now);
  const prevMonthStart = addUtcMonths(monthStart, -1);
  const prevComparableEnd = previousMtdComparableEnd(now, monthStart);
  const tomorrow = addUtcDays(startOfUtcDay(now), 1);
  const queryText =
    `FROM sales SHOW total_sales, net_sales, orders GROUP BY day ` +
    `SINCE ${toYmd(prevMonthStart)} UNTIL ${toYmd(tomorrow)}`;

  const table = await fetchShopifyqlTable(queryText);
  if (!table) {
    return null;
  }

  const currentMonthPrefix = toYmd(monthStart).slice(0, 7);
  const prevMonthPrefix = toYmd(prevMonthStart).slice(0, 7);
  const prevComparableDay = dayOfMonthInReportingZone(prevComparableEnd);
  let currentMtd = 0;
  let currentNetMtd = 0;
  let currentOrders = 0;
  let previousComparable = 0;
  let previousNetComparable = 0;
  let previousOrders = 0;

  table.rows.forEach((row) => {
    const dayRaw = readShopifyqlCell(row, table.columnIndex, "day");
    const dayKey = String(dayRaw || "").slice(0, 10);
    const sales = Number(readShopifyqlCell(row, table.columnIndex, "total_sales"));
    const netSales = Number(readShopifyqlCell(row, table.columnIndex, "net_sales"));
    const orders = Number(readShopifyqlCell(row, table.columnIndex, "orders"));
    if (!dayKey) {
      return;
    }

    if (dayKey.slice(0, 7) === currentMonthPrefix) {
      if (Number.isFinite(sales)) {
        currentMtd += sales;
      }
      if (Number.isFinite(netSales)) {
        currentNetMtd += netSales;
      }
      if (Number.isFinite(orders)) {
        currentOrders += orders;
      }
      return;
    }

    if (dayKey.slice(0, 7) === prevMonthPrefix) {
      const dayOfMonth = Number(dayKey.slice(8, 10));
      if (Number.isFinite(dayOfMonth) && dayOfMonth <= prevComparableDay) {
        if (Number.isFinite(sales)) {
          previousComparable += sales;
        }
        if (Number.isFinite(netSales)) {
          previousNetComparable += netSales;
        }
        if (Number.isFinite(orders)) {
          previousOrders += orders;
        }
      }
    }
  });

  return {
    source: "shopifyql",
    current_mtd: round(currentMtd, 2),
    current_mtd_net_sales: round(currentNetMtd, 2),
    current_mtd_orders: round(currentOrders, 2),
    previous_mtd: round(previousComparable, 2),
    previous_mtd_net_sales: round(previousNetComparable, 2),
    previous_mtd_orders: round(previousOrders, 2),
    range: {
      since: toYmd(prevMonthStart),
      until: toYmd(tomorrow),
      previous_end_utc: prevComparableEnd.toISOString(),
    },
  };
}

async function getShopifySalesYtdComparableSnapshot(now) {
  if (!SHOPIFY_DOMAIN || !SHOPIFY_ACCESS_TOKEN) {
    return null;
  }

  const currentYearStart = startOfUtcYear(now);
  const previousYearStart = addUtcYears(currentYearStart, -1);
  const previousComparableEnd = previousYtdComparableEnd(now);
  const tomorrow = addUtcDays(startOfUtcDay(now), 1);
  const queryText =
    `FROM sales SHOW total_sales, orders GROUP BY day ` +
    `SINCE ${toYmd(previousYearStart)} UNTIL ${toYmd(tomorrow)}`;

  const table = await fetchShopifyqlTable(queryText);
  if (!table) {
    return null;
  }

  const currentYearPrefix = toYmd(currentYearStart).slice(0, 4);
  const previousYearPrefix = toYmd(previousYearStart).slice(0, 4);
  const currentComparableYmd = toYmd(now);
  const previousComparableYmd = toYmd(previousComparableEnd);
  let currentYtdSales = 0;
  let previousYtdSales = 0;
  let currentYtdOrders = 0;
  let previousYtdOrders = 0;
  let previousFullYearSales = 0;
  let previousFullYearOrders = 0;

  table.rows.forEach((row) => {
    const dayRaw = readShopifyqlCell(row, table.columnIndex, "day");
    const dayKey = String(dayRaw || "").slice(0, 10);
    const sales = Number(readShopifyqlCell(row, table.columnIndex, "total_sales"));
    const orders = Number(readShopifyqlCell(row, table.columnIndex, "orders"));
    if (!dayKey) {
      return;
    }

    if (dayKey.slice(0, 4) === currentYearPrefix && dayKey <= currentComparableYmd) {
      if (Number.isFinite(sales)) {
        currentYtdSales += sales;
      }
      if (Number.isFinite(orders)) {
        currentYtdOrders += orders;
      }
      return;
    }

    if (dayKey.slice(0, 4) === previousYearPrefix) {
      if (Number.isFinite(sales)) {
        previousFullYearSales += sales;
      }
      if (Number.isFinite(orders)) {
        previousFullYearOrders += orders;
      }

      if (dayKey <= previousComparableYmd) {
        if (Number.isFinite(sales)) {
          previousYtdSales += sales;
        }
        if (Number.isFinite(orders)) {
          previousYtdOrders += orders;
        }
      }
    }
  });

  return {
    source: "shopifyql",
    current_ytd_sales: round(currentYtdSales, 2),
    previous_ytd_sales: round(previousYtdSales, 2),
    current_ytd_orders: round(currentYtdOrders, 2),
    previous_ytd_orders: round(previousYtdOrders, 2),
    previous_full_year_sales: round(previousFullYearSales, 2),
    previous_full_year_orders: round(previousFullYearOrders, 2),
    range: {
      since: toYmd(previousYearStart),
      until: toYmd(tomorrow),
      previous_end_utc: previousComparableEnd.toISOString(),
    },
  };
}

async function fetchShopifyOrderCountBetween(start, end) {
  if (!SHOPIFY_DOMAIN || !SHOPIFY_ACCESS_TOKEN) {
    return null;
  }

  const endpoint = new URL(`https://${SHOPIFY_DOMAIN}/admin/api/${SHOPIFY_ORDERS_API_VERSION}/orders/count.json`);
  endpoint.searchParams.set("status", "any");
  endpoint.searchParams.set("created_at_min", start.toISOString());
  endpoint.searchParams.set("created_at_max", end.toISOString());

  const res = await fetch(endpoint, {
    headers: {
      "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
    },
  });

  if (!res.ok) {
    throw new Error(`Shopify orders count failed (${res.status}): ${(await res.text()).slice(0, 500)}`);
  }

  const payload = await safeParseJson(res);
  const count = Number(payload?.count);
  return Number.isFinite(count) ? count : null;
}

async function getShopifySameTimeComparableSnapshot(now) {
  if (!SHOPIFY_DOMAIN || !SHOPIFY_ACCESS_TOKEN) {
    return null;
  }

  const monthStart = startOfUtcMonth(now);
  const prevMonthStart = addUtcMonths(monthStart, -1);
  const elapsedMs = now.getTime() - monthStart.getTime();
  const prevComparableEnd = new Date(prevMonthStart.getTime() + elapsedMs);
  const tomorrow = addUtcDays(startOfUtcDay(now), 1);
  const queryText =
    `FROM sales SHOW total_sales, net_sales GROUP BY hour ` +
    `SINCE ${toYmd(prevMonthStart)} UNTIL ${toYmd(tomorrow)}`;

  const [table, currentOrderCount, previousOrderCount] = await Promise.all([
    fetchShopifyqlTable(queryText),
    fetchShopifyOrderCountBetween(monthStart, now),
    fetchShopifyOrderCountBetween(prevMonthStart, prevComparableEnd),
  ]);
  if (!table) {
    return null;
  }

  let currentMtdSales = 0;
  let currentMtdNetSales = 0;
  let previousComparableSales = 0;
  let previousComparableNetSales = 0;

  table.rows.forEach((row) => {
    const hourRaw = readShopifyqlCell(row, table.columnIndex, "hour");
    const ts = Date.parse(String(hourRaw || ""));
    const sales = Number(readShopifyqlCell(row, table.columnIndex, "total_sales"));
    const netSales = Number(readShopifyqlCell(row, table.columnIndex, "net_sales"));
    if (!Number.isFinite(ts)) {
      return;
    }

    if (ts >= monthStart.getTime() && ts <= now.getTime() && Number.isFinite(sales)) {
      currentMtdSales += sales;
    }
    if (ts >= prevMonthStart.getTime() && ts <= prevComparableEnd.getTime() && Number.isFinite(sales)) {
      previousComparableSales += sales;
    }
    if (ts >= monthStart.getTime() && ts <= now.getTime() && Number.isFinite(netSales)) {
      currentMtdNetSales += netSales;
    }
    if (ts >= prevMonthStart.getTime() && ts <= prevComparableEnd.getTime() && Number.isFinite(netSales)) {
      previousComparableNetSales += netSales;
    }
  });

  return {
    source: "shopify_same_time",
    current_mtd_sales: round(currentMtdSales, 2),
    current_mtd_net_sales: round(currentMtdNetSales, 2),
    previous_mtd_sales: round(previousComparableSales, 2),
    previous_mtd_net_sales: round(previousComparableNetSales, 2),
    current_mtd_orders: Number.isFinite(currentOrderCount) ? currentOrderCount : null,
    previous_mtd_orders: Number.isFinite(previousOrderCount) ? previousOrderCount : null,
    range: {
      current_start_utc: monthStart.toISOString(),
      current_end_utc: now.toISOString(),
      previous_start_utc: prevMonthStart.toISOString(),
      previous_end_utc: prevComparableEnd.toISOString(),
    },
  };
}

function filterBetween(rows, key, start, end) {
  const startTs = start.getTime();
  const endTs = end.getTime();
  return rows.filter((row) => {
    const ts = Date.parse(row[key] || "");
    return Number.isFinite(ts) && ts >= startTs && ts <= endTs;
  });
}

function isReportableOrderRow(row) {
  const financialStatus = String(row?.financial_status || "").toLowerCase();
  if (financialStatus === "voided") {
    return false;
  }
  return true;
}

function filterReportableOrderRows(rows) {
  return rows.filter((row) => isReportableOrderRow(row));
}

function filterLinesByOrderIdSet(lines, orderIdSet) {
  if (!(orderIdSet instanceof Set) || orderIdSet.size === 0) {
    return [];
  }
  return lines.filter((line) => orderIdSet.has(line.order_id));
}

async function fetchReportableOrderIdSetWithin(start, end) {
  const orderRows = await fetchOrderRowsSinceIso(start.toISOString());
  const scopedRows = filterBetween(orderRows, "created_at_utc", start, end);
  const reportableRows = filterReportableOrderRows(scopedRows);
  return new Set(reportableRows.map((row) => row.order_id));
}

async function fetchHourlyRowsSinceDays(days) {
  const since = new Date(Date.now() - days * 24 * 60 * 60 * 1000).toISOString();
  const filter = `logged_at_utc=gte.${encodeURIComponent(since)}`;
  const rows = await supabaseSelectAll({
    config: supabase,
    table: supabase.table,
    select: HOURLY_SELECT,
    orderBy: "logged_at_utc.asc",
    filter,
  });
  return rows.map(normalizeHourlyRow);
}

async function fetchHourlyRowsSinceIso(sinceIso) {
  const filter = `logged_at_utc=gte.${encodeURIComponent(sinceIso)}`;
  const rows = await supabaseSelectAll({
    config: supabase,
    table: supabase.table,
    select: HOURLY_SELECT,
    orderBy: "logged_at_utc.asc",
    filter,
  });
  return rows.map(normalizeHourlyRow);
}

async function fetchOrderRowsSinceIso(sinceIso) {
  const filter = `created_at_utc=gte.${encodeURIComponent(sinceIso)}`;
  const rows = await supabaseSelectAll({
    config: supabase,
    table: supabase.ordersTable,
    select: ORDERS_SELECT,
    orderBy: "created_at_utc.asc",
    filter,
  });
  return rows.map(normalizeOrderRow);
}

async function fetchLineRowsSinceIso(sinceIso) {
  const filter = `created_at_utc=gte.${encodeURIComponent(sinceIso)}`;
  const rows = await supabaseSelectAll({
    config: supabase,
    table: supabase.orderLinesTable,
    select: LINES_SELECT,
    orderBy: "created_at_utc.asc",
    filter,
  });
  return rows.map(normalizeLineRow);
}

function normalizeHourlyRow(row) {
  return {
    row_key: row.row_key,
    logged_at_utc: row.logged_at_utc,
    logged_at_local: row.logged_at_local,
    sales_amount: toNum(row.sales_amount),
    orders: toNum(row.orders),
    ad_spend: toNum(row.ad_spend),
    roas: toNum(row.roas),
    source_sales: row.source_sales || null,
    source_marketing: row.source_marketing || null,
    ingested_at_utc: row.ingested_at_utc || null,
  };
}

function normalizeOrderRow(row) {
  return {
    order_id: row.order_id,
    order_name: row.order_name,
    created_at_utc: row.created_at_utc,
    processed_at_utc: row.processed_at_utc,
    currency: row.currency,
    source_name: row.source_name || "unknown",
    customer_id: row.customer_id,
    customer_type: row.customer_type || "unknown",
    gross_sales: toNum(row.gross_sales),
    net_sales: toNum(row.net_sales),
    total_sales: toNum(row.total_sales),
    discounts: toNum(row.discounts),
    returns_amount: toNum(row.returns_amount),
    refunds_count: toNum(row.refunds_count),
    line_items_count: toNum(row.line_items_count),
    financial_status: row.financial_status,
    fulfillment_status: row.fulfillment_status,
  };
}

function normalizeLineRow(row) {
  return {
    order_line_key: row.order_line_key,
    order_id: row.order_id,
    line_item_id: row.line_item_id,
    created_at_utc: row.created_at_utc,
    product_id: row.product_id || null,
    variant_id: row.variant_id || null,
    sku: row.sku || null,
    product_title: row.product_title || "Unknown Product",
    variant_title: row.variant_title || null,
    vendor: row.vendor || null,
    source_name: row.source_name || "unknown",
    customer_type: row.customer_type || "unknown",
    quantity: toNum(row.quantity) || 0,
    gross_revenue: toNum(row.gross_revenue) || 0,
    discount_amount: toNum(row.discount_amount) || 0,
    net_revenue: toNum(row.net_revenue) || 0,
    returned_quantity: toNum(row.returned_quantity) || 0,
    returned_revenue: toNum(row.returned_revenue) || 0,
    net_quantity: toNum(row.net_quantity) || 0,
    net_revenue_after_returns: toNum(row.net_revenue_after_returns) || 0,
  };
}

async function buildSummaryPayload() {
  const now = nowUtc();
  const currentMonthStart = startOfUtcMonth(now);
  const prevMonthStart = addUtcMonths(currentMonthStart, -1);
  const prevComparableEnd = previousMtdComparableEnd(now, currentMonthStart);

  const [orderRows, hourlyRows, salesMonthSnapshot, salesComparableSnapshot, sameTimeSnapshot] = await Promise.all([
    fetchOrderRowsSinceIso(prevMonthStart.toISOString()),
    fetchHourlyRowsSinceIso(prevMonthStart.toISOString()),
    getShopifySalesMonthSnapshot(now).catch((error) => {
      console.warn("ShopifyQL month snapshot unavailable in summary:", error.message);
      return null;
    }),
    getShopifySalesComparableSnapshot(now).catch((error) => {
      console.warn("ShopifyQL comparable snapshot unavailable in summary:", error.message);
      return null;
    }),
    getShopifySameTimeComparableSnapshot(now).catch((error) => {
      console.warn("Shopify same-time snapshot unavailable in summary:", error.message);
      return null;
    }),
  ]);

  const currentOrderRows = filterReportableOrderRows(filterBetween(orderRows, "created_at_utc", currentMonthStart, now));
  const prevOrderRows = filterReportableOrderRows(filterBetween(orderRows, "created_at_utc", prevMonthStart, prevComparableEnd));
  const currentHourlyRows = filterBetween(hourlyRows, "logged_at_utc", currentMonthStart, now);
  const prevHourlyRows = filterBetween(hourlyRows, "logged_at_utc", prevMonthStart, prevComparableEnd);

  const currentOrderTotals = aggregateOrders(currentOrderRows);
  const prevOrderTotals = aggregateOrders(prevOrderRows);
  const currentHourlyTotals = aggregateHourly(currentHourlyRows);
  const prevHourlyTotals = aggregateHourly(prevHourlyRows);

  const currentSales = Number.isFinite(salesComparableSnapshot?.current_mtd)
    ? salesComparableSnapshot.current_mtd
    : Number.isFinite(salesMonthSnapshot?.current?.total_sales)
    ? salesMonthSnapshot.current.total_sales
    : currentOrderTotals.total_sales;
  const previousSales = Number.isFinite(salesComparableSnapshot?.previous_mtd)
    ? salesComparableSnapshot.previous_mtd
    : prevOrderTotals.total_sales;
  const currentOrders = Number.isFinite(salesComparableSnapshot?.current_mtd_orders)
    ? salesComparableSnapshot.current_mtd_orders
    : currentOrderTotals.orders_count;
  const previousOrders = Number.isFinite(salesComparableSnapshot?.previous_mtd_orders)
    ? salesComparableSnapshot.previous_mtd_orders
    : prevOrderTotals.orders_count;
  const currentAdSpend = currentHourlyTotals.ad_spend;
  const previousAdSpend = prevHourlyTotals.ad_spend;
  const currentNetSalesForAov = Number.isFinite(salesComparableSnapshot?.current_mtd_net_sales)
    ? salesComparableSnapshot.current_mtd_net_sales
    : Number.isFinite(salesMonthSnapshot?.current?.net_sales)
    ? salesMonthSnapshot.current.net_sales
    : currentOrderTotals.net_sales;
  const previousNetSalesForAov = Number.isFinite(salesComparableSnapshot?.previous_mtd_net_sales)
    ? salesComparableSnapshot.previous_mtd_net_sales
    : prevOrderTotals.net_sales;
  const currentAov = currentOrders > 0 ? currentNetSalesForAov / currentOrders : null;
  const previousAov = previousOrders > 0 ? previousNetSalesForAov / previousOrders : null;
  const currentRoas = currentAdSpend > 0 ? currentSales / currentAdSpend : null;
  const previousRoas = previousAdSpend > 0 ? previousSales / previousAdSpend : null;
  const currentSalesForChange = Number.isFinite(sameTimeSnapshot?.current_mtd_sales) ? sameTimeSnapshot.current_mtd_sales : currentSales;
  const previousSalesForChange = Number.isFinite(sameTimeSnapshot?.previous_mtd_sales)
    ? sameTimeSnapshot.previous_mtd_sales
    : previousSales;
  const currentOrdersForChange = Number.isFinite(sameTimeSnapshot?.current_mtd_orders)
    ? sameTimeSnapshot.current_mtd_orders
    : currentOrders;
  const previousOrdersForChange = Number.isFinite(sameTimeSnapshot?.previous_mtd_orders)
    ? sameTimeSnapshot.previous_mtd_orders
    : previousOrders;
  const currentNetSalesForChange = Number.isFinite(sameTimeSnapshot?.current_mtd_net_sales)
    ? sameTimeSnapshot.current_mtd_net_sales
    : currentNetSalesForAov;
  const previousNetSalesForChange = Number.isFinite(sameTimeSnapshot?.previous_mtd_net_sales)
    ? sameTimeSnapshot.previous_mtd_net_sales
    : previousNetSalesForAov;
  const currentAovForChange = currentOrdersForChange > 0 ? currentNetSalesForChange / currentOrdersForChange : null;
  const previousAovForChange = previousOrdersForChange > 0 ? previousNetSalesForChange / previousOrdersForChange : null;
  const salesPctChange = pctChange(currentSalesForChange, previousSalesForChange);
  const ordersPctChange = pctChange(currentOrdersForChange, previousOrdersForChange);

  const current = {
    sales_amount: round(currentSales, 2),
    orders: round(currentOrders, 2),
    ad_spend: round(currentAdSpend, 2),
    roas: round(currentRoas, 4),
    aov: round(currentAov, 2),
    row_count: currentOrderRows.length,
  };

  const previous = {
    sales_amount: round(previousSales, 2),
    orders: round(previousOrders, 2),
    ad_spend: round(previousAdSpend, 2),
    roas: round(previousRoas, 4),
    aov: round(previousAov, 2),
    row_count: prevOrderRows.length,
  };

  const summary = {
    mtd_sales: current.sales_amount,
    mtd_orders: current.orders,
    mtd_ad_spend: current.ad_spend,
    mtd_roas: current.roas,
    mtd_aov: current.aov,
    sales_source:
      Number.isFinite(salesComparableSnapshot?.current_mtd) || Number.isFinite(salesMonthSnapshot?.current?.total_sales)
        ? "shopifyql"
        : "orders_table",
  };

  const kpis = {
    current,
    previous,
    change: {
      sales_amount_pct: Number.isFinite(salesPctChange) ? round(salesPctChange, 0) : salesPctChange,
      orders_pct: Number.isFinite(ordersPctChange) ? round(ordersPctChange, 0) : ordersPctChange,
      ad_spend_pct: pctChange(current.ad_spend, previous.ad_spend),
      roas_pct: pctChange(current.roas, previous.roas),
      aov_pct: pctChange(currentAov, previousAov),
    },
  };

  return {
    updatedAt: now.toISOString(),
    window: {
      current_start_utc: currentMonthStart.toISOString(),
      current_end_utc: now.toISOString(),
      previous_start_utc: prevMonthStart.toISOString(),
      previous_end_utc: prevComparableEnd.toISOString(),
      reporting_timezone: REPORTING_TIMEZONE,
    },
    summary,
    kpis,
  };
}

async function getYtdComparisonPayload() {
  const now = nowUtc();
  const currentYearStart = startOfUtcYear(now);
  const previousYearStart = addUtcYears(currentYearStart, -1);
  const previousComparableEnd = previousYtdComparableEnd(now);
  const previousYearEnd = new Date(currentYearStart.getTime() - 1);

  let shopifyYtdSnapshot = null;
  try {
    shopifyYtdSnapshot = await getShopifySalesYtdComparableSnapshot(now);
  } catch (error) {
    console.warn("ShopifyQL YTD snapshot unavailable:", error.message);
  }

  const orderRows = await fetchOrderRowsSinceIso(previousYearStart.toISOString());
  const currentRows = filterReportableOrderRows(filterBetween(orderRows, "created_at_utc", currentYearStart, now));
  const previousRows = filterReportableOrderRows(
    filterBetween(orderRows, "created_at_utc", previousYearStart, previousComparableEnd)
  );
  const previousFullYearRows = filterReportableOrderRows(
    filterBetween(orderRows, "created_at_utc", previousYearStart, previousYearEnd)
  );

  const currentTotals = aggregateOrders(currentRows);
  const previousTotals = aggregateOrders(previousRows);
  const previousFullYearTotals = aggregateOrders(previousFullYearRows);

  const currentSales = Number.isFinite(shopifyYtdSnapshot?.current_ytd_sales)
    ? shopifyYtdSnapshot.current_ytd_sales
    : currentRows.length
      ? currentTotals.total_sales
      : null;
  const previousSales = Number.isFinite(shopifyYtdSnapshot?.previous_ytd_sales)
    ? shopifyYtdSnapshot.previous_ytd_sales
    : previousRows.length
      ? previousTotals.total_sales
      : null;
  const currentOrders = Number.isFinite(shopifyYtdSnapshot?.current_ytd_orders)
    ? shopifyYtdSnapshot.current_ytd_orders
    : currentRows.length
      ? currentTotals.orders_count
      : null;
  const previousOrders = Number.isFinite(shopifyYtdSnapshot?.previous_ytd_orders)
    ? shopifyYtdSnapshot.previous_ytd_orders
    : previousRows.length
      ? previousTotals.orders_count
      : null;
  const previousYearTotalSales = Number.isFinite(shopifyYtdSnapshot?.previous_full_year_sales)
    ? shopifyYtdSnapshot.previous_full_year_sales
    : previousFullYearRows.length
      ? previousFullYearTotals.total_sales
      : null;
  const previousYearTotalOrders = Number.isFinite(shopifyYtdSnapshot?.previous_full_year_orders)
    ? shopifyYtdSnapshot.previous_full_year_orders
    : previousFullYearRows.length
      ? previousFullYearTotals.orders_count
      : null;
  const salesPct = pctChange(currentSales, previousSales);
  const ordersPct = pctChange(currentOrders, previousOrders);

  return {
    updatedAt: now.toISOString(),
    period: {
      current_start_utc: currentYearStart.toISOString(),
      current_end_utc: now.toISOString(),
      previous_start_utc: previousYearStart.toISOString(),
      previous_end_utc: previousComparableEnd.toISOString(),
      reporting_timezone: REPORTING_TIMEZONE,
      comparison_basis: "same_local_datetime_previous_year",
    },
    current: {
      sales_amount: round(currentSales, 2),
      orders: round(currentOrders, 0),
      row_count: currentRows.length,
    },
    previous: {
      sales_amount: round(previousSales, 2),
      orders: round(previousOrders, 0),
      row_count: previousRows.length,
    },
    previous_year: {
      sales_amount: round(previousYearTotalSales, 2),
      orders: round(previousYearTotalOrders, 0),
      row_count: previousFullYearRows.length,
      start_utc: previousYearStart.toISOString(),
      end_utc: previousYearEnd.toISOString(),
    },
    change: {
      sales_amount_pct: salesPct,
      orders_pct: ordersPct,
      growth_rate_pct: salesPct,
    },
    source: {
      sales: Number.isFinite(shopifyYtdSnapshot?.current_ytd_sales) ? "shopifyql" : currentRows.length ? "orders_table" : "unavailable",
      orders:
        Number.isFinite(shopifyYtdSnapshot?.current_ytd_orders) ? "shopifyql" : currentRows.length ? "orders_table" : "unavailable",
    },
  };
}

function aggregateHourly(rows) {
  const totals = rows.reduce(
    (acc, row) => {
      acc.sales_amount += row.sales_amount || 0;
      acc.orders += row.orders || 0;
      acc.ad_spend += row.ad_spend || 0;
      return acc;
    },
    { sales_amount: 0, orders: 0, ad_spend: 0 }
  );

  const roas = totals.ad_spend > 0 ? totals.sales_amount / totals.ad_spend : null;
  const aov = totals.orders > 0 ? totals.sales_amount / totals.orders : null;

  return {
    sales_amount: round(totals.sales_amount, 2),
    orders: round(totals.orders, 2),
    ad_spend: round(totals.ad_spend, 2),
    roas: round(roas, 4),
    aov: round(aov, 2),
    row_count: rows.length,
  };
}

function aggregateOrders(rows) {
  const totals = rows.reduce(
    (acc, row) => {
      if (!isReportableOrderRow(row)) {
        return acc;
      }
      acc.gross_sales += row.gross_sales || 0;
      acc.net_sales += row.net_sales || 0;
      acc.total_sales += row.total_sales || 0;
      acc.discounts += row.discounts || 0;
      acc.returns_amount += row.returns_amount || 0;
      acc.orders_count += 1;
      return acc;
    },
    { gross_sales: 0, net_sales: 0, total_sales: 0, discounts: 0, returns_amount: 0, orders_count: 0 }
  );
  return {
    gross_sales: round(totals.gross_sales, 2),
    net_sales: round(totals.net_sales, 2),
    total_sales: round(totals.total_sales, 2),
    discounts: round(totals.discounts, 2),
    returns_amount: round(totals.returns_amount, 2),
    orders_count: totals.orders_count,
  };
}

function getProductKey(line) {
  if (line.product_id) {
    return line.product_id;
  }
  return `title:${line.product_title}`;
}

function aggregateProducts(lines) {
  const map = {};
  lines.forEach((line) => {
    const key = getProductKey(line);
    if (!map[key]) {
      map[key] = {
        product_key: key,
        product_id: line.product_id,
        title: line.product_title || "Unknown Product",
        units: 0,
        revenue: 0,
        gross_revenue: 0,
        returned_units: 0,
        returned_revenue: 0,
      };
    }
    map[key].units += line.net_quantity || 0;
    map[key].revenue += line.net_revenue_after_returns || 0;
    map[key].gross_revenue += line.gross_revenue || 0;
    map[key].returned_units += line.returned_quantity || 0;
    map[key].returned_revenue += line.returned_revenue || 0;
  });
  return Object.values(map);
}

async function getTopProductsByUnits(limit) {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const [lines, includedOrderIds] = await Promise.all([
    fetchLineRowsSinceIso(monthStart.toISOString()),
    fetchReportableOrderIdSetWithin(monthStart, now),
  ]);
  const scopedLines = filterLinesByOrderIdSet(filterBetween(lines, "created_at_utc", monthStart, now), includedOrderIds);
  const products = aggregateProducts(scopedLines);
  const totalUnits = products.reduce((sum, p) => sum + p.units, 0);

  const top = products
    .sort((a, b) => b.units - a.units)
    .slice(0, limit)
    .map((p, i) => ({
      rank: i + 1,
      product_key: p.product_key,
      product_id: p.product_id,
      title: p.title,
      units: round(p.units, 2),
      revenue: round(p.revenue, 2),
      unit_share_pct: totalUnits > 0 ? round((p.units / totalUnits) * 100, 2) : null,
    }));

  return {
    updatedAt: new Date().toISOString(),
    period: {
      start_utc: monthStart.toISOString(),
      end_utc: now.toISOString(),
    },
    total_units: round(totalUnits, 2),
    products: top,
  };
}

async function getTopProductsByRevenue(limit) {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const [lines, includedOrderIds] = await Promise.all([
    fetchLineRowsSinceIso(monthStart.toISOString()),
    fetchReportableOrderIdSetWithin(monthStart, now),
  ]);
  const scopedLines = filterLinesByOrderIdSet(filterBetween(lines, "created_at_utc", monthStart, now), includedOrderIds);
  const products = aggregateProducts(scopedLines);
  const totalRevenue = products.reduce((sum, p) => sum + p.revenue, 0);

  const top = products
    .sort((a, b) => b.revenue - a.revenue)
    .slice(0, limit)
    .map((p, i) => ({
      rank: i + 1,
      product_key: p.product_key,
      product_id: p.product_id,
      title: p.title,
      revenue: round(p.revenue, 2),
      units: round(p.units, 2),
      revenue_share_pct: totalRevenue > 0 ? round((p.revenue / totalRevenue) * 100, 2) : null,
    }));

  return {
    updatedAt: new Date().toISOString(),
    period: {
      start_utc: monthStart.toISOString(),
      end_utc: now.toISOString(),
    },
    total_revenue: round(totalRevenue, 2),
    products: top,
  };
}

async function getProductMomentum({ limit, metric }) {
  const now = nowUtc();
  const todayStart = startOfUtcDay(now);
  const thisWeekStart = addUtcDays(todayStart, -6);
  const prevWeekStart = addUtcDays(thisWeekStart, -7);
  const prevWeekEnd = new Date(thisWeekStart.getTime() - 1000);

  const [lines, includedOrderIds] = await Promise.all([
    fetchLineRowsSinceIso(prevWeekStart.toISOString()),
    fetchReportableOrderIdSetWithin(prevWeekStart, now),
  ]);
  const scopedLines = filterLinesByOrderIdSet(filterBetween(lines, "created_at_utc", prevWeekStart, now), includedOrderIds);
  const thisWeekLines = filterBetween(scopedLines, "created_at_utc", thisWeekStart, now);
  const prevWeekLines = filterBetween(scopedLines, "created_at_utc", prevWeekStart, prevWeekEnd);

  const thisMap = aggregateProducts(thisWeekLines);
  const prevMap = aggregateProducts(prevWeekLines);
  const prevByKey = {};
  prevMap.forEach((p) => {
    prevByKey[p.product_key] = p;
  });

  const merged = thisMap.map((curr) => {
    const prev = prevByKey[curr.product_key] || { units: 0, revenue: 0, title: curr.title };
    const currValue = metric === "units" ? curr.units : curr.revenue;
    const prevValue = metric === "units" ? prev.units : prev.revenue;
    const delta = currValue - prevValue;
    return {
      product_key: curr.product_key,
      title: curr.title,
      metric,
      current_value: round(currValue, 2),
      previous_value: round(prevValue, 2),
      delta: round(delta, 2),
      delta_pct: pctChange(currValue, prevValue),
    };
  });

  const products = merged.sort((a, b) => (b.delta || 0) - (a.delta || 0)).slice(0, limit);
  return {
    updatedAt: new Date().toISOString(),
    metric,
    windows: {
      this_week_start_utc: thisWeekStart.toISOString(),
      this_week_end_utc: now.toISOString(),
      prev_week_start_utc: prevWeekStart.toISOString(),
      prev_week_end_utc: prevWeekEnd.toISOString(),
    },
    products,
  };
}

async function getDailySalesPace() {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const prevMonthStart = addUtcMonths(monthStart, -1);
  const prevMonthEnd = new Date(monthStart.getTime() - 1000);

  const [monthlyOrdersRaw, prevMonthOrdersRaw, salesMonthSnapshot] = await Promise.all([
    fetchOrderRowsSinceIso(monthStart.toISOString()),
    fetchOrderRowsSinceIso(prevMonthStart.toISOString()),
    getShopifySalesMonthSnapshot(now).catch((error) => {
      console.warn("ShopifyQL month snapshot unavailable in pace:", error.message);
      return null;
    }),
  ]);

  const monthlyOrders = filterBetween(monthlyOrdersRaw, "created_at_utc", monthStart, now);
  const prevOnly = filterBetween(prevMonthOrdersRaw, "created_at_utc", prevMonthStart, prevMonthEnd);
  const prevOrderTotals = aggregateOrders(prevOnly);
  const mtdOrderTotals = aggregateOrders(monthlyOrders);

  const previousMonthTotalSales = Number.isFinite(salesMonthSnapshot?.previous?.total_sales)
    ? salesMonthSnapshot.previous.total_sales
    : prevOrderTotals.total_sales;
  const previousMonthGrossSales = Number.isFinite(salesMonthSnapshot?.previous?.gross_sales)
    ? salesMonthSnapshot.previous.gross_sales
    : prevOrderTotals.gross_sales;
  const previousMonthNetSales = Number.isFinite(salesMonthSnapshot?.previous?.net_sales)
    ? salesMonthSnapshot.previous.net_sales
    : prevOrderTotals.net_sales;
  const mtdTotalSales = Number.isFinite(salesMonthSnapshot?.current?.total_sales)
    ? salesMonthSnapshot.current.total_sales
    : mtdOrderTotals.total_sales;
  const mtdGrossSales = Number.isFinite(salesMonthSnapshot?.current?.gross_sales)
    ? salesMonthSnapshot.current.gross_sales
    : mtdOrderTotals.gross_sales;
  const mtdNetSales = Number.isFinite(salesMonthSnapshot?.current?.net_sales)
    ? salesMonthSnapshot.current.net_sales
    : mtdOrderTotals.net_sales;

  const targetOverride = toNum(process.env.MONTHLY_SALES_TARGET);
  const monthlyTarget = Number.isFinite(targetOverride) ? targetOverride : previousMonthTotalSales || null;

  const todayStart = startOfUtcDay(now);
  const todayRows = filterBetween(monthlyOrders, "created_at_utc", todayStart, now);
  const todaySales = aggregateOrders(todayRows).total_sales || 0;

  const dayOfMonth = dayOfMonthInReportingZone(now);
  const totalDays = daysInUtcMonth(now);
  const daysRemaining = Math.max(0, totalDays - dayOfMonth);
  const requiredDailyPace =
    Number.isFinite(monthlyTarget) && monthlyTarget > 0
      ? Math.max(0, (monthlyTarget - (mtdTotalSales || 0)) / Math.max(1, daysRemaining))
      : null;

  return {
    updatedAt: now.toISOString(),
    target_source: Number.isFinite(targetOverride) ? "env_monthly_sales_target" : "previous_month_total_sales",
    sales_source: Number.isFinite(salesMonthSnapshot?.current?.total_sales) ? "shopifyql" : "orders_table",
    month_goal: Number.isFinite(monthlyTarget) ? round(monthlyTarget, 2) : null,
    mtd_sales: round(mtdTotalSales, 2),
    mtd_gross_sales: round(mtdGrossSales, 2),
    mtd_net_sales: round(mtdNetSales, 2),
    today_sales: round(todaySales, 2),
    required_daily_pace: round(requiredDailyPace, 2),
    on_track_today: Number.isFinite(requiredDailyPace) ? todaySales >= requiredDailyPace : null,
    previous_month_gross_sales: round(previousMonthGrossSales, 2),
    previous_month_total_sales: round(previousMonthTotalSales, 2),
    previous_month_net_sales: round(previousMonthNetSales, 2),
    previous_month_orders: prevOrderTotals.orders_count,
    days_elapsed: dayOfMonth,
    days_remaining: daysRemaining,
  };
}

async function getMtdProjection() {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const [monthlyOrdersRaw, salesMonthSnapshot] = await Promise.all([
    fetchOrderRowsSinceIso(monthStart.toISOString()),
    getShopifySalesMonthSnapshot(now).catch((error) => {
      console.warn("ShopifyQL month snapshot unavailable in projection:", error.message);
      return null;
    }),
  ]);
  const monthlyOrders = filterBetween(monthlyOrdersRaw, "created_at_utc", monthStart, now);
  const mtdOrderTotals = aggregateOrders(monthlyOrders);
  const mtdSales = Number.isFinite(salesMonthSnapshot?.current?.total_sales)
    ? salesMonthSnapshot.current.total_sales
    : mtdOrderTotals.total_sales;
  const mtdGrossSales = Number.isFinite(salesMonthSnapshot?.current?.gross_sales)
    ? salesMonthSnapshot.current.gross_sales
    : mtdOrderTotals.gross_sales;
  const mtdNetSales = Number.isFinite(salesMonthSnapshot?.current?.net_sales)
    ? salesMonthSnapshot.current.net_sales
    : mtdOrderTotals.net_sales;

  const dayOfMonth = dayOfMonthInReportingZone(now);
  const totalDays = daysInUtcMonth(now);
  const runRatePerDay = dayOfMonth > 0 ? (mtdSales || 0) / dayOfMonth : null;
  const projectedMonthEnd = Number.isFinite(runRatePerDay) ? runRatePerDay * totalDays : null;

  const targetOverride = toNum(process.env.MONTHLY_SALES_TARGET);
  const monthGoal = Number.isFinite(targetOverride) ? targetOverride : null;

  return {
    updatedAt: now.toISOString(),
    sales_source: Number.isFinite(salesMonthSnapshot?.current?.total_sales) ? "shopifyql" : "orders_table",
    mtd_sales: round(mtdSales, 2),
    mtd_gross_sales: round(mtdGrossSales, 2),
    mtd_net_sales: round(mtdNetSales, 2),
    progress_pct_of_target: Number.isFinite(monthGoal) && monthGoal > 0 ? round(((mtdSales || 0) / monthGoal) * 100, 2) : null,
    projected_month_end_sales: round(projectedMonthEnd, 2),
    projected_vs_target_pct:
      Number.isFinite(monthGoal) && monthGoal > 0 && Number.isFinite(projectedMonthEnd)
        ? round(((projectedMonthEnd - monthGoal) / monthGoal) * 100, 2)
        : null,
    run_rate_daily_sales: round(runRatePerDay, 2),
    month_goal: Number.isFinite(monthGoal) ? round(monthGoal, 2) : null,
    days_elapsed: dayOfMonth,
    days_in_month: totalDays,
  };
}

async function getGrossNetReturns() {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const ordersRaw = await fetchOrderRowsSinceIso(monthStart.toISOString());
  const orders = filterBetween(ordersRaw, "created_at_utc", monthStart, now);
  const totals = aggregateOrders(orders);
  const returnsRateOnNet = totals.net_sales > 0 ? (totals.returns_amount / totals.net_sales) * 100 : null;

  return {
    updatedAt: now.toISOString(),
    period: {
      start_utc: monthStart.toISOString(),
      end_utc: now.toISOString(),
    },
    gross_sales: totals.gross_sales,
    net_sales: totals.net_sales,
    total_sales: totals.total_sales,
    returns_amount: totals.returns_amount,
    returns_rate_pct_of_net: round(returnsRateOnNet, 2),
    orders_count: totals.orders_count,
  };
}

async function getAovPayload() {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const prevMonthStart = addUtcMonths(monthStart, -1);
  const prevComparableEnd = previousMtdComparableEnd(now, monthStart);

  let ordersFetchError = "";
  let shopifyComparableError = "";
  const [orders, salesComparableSnapshot] = await Promise.all([
    fetchOrderRowsSinceIso(prevMonthStart.toISOString()).catch((error) => {
      ordersFetchError = errorMessage(error, "Orders table fetch failed");
      console.warn("Orders fallback unavailable in aov:", ordersFetchError);
      return [];
    }),
    getShopifySalesComparableSnapshot(now).catch((error) => {
      shopifyComparableError = errorMessage(error, "ShopifyQL comparable snapshot unavailable");
      console.warn("ShopifyQL comparable snapshot unavailable in aov:", shopifyComparableError);
      return null;
    }),
  ]);

  const safeOrders = Array.isArray(orders) ? orders : [];
  const currentOrders = filterBetween(safeOrders, "created_at_utc", monthStart, now);
  const prevOrders = filterBetween(safeOrders, "created_at_utc", prevMonthStart, prevComparableEnd);
  const current = aggregateOrders(currentOrders);
  const previous = aggregateOrders(prevOrders);

  const hasShopifySales =
    Number.isFinite(salesComparableSnapshot?.current_mtd_net_sales) &&
    Number.isFinite(salesComparableSnapshot?.previous_mtd_net_sales);
  const hasShopifyOrders =
    Number.isFinite(salesComparableSnapshot?.current_mtd_orders) &&
    Number.isFinite(salesComparableSnapshot?.previous_mtd_orders);

  const currentNetSales = hasShopifySales ? salesComparableSnapshot.current_mtd_net_sales : current.net_sales;
  const previousNetSales = hasShopifySales ? salesComparableSnapshot.previous_mtd_net_sales : previous.net_sales;
  const currentOrderCount = hasShopifyOrders ? salesComparableSnapshot.current_mtd_orders : current.orders_count;
  const previousOrderCount = hasShopifyOrders ? salesComparableSnapshot.previous_mtd_orders : previous.orders_count;

  const currentAov = currentOrderCount > 0 ? currentNetSales / currentOrderCount : null;
  const previousAov = previousOrderCount > 0 ? previousNetSales / previousOrderCount : null;
  const hasAov = Number.isFinite(currentAov) && Number.isFinite(previousAov);
  const unavailableReason = hasAov
    ? ""
    : shopifyComparableError
      ? shopifyComparableError
      : ordersFetchError
        ? ordersFetchError
        : "AOV data unavailable";

  return {
    updatedAt: now.toISOString(),
    status: hasAov ? "ok" : "unavailable",
    source_sales: hasShopifySales ? "shopifyql" : "orders_table",
    source_orders: hasShopifyOrders ? "shopifyql" : "orders_table",
    unavailable_reason: unavailableReason,
    period: {
      current_start_utc: monthStart.toISOString(),
      current_end_utc: now.toISOString(),
      previous_start_utc: prevMonthStart.toISOString(),
      previous_end_utc: prevComparableEnd.toISOString(),
    },
    mtd_aov: round(currentAov, 2),
    previous_period_aov: round(previousAov, 2),
    aov_change_pct: pctChange(currentAov, previousAov),
    mtd_orders: round(currentOrderCount, 2),
    mtd_net_sales: round(currentNetSales, 2),
    previous_mtd_net_sales: round(previousNetSales, 2),
    mtd_sales: round(currentNetSales, 2),
  };
}

async function getWebsiteSessionsMtdPayload() {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const prevMonthStart = addUtcMonths(monthStart, -1);
  const prevComparableEnd = previousMtdComparableEnd(now, monthStart);

  let sessionsSnapshot = null;
  let unavailableReason = "";
  try {
    sessionsSnapshot = await getShopifySessionsComparableSnapshot(now);
  } catch (error) {
    unavailableReason = errorMessage(error, "Shopify sessions snapshot unavailable");
    console.warn("ShopifyQL sessions snapshot unavailable:", unavailableReason);
  }

  const mtdSessions = Number.isFinite(sessionsSnapshot?.current_mtd) ? sessionsSnapshot.current_mtd : null;
  const previousMtdSessions = Number.isFinite(sessionsSnapshot?.previous_mtd) ? sessionsSnapshot.previous_mtd : null;

  if (!Number.isFinite(mtdSessions) || !Number.isFinite(previousMtdSessions)) {
    try {
      const scopes = await fetchShopifyAccessScopes();
      if (Array.isArray(scopes) && scopes.length && !scopes.includes("read_reports")) {
        unavailableReason = unavailableReason || "Missing Shopify app scope: read_reports";
      }
    } catch (error) {
      console.warn("Shopify access scopes check failed for sessions:", errorMessage(error));
    }
  }

  const sessionsDelta =
    Number.isFinite(mtdSessions) && Number.isFinite(previousMtdSessions) ? mtdSessions - previousMtdSessions : null;

  return {
    updatedAt: now.toISOString(),
    status: Number.isFinite(mtdSessions) && Number.isFinite(previousMtdSessions) ? "ok" : "unavailable",
    source: sessionsSnapshot?.source || "shopifyql",
    metric: sessionsSnapshot?.metric || null,
    query_used: sessionsSnapshot?.query_used || "",
    unavailable_reason: unavailableReason || "",
    period: {
      current_start_utc: monthStart.toISOString(),
      current_end_utc: now.toISOString(),
      previous_start_utc: prevMonthStart.toISOString(),
      previous_end_utc: prevComparableEnd.toISOString(),
    },
    mtd_sessions: round(mtdSessions, 0),
    previous_mtd_sessions: round(previousMtdSessions, 0),
    sessions_change: round(sessionsDelta, 0),
    sessions_change_pct: pctChange(mtdSessions, previousMtdSessions),
  };
}

async function getShopifySessionsComparableSnapshot(now) {
  if (!SHOPIFY_DOMAIN || !SHOPIFY_ACCESS_TOKEN) {
    return null;
  }

  const monthStart = startOfUtcMonth(now);
  const prevMonthStart = addUtcMonths(monthStart, -1);
  const prevComparableEnd = previousMtdComparableEnd(now, monthStart);
  const tomorrow = addUtcDays(startOfUtcDay(now), 1);
  const monthStartYmd = toYmd(monthStart);

  const compareCandidates = [
    {
      metric: "sessions",
      query: `FROM sales, sessions SHOW sessions SINCE ${monthStartYmd} UNTIL today COMPARE TO previous_period`,
    },
    {
      metric: "online_store_visitors",
      query: `FROM sales, sessions SHOW online_store_visitors SINCE ${monthStartYmd} UNTIL today COMPARE TO previous_period`,
    },
    {
      metric: "online_store_sessions",
      query: `FROM sales, sessions SHOW online_store_sessions SINCE ${monthStartYmd} UNTIL today COMPARE TO previous_period`,
    },
  ];

  for (const candidate of compareCandidates) {
    try {
      const table = await fetchShopifyqlTable(candidate.query);
      if (!table || !Array.isArray(table.rows) || !table.rows.length) {
        continue;
      }

      const metricKey = pickShopifyqlMetricColumn(table.columnIndex, [
        candidate.metric,
        "sessions",
        "online_store_visitors",
        "online_store_sessions",
      ]);
      if (!metricKey) {
        continue;
      }

      const comparisonKey = Object.keys(table.columnIndex).find(
        (key) => key.startsWith(`comparison_${metricKey}`) && key.includes("previous_period")
      );
      if (!comparisonKey) {
        continue;
      }

      const row = table.rows[0];
      const currentMtd = Number(readShopifyqlCell(row, table.columnIndex, metricKey));
      const previousMtd = Number(readShopifyqlCell(row, table.columnIndex, comparisonKey));
      if (!Number.isFinite(currentMtd) || !Number.isFinite(previousMtd)) {
        continue;
      }

      return {
        source: "shopifyql",
        metric: metricKey,
        query_used: candidate.query,
        current_mtd: round(currentMtd, 2),
        previous_mtd: round(previousMtd, 2),
        range: {
          since: monthStartYmd,
          until: "today",
          previous_end_utc: prevComparableEnd.toISOString(),
        },
      };
    } catch (_error) {
      // Fall through to next candidate and then to daily fallback below.
    }
  }

  const since = toYmd(prevMonthStart);
  const until = toYmd(tomorrow);
  const queryCandidates = [
    `FROM sales, sessions SHOW day, sessions GROUP BY day SINCE ${since} UNTIL ${until} ORDER BY day`,
    `FROM sales, sessions SHOW day, online_store_visitors GROUP BY day SINCE ${since} UNTIL ${until} ORDER BY day`,
    `FROM sales, sessions SHOW day, online_store_sessions GROUP BY day SINCE ${since} UNTIL ${until} ORDER BY day`,
    `FROM sales, sessions SHOW day, visitors GROUP BY day SINCE ${since} UNTIL ${until} ORDER BY day`,
    `FROM sales, sessions SHOW sessions SINCE ${since} UNTIL ${until} TIMESERIES day ORDER BY day`,
    `FROM sales, sessions SHOW online_store_visitors SINCE ${since} UNTIL ${until} TIMESERIES day ORDER BY day`,
    `FROM sales, sessions SHOW online_store_sessions SINCE ${since} UNTIL ${until} TIMESERIES day ORDER BY day`,
    `FROM sales, sessions SHOW visitors SINCE ${since} UNTIL ${until} TIMESERIES day ORDER BY day`,
    `FROM sessions SHOW sessions TIMESERIES day SINCE ${since} UNTIL ${until}`,
    `FROM sessions SHOW online_store_sessions TIMESERIES day SINCE ${since} UNTIL ${until}`,
    `FROM sessions SHOW online_store_visitors TIMESERIES day SINCE ${since} UNTIL ${until}`,
    `FROM visits SHOW sessions TIMESERIES day SINCE ${since} UNTIL ${until}`,
    `FROM visits SHOW online_store_sessions TIMESERIES day SINCE ${since} UNTIL ${until}`,
    `FROM visits SHOW online_store_visitors TIMESERIES day SINCE ${since} UNTIL ${until}`,
    `FROM sessions SHOW sessions GROUP BY day SINCE ${since} UNTIL ${until}`,
    `FROM sessions SHOW online_store_sessions GROUP BY day SINCE ${since} UNTIL ${until}`,
    `FROM sessions SHOW online_store_visitors GROUP BY day SINCE ${since} UNTIL ${until}`,
    `FROM sales, sessions SHOW sessions TIMESERIES day SINCE ${since} UNTIL ${until}`,
  ];

  let table = null;
  let metricKey = "";
  let queryUsed = "";
  let lastError = null;

  for (const queryText of queryCandidates) {
    try {
      const candidate = await fetchShopifyqlTable(queryText);
      if (!candidate) {
        continue;
      }
      const candidateMetric = pickShopifyqlMetricColumn(candidate.columnIndex, [
        "sessions",
        "online_store_visitors",
        "online_store_sessions",
        "visitors",
      ]);
      const dayColumn = pickShopifyqlMetricColumn(candidate.columnIndex, ["day", "date", "time"]);
      if (!candidateMetric || !dayColumn) {
        continue;
      }
      table = candidate;
      metricKey = candidateMetric;
      queryUsed = queryText;
      break;
    } catch (error) {
      lastError = error;
    }
  }

  if (!table || !metricKey) {
    if (lastError) {
      throw lastError;
    }
    return null;
  }

  const dayColumn = pickShopifyqlMetricColumn(table.columnIndex, ["day", "date", "time"]);
  if (!dayColumn) {
    return null;
  }

  const currentMonthPrefix = toYmd(monthStart).slice(0, 7);
  const prevMonthPrefix = toYmd(prevMonthStart).slice(0, 7);
  const prevComparableDay = dayOfMonthInReportingZone(prevComparableEnd);
  let currentMtd = 0;
  let previousComparable = 0;

  table.rows.forEach((row) => {
    const dayRaw = readShopifyqlCell(row, table.columnIndex, dayColumn);
    const dayKey = String(dayRaw || "").slice(0, 10);
    const sessions = Number(readShopifyqlCell(row, table.columnIndex, metricKey));
    if (!dayKey || !Number.isFinite(sessions)) {
      return;
    }

    if (dayKey.slice(0, 7) === currentMonthPrefix) {
      currentMtd += sessions;
      return;
    }

    if (dayKey.slice(0, 7) === prevMonthPrefix) {
      const dayOfMonth = Number(dayKey.slice(8, 10));
      if (Number.isFinite(dayOfMonth) && dayOfMonth <= prevComparableDay) {
        previousComparable += sessions;
      }
    }
  });

  return {
    source: "shopifyql",
    metric: metricKey,
    query_used: queryUsed,
    current_mtd: round(currentMtd, 2),
    previous_mtd: round(previousComparable, 2),
    range: {
      since,
      until,
      previous_end_utc: prevComparableEnd.toISOString(),
    },
  };
}

async function fetchShopifyAccessScopes() {
  if (!SHOPIFY_DOMAIN || !SHOPIFY_ACCESS_TOKEN) {
    return [];
  }

  if (
    shopifyAccessScopesCache &&
    Array.isArray(shopifyAccessScopesCache.scopes) &&
    Date.now() - shopifyAccessScopesCache.fetchedAt <= SHOPIFY_ACCESS_SCOPES_CACHE_TTL_MS
  ) {
    return shopifyAccessScopesCache.scopes;
  }

  const endpoint = new URL(`https://${SHOPIFY_DOMAIN}/admin/api/${SHOPIFY_ORDERS_API_VERSION}/access_scopes.json`);
  const res = await fetch(endpoint, {
    headers: {
      "X-Shopify-Access-Token": SHOPIFY_ACCESS_TOKEN,
    },
  });

  if (!res.ok) {
    throw new Error(`Shopify access scopes failed (${res.status}): ${(await res.text()).slice(0, 500)}`);
  }

  const payload = await safeParseJson(res);
  const scopes = Array.isArray(payload?.access_scopes)
    ? payload.access_scopes
        .map((item) => String(item?.handle || "").trim())
        .filter(Boolean)
    : [];

  shopifyAccessScopesCache = {
    fetchedAt: Date.now(),
    scopes,
  };
  return scopes;
}

async function getNewVsReturningPayload() {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const ordersRaw = await fetchOrderRowsSinceIso(monthStart.toISOString());
  const orders = filterReportableOrderRows(filterBetween(ordersRaw, "created_at_utc", monthStart, now));

  const totals = {
    new_revenue: 0,
    returning_revenue: 0,
    unknown_revenue: 0,
    new_orders: 0,
    returning_orders: 0,
    unknown_orders: 0,
  };

  orders.forEach((row) => {
    const revenue = row.total_sales || 0;
    if (row.customer_type === "new") {
      totals.new_revenue += revenue;
      totals.new_orders += 1;
      return;
    }
    if (row.customer_type === "returning") {
      totals.returning_revenue += revenue;
      totals.returning_orders += 1;
      return;
    }
    totals.unknown_revenue += revenue;
    totals.unknown_orders += 1;
  });

  const totalRevenue = totals.new_revenue + totals.returning_revenue + totals.unknown_revenue;
  return {
    updatedAt: now.toISOString(),
    period: {
      start_utc: monthStart.toISOString(),
      end_utc: now.toISOString(),
    },
    revenue: {
      new: round(totals.new_revenue, 2),
      returning: round(totals.returning_revenue, 2),
      unknown: round(totals.unknown_revenue, 2),
    },
    orders: {
      new: totals.new_orders,
      returning: totals.returning_orders,
      unknown: totals.unknown_orders,
    },
    shares_pct: {
      new: totalRevenue > 0 ? round((totals.new_revenue / totalRevenue) * 100, 2) : null,
      returning: totalRevenue > 0 ? round((totals.returning_revenue / totalRevenue) * 100, 2) : null,
      unknown: totalRevenue > 0 ? round((totals.unknown_revenue / totalRevenue) * 100, 2) : null,
    },
  };
}

async function getChannelSplitPayload() {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const ordersRaw = await fetchOrderRowsSinceIso(monthStart.toISOString());
  const orders = filterReportableOrderRows(filterBetween(ordersRaw, "created_at_utc", monthStart, now));

  const byChannel = {};
  orders.forEach((row) => {
    const channel = row.source_name || "unknown";
    if (!byChannel[channel]) {
      byChannel[channel] = { channel, revenue: 0, orders: 0 };
    }
    byChannel[channel].revenue += row.total_sales || 0;
    byChannel[channel].orders += 1;
  });

  const totalRevenue = Object.values(byChannel).reduce((sum, c) => sum + c.revenue, 0);
  const channels = Object.values(byChannel)
    .sort((a, b) => b.revenue - a.revenue)
    .map((c) => ({
      channel: c.channel,
      revenue: round(c.revenue, 2),
      orders: c.orders,
      revenue_share_pct: totalRevenue > 0 ? round((c.revenue / totalRevenue) * 100, 2) : null,
    }));

  return {
    updatedAt: now.toISOString(),
    period: {
      start_utc: monthStart.toISOString(),
      end_utc: now.toISOString(),
    },
    total_revenue: round(totalRevenue, 2),
    channels,
  };
}

async function getDiscountImpactPayload() {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const ordersRaw = await fetchOrderRowsSinceIso(monthStart.toISOString());
  const orders = filterReportableOrderRows(filterBetween(ordersRaw, "created_at_utc", monthStart, now));
  const totals = aggregateOrders(orders);

  const discountedOrders = orders.filter((o) => (o.discounts || 0) > 0).length;
  const discountedOrdersPct = totals.orders_count > 0 ? (discountedOrders / totals.orders_count) * 100 : null;
  const discountRate = totals.gross_sales > 0 ? (totals.discounts / totals.gross_sales) * 100 : null;

  return {
    updatedAt: now.toISOString(),
    period: {
      start_utc: monthStart.toISOString(),
      end_utc: now.toISOString(),
    },
    discounted_orders_count: discountedOrders,
    discounted_orders_pct: round(discountedOrdersPct, 2),
    total_discounts: totals.discounts,
    avg_discount_per_order: totals.orders_count > 0 ? round(totals.discounts / totals.orders_count, 2) : null,
    discount_rate_pct_of_gross: round(discountRate, 2),
  };
}

async function getHourlyHeatmapToday() {
  const now = nowUtc();
  const todayStart = startOfUtcDay(now);
  const rows = await fetchHourlyRowsSinceIso(todayStart.toISOString());
  const todaysRows = filterBetween(rows, "logged_at_utc", todayStart, now);

  const byHour = {};
  for (let h = 0; h < 24; h += 1) {
    const key = String(h).padStart(2, "0");
    byHour[key] = { hour_utc: key, sales_amount: 0, orders: 0 };
  }

  todaysRows.forEach((row) => {
    const ts = new Date(row.logged_at_utc);
    if (Number.isNaN(ts.getTime())) {
      return;
    }
    const h = String(ts.getUTCHours()).padStart(2, "0");
    byHour[h].sales_amount += row.sales_amount || 0;
    byHour[h].orders += row.orders || 0;
  });

  const heatmap = Object.keys(byHour)
    .sort()
    .map((h) => ({
      hour_utc: h,
      sales_amount: round(byHour[h].sales_amount, 2),
      orders: round(byHour[h].orders, 2),
    }));

  return {
    updatedAt: now.toISOString(),
    day_utc: todayStart.toISOString().slice(0, 10),
    heatmap,
  };
}

async function getRefundWatchlist(limit) {
  const now = nowUtc();
  const monthStart = startOfUtcMonth(now);
  const [lines, includedOrderIds] = await Promise.all([
    fetchLineRowsSinceIso(monthStart.toISOString()),
    fetchReportableOrderIdSetWithin(monthStart, now),
  ]);
  const scopedLines = filterLinesByOrderIdSet(filterBetween(lines, "created_at_utc", monthStart, now), includedOrderIds);
  const products = aggregateProducts(scopedLines);

  const watchlist = products
    .map((p) => {
      const sold = p.units + p.returned_units;
      const returnRate = sold > 0 ? (p.returned_units / sold) * 100 : null;
      return {
        product_key: p.product_key,
        product_id: p.product_id,
        title: p.title,
        sold_units: round(sold, 2),
        returned_units: round(p.returned_units, 2),
        return_rate_pct: round(returnRate, 2),
        returned_revenue: round(p.returned_revenue, 2),
      };
    })
    .filter((p) => (p.returned_units || 0) > 0)
    .sort((a, b) => {
      const ar = Number.isFinite(a.return_rate_pct) ? a.return_rate_pct : -1;
      const br = Number.isFinite(b.return_rate_pct) ? b.return_rate_pct : -1;
      if (br !== ar) {
        return br - ar;
      }
      return (b.returned_revenue || 0) - (a.returned_revenue || 0);
    })
    .slice(0, limit);

  return {
    updatedAt: now.toISOString(),
    period: {
      start_utc: monthStart.toISOString(),
      end_utc: now.toISOString(),
    },
    products: watchlist,
  };
}

function buildQuality(rows) {
  if (!rows.length) {
    return {
      row_count: 0,
      unique_hour_keys: 0,
      expected_hours: 0,
      missing_hours: 0,
      duplicate_rows: 0,
      hours_without_sales: 0,
      hours_without_marketing: 0,
      latest_row_age_minutes: null,
      rows_per_day: [],
    };
  }

  const seenKey = new Set();
  const seenHour = new Set();
  let duplicateRows = 0;
  let hoursWithoutSales = 0;
  let hoursWithoutMarketing = 0;
  const perDay = {};

  rows.forEach((row) => {
    const key = String(row.row_key || "").trim();
    if (key) {
      if (seenKey.has(key)) {
        duplicateRows += 1;
      } else {
        seenKey.add(key);
      }
    }

    const hour = String(row.logged_at_utc || "").slice(0, 13);
    if (hour) {
      seenHour.add(hour);
    }

    const day = String(row.logged_at_utc || "").slice(0, 10);
    if (day) {
      perDay[day] = (perDay[day] || 0) + 1;
    }

    if (!Number.isFinite(row.sales_amount) && !Number.isFinite(row.orders)) {
      hoursWithoutSales += 1;
    }
    if (!Number.isFinite(row.ad_spend) && !Number.isFinite(row.roas)) {
      hoursWithoutMarketing += 1;
    }
  });

  const firstTs = Date.parse(rows[0].logged_at_utc || "");
  const lastTs = Date.parse(rows[rows.length - 1].logged_at_utc || "");
  const expectedHours =
    Number.isFinite(firstTs) && Number.isFinite(lastTs) && lastTs >= firstTs
      ? Math.floor((lastTs - firstTs) / (60 * 60 * 1000)) + 1
      : seenHour.size;
  const missingHours = Math.max(0, expectedHours - seenHour.size);
  const latestAgeMins = Number.isFinite(lastTs) ? Math.floor((Date.now() - lastTs) / (60 * 1000)) : null;

  const rowsPerDay = Object.keys(perDay)
    .sort()
    .map((day) => ({ day, rows: perDay[day] }));

  return {
    row_count: rows.length,
    unique_hour_keys: seenHour.size,
    expected_hours: expectedHours,
    missing_hours: missingHours,
    duplicate_rows: duplicateRows,
    hours_without_sales: hoursWithoutSales,
    hours_without_marketing: hoursWithoutMarketing,
    latest_row_age_minutes: latestAgeMins,
    rows_per_day: rowsPerDay,
  };
}

function buildSourceCoverage(rows) {
  const sourceSales = {};
  const sourceMarketing = {};
  let both = 0;
  let salesOnly = 0;
  let marketingOnly = 0;
  let neither = 0;

  rows.forEach((row) => {
    const s = row.source_sales || "unknown";
    const m = row.source_marketing || "unknown";
    sourceSales[s] = (sourceSales[s] || 0) + 1;
    sourceMarketing[m] = (sourceMarketing[m] || 0) + 1;

    const hasSales = Number.isFinite(row.sales_amount) || Number.isFinite(row.orders);
    const hasMarketing = Number.isFinite(row.ad_spend) || Number.isFinite(row.roas);

    if (hasSales && hasMarketing) {
      both += 1;
      return;
    }
    if (hasSales) {
      salesOnly += 1;
      return;
    }
    if (hasMarketing) {
      marketingOnly += 1;
      return;
    }
    neither += 1;
  });

  return {
    rows: rows.length,
    source_sales_counts: sourceSales,
    source_marketing_counts: sourceMarketing,
    coverage: {
      both,
      sales_only: salesOnly,
      marketing_only: marketingOnly,
      neither,
    },
  };
}
