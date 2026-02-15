(function () {
  const DASHBOARD_CURRENCY = "EUR";

  const fallbackConfig = {
    appsScriptUrl: "",
    backendApiUrl: "",
    sheetName: "Triple Whale Hourly",
    refreshIntervalMs: 60 * 60 * 1000,
    cycleIntervalMs: 15 * 1000,
    passcode: "",
    salesTargetMultiplier: 1,
    salesTargetValue: null,
    salesTargetsByMonth: {},
    currencySymbol: DASHBOARD_CURRENCY,
  };

  const cfg = normalizeConfig({ ...fallbackConfig, ...(window.ELAVE_DASH_CONFIG || {}) });

  const LAYOUT_STORAGE_KEY = "elave_dash_widget_layout_v2";
  const MONTHLY_TARGETS_STORAGE_KEY = "elave_dash_monthly_targets_v1";
  const METRIC_WIDGET_PREFIX = "metric:";
  const ADVANCED_WIDGET_PREFIX = "advanced:";
  const ADVANCED_CELL_LIST_LIMIT = 5;
  const clockTimeFormatter = new Intl.DateTimeFormat(undefined, {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
    hourCycle: "h23",
  });
  const clockDateFormatter = new Intl.DateTimeFormat(undefined, {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
  });
  const fixedWidgetSeed = {
    goal_panel: {
      key: "goal_panel",
      label: "Target Progress",
      type: "fixed",
      view: "goal",
      colSpan: 6,
      rowSpan: 2,
      accentClass: "widget-goal",
    },
    chart_panel: {
      key: "chart_panel",
      label: "Trend Chart",
      type: "fixed",
      view: "chart",
      colSpan: 6,
      rowSpan: 2,
      accentClass: "widget-chart",
    },
    clock_panel: {
      key: "clock_panel",
      label: "Clock + Date/Day",
      type: "fixed",
      view: "clock",
      colSpan: 3,
      rowSpan: 1,
      accentClass: "metric-clock",
    },
  };
  const DEFAULT_CARD_KEYS = [
    "goal_panel",
    `${METRIC_WIDGET_PREFIX}sales`,
    `${METRIC_WIDGET_PREFIX}orders`,
    `${METRIC_WIDGET_PREFIX}aov`,
    `${METRIC_WIDGET_PREFIX}roas`,
    "chart_panel",
  ];
  const metricSeed = {
    sales: {
      key: "sales",
      label: "Total Sales",
      valueCol: "Order Revenue (Current)",
      changeCol: "Order Revenue (% Change)",
      directionCol: "Order Revenue (Direction)",
      formatType: "currency",
      accentClass: "metric-sales",
    },
    orders: {
      key: "orders",
      label: "Orders",
      valueCol: "Orders (Current)",
      changeCol: "Orders (% Change)",
      directionCol: "Orders (Direction)",
      formatType: "count",
      accentClass: "metric-orders",
    },
    aov: {
      key: "aov",
      label: "AOV",
      valueCol: "True AOV (Current)",
      changeCol: "True AOV (% Change)",
      directionCol: "True AOV (Direction)",
      formatType: "currency",
      accentClass: "metric-aov",
    },
    roas: {
      key: "roas",
      label: "ROAS",
      valueCol: "Blended ROAS (Current)",
      changeCol: "Blended ROAS (% Change)",
      directionCol: "Blended ROAS (Direction)",
      formatType: "ratio",
      accentClass: "metric-roas",
    },
  };
  const advancedCellSeed = [
    { key: "top_products_units", label: "Top Products (MTD by Units)", colSpan: 3, rowSpan: 2 },
    { key: "top_products_revenue", label: "Top Products (MTD by Revenue)", colSpan: 3, rowSpan: 2 },
    { key: "product_momentum", label: "Product Momentum", colSpan: 3, rowSpan: 2 },
    { key: "daily_sales_pace", label: "Daily Sales Pace vs Target", colSpan: 3, rowSpan: 2 },
    { key: "mtd_projection", label: "MTD Progress + Projection", colSpan: 3, rowSpan: 2 },
    { key: "gross_net_returns", label: "Gross / Net / Returns", colSpan: 3, rowSpan: 2 },
    { key: "aov", label: "AOV (MTD)", colSpan: 3, rowSpan: 2 },
    { key: "new_vs_returning", label: "New vs Returning Revenue", colSpan: 3, rowSpan: 2 },
    { key: "channel_split", label: "Channel Split", colSpan: 3, rowSpan: 2 },
    { key: "discount_impact", label: "Discount Impact", colSpan: 3, rowSpan: 2 },
    { key: "hourly_heatmap_today", label: "Hourly Heatmap (Today)", colSpan: 3, rowSpan: 2 },
    { key: "refund_watchlist", label: "Refund Watchlist", colSpan: 3, rowSpan: 2 },
  ];

  const state = {
    rows: [],
    cellsPayload: null,
    metricCatalog: {},
    cardLayout: [],
    monthlyTargets: {},
    hasCustomLayout: false,
    chart: null,
    currentMetric: "sales",
    chartType: "line",
    chartVisible: false,
    activeDragMetric: "",
    activeDragSource: "",
    activeDragCardEl: null,
    dragDropSlot: null,
    cycleTimer: null,
    now: new Date(),
    clockTimer: null,
    bootHideTimer: null,
  };

  const localFallbackRows = [
    {
      MONTH: "Jan",
      "Current Date Range": "Jan 1-Jan 31, 2026",
      "Order Revenue (Current)": 40250,
      "Order Revenue (% Change)": 4.1,
      "Order Revenue (Direction)": "up",
      "Blended ROAS (Current)": "1.66",
      "Blended ROAS (% Change)": "3.20",
      "Blended ROAS (Direction)": "up",
      "True AOV (Current)": 28,
      "True AOV (% Change)": "1.80",
      "True AOV (Direction)": "up",
      "Orders (Current)": 1460,
      "Orders (% Change)": 2.9,
      "Orders (Direction)": "up",
    },
    {
      MONTH: "Feb",
      "Current Date Range": "Feb 1-Feb 28, 2026",
      "Order Revenue (Current)": 41780,
      "Order Revenue (% Change)": 3.8,
      "Order Revenue (Direction)": "up",
      "Blended ROAS (Current)": "1.70",
      "Blended ROAS (% Change)": "2.40",
      "Blended ROAS (Direction)": "up",
      "True AOV (Current)": 29,
      "True AOV (% Change)": "3.60",
      "True AOV (Direction)": "up",
      "Orders (Current)": 1502,
      "Orders (% Change)": 2.8,
      "Orders (Direction)": "up",
    },
    {
      MONTH: "Mar",
      "Current Date Range": "Mar 1-Mar 31, 2026",
      "Order Revenue (Current)": 39840,
      "Order Revenue (% Change)": -4.6,
      "Order Revenue (Direction)": "down",
      "Blended ROAS (Current)": "1.58",
      "Blended ROAS (% Change)": "-7.10",
      "Blended ROAS (Direction)": "down",
      "True AOV (Current)": 27,
      "True AOV (% Change)": "-5.10",
      "True AOV (Direction)": "down",
      "Orders (Current)": 1442,
      "Orders (% Change)": -4.0,
      "Orders (Direction)": "down",
    },
    {
      MONTH: "Apr",
      "Current Date Range": "Apr 1-Apr 30, 2026",
      "Order Revenue (Current)": 43260,
      "Order Revenue (% Change)": 8.6,
      "Order Revenue (Direction)": "up",
      "Blended ROAS (Current)": "1.79",
      "Blended ROAS (% Change)": "13.20",
      "Blended ROAS (Direction)": "up",
      "True AOV (Current)": 30,
      "True AOV (% Change)": "8.90",
      "True AOV (Direction)": "up",
      "Orders (Current)": 1558,
      "Orders (% Change)": 8.0,
      "Orders (Direction)": "up",
    },
    {
      MONTH: "May",
      "Current Date Range": "May 1-May 31, 2026",
      "Order Revenue (Current)": 44610,
      "Order Revenue (% Change)": 3.1,
      "Order Revenue (Direction)": "up",
      "Blended ROAS (Current)": "1.82",
      "Blended ROAS (% Change)": "1.90",
      "Blended ROAS (Direction)": "up",
      "True AOV (Current)": 30,
      "True AOV (% Change)": "1.70",
      "True AOV (Direction)": "up",
      "Orders (Current)": 1604,
      "Orders (% Change)": 2.9,
      "Orders (Direction)": "up",
    },
    {
      MONTH: "Jun",
      "Current Date Range": "Jun 1-Jun 30, 2026",
      "Order Revenue (Current)": 45880,
      "Order Revenue (% Change)": 2.8,
      "Order Revenue (Direction)": "up",
      "Blended ROAS (Current)": "1.86",
      "Blended ROAS (% Change)": "2.10",
      "Blended ROAS (Direction)": "up",
      "True AOV (Current)": 31,
      "True AOV (% Change)": "2.30",
      "True AOV (Direction)": "up",
      "Orders (Current)": 1648,
      "Orders (% Change)": 2.7,
      "Orders (Direction)": "up",
    },
    {
      MONTH: "Jul",
      "Current Date Range": "Jul 1-Jul 31, 2026",
      "Order Revenue (Current)": 47240,
      "Order Revenue (% Change)": 3.0,
      "Order Revenue (Direction)": "up",
      "Blended ROAS (Current)": "1.91",
      "Blended ROAS (% Change)": "2.70",
      "Blended ROAS (Direction)": "up",
      "True AOV (Current)": 31,
      "True AOV (% Change)": "0.00",
      "True AOV (Direction)": "up",
      "Orders (Current)": 1692,
      "Orders (% Change)": 2.7,
      "Orders (Direction)": "up",
    },
    {
      MONTH: "Aug",
      "Current Date Range": "Aug 1-Aug 31, 2026",
      "Order Revenue (Current)": 48910,
      "Order Revenue (% Change)": 3.5,
      "Order Revenue (Direction)": "up",
      "Blended ROAS (Current)": "1.95",
      "Blended ROAS (% Change)": "2.10",
      "Blended ROAS (Direction)": "up",
      "True AOV (Current)": 32,
      "True AOV (% Change)": "3.20",
      "True AOV (Direction)": "up",
      "Orders (Current)": 1740,
      "Orders (% Change)": 2.8,
      "Orders (Direction)": "up",
    },
    {
      MONTH: "Sep",
      "Current Date Range": "Sep 1-Sep 30, 2026",
      "Order Revenue (Current)": 45190,
      "Order Revenue (% Change)": -7.6,
      "Order Revenue (Direction)": "down",
      "Blended ROAS (Current)": "1.73",
      "Blended ROAS (% Change)": "-11.30",
      "Blended ROAS (Direction)": "down",
      "True AOV (Current)": 29,
      "True AOV (% Change)": "-9.40",
      "True AOV (Direction)": "down",
      "Orders (Current)": 1598,
      "Orders (% Change)": -8.2,
      "Orders (Direction)": "down",
    },
    {
      MONTH: "Oct",
      "Current Date Range": "Oct 1-Oct 31, 2026",
      "Order Revenue (Current)": 48950,
      "Order Revenue (% Change)": 8.32,
      "Order Revenue (Direction)": "up",
      "Blended ROAS (Current)": "1.88",
      "Blended ROAS (% Change)": "8.67",
      "Blended ROAS (Direction)": "up",
      "True AOV (Current)": 31,
      "True AOV (% Change)": "6.90",
      "True AOV (Direction)": "up",
      "Orders (Current)": 1716,
      "Orders (% Change)": 7.39,
      "Orders (Direction)": "up",
    },
  ];

  const dom = {
    bootStatus: document.getElementById("boot-status"),
    app: document.getElementById("app"),
    lockScreen: document.getElementById("lock-screen"),
    unlockForm: document.getElementById("unlock-form"),
    unlockBtn: document.getElementById("unlock-btn"),
    passcodeInput: document.getElementById("passcode-input"),
    showPassToggle: document.getElementById("show-pass-toggle"),
    lockError: document.getElementById("lock-error"),
    metricSelect: document.getElementById("metric-select"),
    chartTypeSelect: document.getElementById("chart-type-select"),
    cycleToggle: document.getElementById("cycle-toggle"),
    toggleChartBtn: document.getElementById("toggle-chart-btn"),
    chartContent: document.getElementById("chart-content"),
    metricsBoard: document.getElementById("metrics-board"),
    cellMenu: document.getElementById("cell-menu"),
    cellMenuToggle: document.getElementById("cell-menu-toggle"),
    cellMenuClose: document.getElementById("cell-menu-close"),
    cellLibrary: document.getElementById("cell-library"),
    selectedMetricTrend: document.getElementById("selected-metric-trend"),
    refreshBtn: document.getElementById("refresh-btn"),
    lastUpdated: document.getElementById("last-updated"),
    clockDisplay: document.getElementById("clock-display"),
    clockTime: document.getElementById("clock-time"),
    clockHours: document.getElementById("clock-hours"),
    clockMinutes: document.getElementById("clock-minutes"),
    clockSeconds: document.getElementById("clock-seconds"),
    clockDate: document.getElementById("clock-date"),
    salesGoalStatus: document.getElementById("sales-goal-status"),
    chartCanvas: document.getElementById("main-chart"),
    salesGoalCurrentValue: document.getElementById("sales-goal-current-value"),
    salesGoalTargetValue: document.getElementById("sales-goal-target-value"),
    salesGoalPrevValue: document.getElementById("sales-goal-prev-value"),
    salesGoalPrevNote: document.getElementById("sales-goal-prev-note"),
    salesGoalMilestonePrev: document.getElementById("sales-goal-milestone-prev"),
    salesGoalMilestoneTarget: document.getElementById("sales-goal-milestone-target"),
    salesGoalProgressLabel: document.getElementById("sales-goal-progress-label"),
    salesGoalProgressFill: document.getElementById("sales-goal-progress-fill"),
    salesGoalGapPrev: document.getElementById("sales-goal-gap-prev"),
    salesGoalGapTarget: document.getElementById("sales-goal-gap-target"),
  };

  init();

  function init() {
    setBootStatus("Init start");
    try {
      setBootStatus("Binding events");
      bindEvents();
      const savedLayout = loadCardLayout();
      if (Array.isArray(savedLayout)) {
        state.cardLayout = savedLayout;
        state.hasCustomLayout = true;
      }
      state.monthlyTargets = loadMonthlyTargets();
      setBootStatus("Checking lock");
      checkLock();
      setBootStatus("Init complete", "ok");
    } catch (error) {
      console.error(error);
      setBootStatus("Init error", "err");
      if (dom.lockError) {
        dom.lockError.textContent = "Dashboard script failed to initialize.";
      }
    }
  }

  function bindEvents() {
    if (dom.unlockForm) {
      dom.unlockForm.addEventListener("submit", onUnlock);
    }
    if (dom.unlockBtn) {
      dom.unlockBtn.addEventListener("click", onUnlock);
    }
    if (dom.passcodeInput) {
      dom.passcodeInput.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          onUnlock(event);
        }
      });
    }
    if (dom.showPassToggle) {
      dom.showPassToggle.addEventListener("change", onToggleShowPasscode);
    }
    if (dom.refreshBtn) {
      dom.refreshBtn.addEventListener("click", fetchAndRender);
    }
    if (dom.cellMenuToggle) {
      dom.cellMenuToggle.addEventListener("click", () => {
        toggleCellMenu();
      });
    }
    if (dom.cellMenuClose) {
      dom.cellMenuClose.addEventListener("click", () => {
        toggleCellMenu(false);
      });
    }
    if (dom.cellLibrary) {
      dom.cellLibrary.addEventListener("dragstart", onLibraryDragStart);
      dom.cellLibrary.addEventListener("click", onLibraryClickAdd);
    }
    if (dom.metricsBoard) {
      dom.metricsBoard.addEventListener("click", onMetricsBoardClick);
      dom.metricsBoard.addEventListener("change", onMetricsBoardChange);
      dom.metricsBoard.addEventListener("dragstart", onMetricCardDragStart);
      dom.metricsBoard.addEventListener("dragover", onMetricsBoardDragOver);
      dom.metricsBoard.addEventListener("drop", onMetricsBoardDrop);
      dom.metricsBoard.addEventListener("dragleave", onMetricsBoardDragLeave);
      dom.metricsBoard.addEventListener("dragend", onAnyDragEnd);
    }
    window.addEventListener("dragend", onAnyDragEnd);
  }

  function onToggleChartVisibility() {
    state.chartVisible = !state.chartVisible;
    if (dom.chartContent) {
      dom.chartContent.classList.toggle("hidden", !state.chartVisible);
    }
    if (dom.toggleChartBtn) {
      dom.toggleChartBtn.textContent = state.chartVisible ? "Hide Chart" : "Show Chart";
    }
    if (state.chartVisible) {
      renderChart();
      return;
    }
    if (state.chart) {
      state.chart.destroy();
      state.chart = null;
    }
  }

  function checkLock() {
    if (!cfg.passcode) {
      unlockDashboard();
      return;
    }

    const queryPasscode = getQueryPasscode();
    if (queryPasscode && queryPasscode === normalizePasscodeForCompare(cfg.passcode)) {
      unlockDashboard();
      return;
    }
  }

  function onToggleShowPasscode(event) {
    if (!dom.passcodeInput) {
      return;
    }
    dom.passcodeInput.type = event.target.checked ? "text" : "password";
  }

  function onUnlock(event) {
    if (event && typeof event.preventDefault === "function") {
      event.preventDefault();
    }
    if (!dom.passcodeInput) {
      return;
    }
    const entered = normalizePasscodeForCompare(dom.passcodeInput.value);
    const expected = normalizePasscodeForCompare(cfg.passcode);
    if (!expected) {
      unlockDashboard();
      return;
    }
    if (entered !== expected) {
      dom.lockError.textContent = `Incorrect passcode (expected ${expected.length} digits)`;
      return;
    }
    dom.lockError.textContent = "";
    dom.passcodeInput.value = "";
    unlockDashboard();
  }

  function unlockDashboard() {
    setBootStatus("Unlocking dashboard");
    if (dom.lockScreen) {
      dom.lockScreen.classList.add("hidden");
    }
    if (dom.app) {
      dom.app.classList.remove("hidden");
    }
    startClockTicker();
    fetchAndRender();
    window.setInterval(fetchAndRender, cfg.refreshIntervalMs);
  }

  function startClockTicker() {
    if (state.clockTimer) {
      return;
    }
    updateNow();
    state.clockTimer = window.setInterval(updateNow, 1000);
  }

  function updateNow() {
    state.now = new Date();
    updateClockWidget();
  }

  function updateClockWidget() {
    if (
      !dom.clockDisplay ||
      !dom.clockTime ||
      !dom.clockHours ||
      !dom.clockMinutes ||
      !dom.clockSeconds ||
      !dom.clockDate
    ) {
      return;
    }
    const clockParts = getClockDisplayParts(state.now);
    dom.clockHours.textContent = clockParts.hours;
    dom.clockMinutes.textContent = clockParts.minutes;
    dom.clockSeconds.textContent = clockParts.seconds;
    dom.clockDisplay.classList.toggle("is-colon-dim", clockParts.isColonDim);
    dom.clockTime.setAttribute("aria-label", `Local time ${clockParts.time}`);
    dom.clockDate.textContent = clockParts.dateDay;
  }

  async function fetchAndRender() {
    setBootStatus("Loading data");
    try {
      const [rows, cellsPayload] = await Promise.all([loadRows(), loadAdvancedCells()]);
      state.rows = rows;
      state.cellsPayload = cellsPayload;
      updateKpis(rows);
      if (state.chartVisible) {
        renderChart();
      }
      dom.lastUpdated.textContent = `Last updated: ${new Date().toLocaleString()}`;
      setBootStatus(`Loaded ${rows.length} rows${cellsPayload ? " + advanced data" : ""}`, "ok");
    } catch (error) {
      dom.lastUpdated.textContent = "Last updated: error";
      setBootStatus("Data/chart error", "err");
      console.error(error);
    }
  }

  async function loadAdvancedCells() {
    if (!cfg.backendApiUrl) {
      return null;
    }
    try {
      const payload = await fetchBackendPayload("/api/cells");
      return payload && typeof payload === "object" ? payload : null;
    } catch (error) {
      console.warn("Advanced cells endpoint unavailable; attempting partial backend payload.", error);
      try {
        const endpoints = [
          ["/api/summary", "summary"],
          ["/api/pace", "daily_sales_pace"],
          ["/api/projection", "mtd_projection"],
          ["/api/finance/gross-net-returns", "gross_net_returns"],
          ["/api/aov", "aov"],
          ["/api/customers/new-vs-returning", "new_vs_returning"],
          ["/api/channels", "channel_split"],
          ["/api/discount-impact", "discount_impact"],
          ["/api/heatmap/today", "hourly_heatmap_today"],
          [`/api/refund-watchlist?limit=${ADVANCED_CELL_LIST_LIMIT}`, "refund_watchlist"],
          [`/api/products/top-units?limit=${ADVANCED_CELL_LIST_LIMIT}`, "top_products_units"],
          [`/api/products/top-revenue?limit=${ADVANCED_CELL_LIST_LIMIT}`, "top_products_revenue"],
          [`/api/products/momentum?metric=revenue&limit=${ADVANCED_CELL_LIST_LIMIT}`, "product_momentum"],
        ];

        const settled = await Promise.allSettled(endpoints.map(([path]) => fetchBackendPayload(path)));
        const partial = {
          updatedAt: new Date().toISOString(),
          errors: {},
        };

        settled.forEach((result, idx) => {
          const [, key] = endpoints[idx];
          if (result.status === "fulfilled") {
            partial[key] = result.value;
            return;
          }
          partial[key] = null;
          partial.errors[key] = String(
            result.reason && result.reason.message ? result.reason.message : result.reason || "Unknown backend error"
          );
        });

        if (partial.summary && partial.summary.summary) {
          partial.kpis = partial.summary.kpis || null;
          partial.summary = partial.summary.summary;
        } else if (!partial.kpis) {
          partial.kpis = null;
        }

        const hasUsefulData =
          partial.kpis ||
          partial.daily_sales_pace ||
          partial.mtd_projection ||
          partial.top_products_units ||
          partial.top_products_revenue;

        return hasUsefulData ? partial : null;
      } catch (fallbackError) {
        console.warn("Partial backend payload unavailable.", fallbackError);
        return null;
      }
    }
  }

  async function loadRows() {
    if (cfg.appsScriptUrl) {
      const pipelineRows = await loadRowsFromPipeline();
      if (pipelineRows.length) {
        return pipelineRows;
      }

      const payload = await fetchAppsScriptPayload(`sheet=${encodeURIComponent(cfg.sheetName)}`);
      const data = extractPayloadRows(payload);
      return prepareRows(data.filter(rowHasData));
    }

    return prepareRows(localFallbackRows.filter(rowHasData));
  }

  async function loadRowsFromPipeline() {
    if (cfg.backendApiUrl) {
      try {
        const payload = await fetchBackendPayload("/api/clean");
        const cleanRows = extractPayloadRows(payload);
        if (cleanRows.length) {
          return prepareMtdRowsFromClean(cleanRows);
        }
      } catch (error) {
        console.warn("Backend clean endpoint unavailable; falling back to Apps Script.", error);
      }
    }

    try {
      const payload = await fetchAppsScriptPayload("mode=clean");
      const cleanRows = extractPayloadRows(payload);
      if (!cleanRows.length) {
        return [];
      }
      return prepareMtdRowsFromClean(cleanRows);
    } catch (error) {
      console.warn("Pipeline clean mode unavailable; falling back to sheet endpoint.", error);
      return [];
    }
  }

  async function fetchBackendPayload(path) {
    const base = String(cfg.backendApiUrl || "").replace(/\/+$/, "");
    const safePath = path.startsWith("/") ? path : `/${path}`;
    const url = `${base}${safePath}?_=${Date.now()}`;
    const res = await fetch(url);
    if (!res.ok) {
      throw new Error(`Backend request failed (${res.status})`);
    }
    return res.json();
  }

  async function fetchAppsScriptPayload(query) {
    const separator = cfg.appsScriptUrl.includes("?") ? "&" : "?";
    const url = `${cfg.appsScriptUrl}${separator}${query}&_=${Date.now()}`;
    const res = await fetch(url);
    if (!res.ok) {
      throw new Error(`Apps Script request failed (${res.status})`);
    }
    return res.json();
  }

  function extractPayloadRows(payload) {
    const rows = Array.isArray(payload) ? payload : payload && payload.data;
    if (!Array.isArray(rows)) {
      throw new Error("Apps Script response has no data array");
    }
    return rows;
  }

  function prepareMtdRowsFromClean(cleanRows) {
    const parsed = cleanRows
      .map((row) => {
        const ts = getCleanRowTimestamp(row);
        if (!(ts instanceof Date) || !Number.isFinite(ts.getTime())) {
          return null;
        }
        const rowKey = String(row.row_key || "").trim();
        return {
          ts,
          rowKey,
          sales: parseNumber(row.sales_amount),
          orders: parseNumber(row.orders),
          adSpend: parseNumber(row.ad_spend),
          roas: parseNumber(row.roas),
        };
      })
      .filter(Boolean)
      .sort((a, b) => a.ts.getTime() - b.ts.getTime());

    if (!parsed.length) {
      return [];
    }

    const dedupedByKey = [];
    const seen = new Set();
    parsed.forEach((item) => {
      const fallbackKey = `${item.ts.getFullYear()}-${item.ts.getMonth()}-${item.ts.getDate()}-${item.ts.getHours()}`;
      const key = item.rowKey || fallbackKey;
      if (seen.has(key)) {
        return;
      }
      seen.add(key);
      dedupedByKey.push(item);
    });

    const monthRows = dedupedByKey.filter((item) => isCurrentMonth(item.ts));
    const rowsForMtd = monthRows.length ? monthRows : dedupedByKey;

    let runningSales = 0;
    let runningOrders = 0;
    let runningAdSpend = 0;
    let prevSales = NaN;
    let prevOrders = NaN;
    let prevAov = NaN;
    let prevRoas = NaN;

    const transformed = rowsForMtd.map((item) => {
      const sales = Number.isFinite(item.sales) ? item.sales : 0;
      const orders = Number.isFinite(item.orders) ? item.orders : 0;
      const adSpend = Number.isFinite(item.adSpend) ? item.adSpend : 0;

      runningSales += sales;
      runningOrders += orders;
      runningAdSpend += adSpend;

      const currentSales = roundTo(runningSales, 2);
      const currentOrders = roundTo(runningOrders, 2);
      const currentAov = runningOrders > 0 ? roundTo(runningSales / runningOrders, 2) : NaN;
      const currentRoas = runningAdSpend > 0 ? roundTo(runningSales / runningAdSpend, 4) : item.roas;

      const salesChange = calcPercentChange(currentSales, prevSales);
      const ordersChange = calcPercentChange(currentOrders, prevOrders);
      const aovChange = calcPercentChange(currentAov, prevAov);
      const roasChange = calcPercentChange(currentRoas, prevRoas);

      prevSales = currentSales;
      prevOrders = currentOrders;
      prevAov = currentAov;
      prevRoas = currentRoas;

      return {
        "Logged At": item.ts.toISOString(),
        "Current Date Range": "",
        "Order Revenue (Current)": currentSales,
        "Order Revenue (% Change)": Number.isFinite(salesChange) ? roundTo(salesChange, 2) : "",
        "Order Revenue (Direction)": trendFromNumber(salesChange),
        "Orders (Current)": currentOrders,
        "Orders (% Change)": Number.isFinite(ordersChange) ? roundTo(ordersChange, 2) : "",
        "Orders (Direction)": trendFromNumber(ordersChange),
        "True AOV (Current)": Number.isFinite(currentAov) ? currentAov : "",
        "True AOV (% Change)": Number.isFinite(aovChange) ? roundTo(aovChange, 2) : "",
        "True AOV (Direction)": trendFromNumber(aovChange),
        "Blended ROAS (Current)": Number.isFinite(currentRoas) ? currentRoas : "",
        "Blended ROAS (% Change)": Number.isFinite(roasChange) ? roundTo(roasChange, 2) : "",
        "Blended ROAS (Direction)": trendFromNumber(roasChange),
      };
    });

    return prepareRows(transformed.filter(rowHasData));
  }

  function getCleanRowTimestamp(row) {
    return parseTimestamp(row.logged_at_local || row.logged_at_utc || row["Logged At"]);
  }

  function buildAdvancedCellContent(cellKey, payload) {
    if (!payload || typeof payload !== "object") {
      return {
        subtitle: "",
        body: '<p class="advanced-empty">Advanced data unavailable. Start backend API and run Supabase sync.</p>',
      };
    }

    const sectionPayload = payload[cellKey];

    if (cellKey === "top_products_units") {
      return buildTopProductsUnitsContent(sectionPayload);
    }
    if (cellKey === "top_products_revenue") {
      return buildTopProductsRevenueContent(sectionPayload);
    }
    if (cellKey === "product_momentum") {
      return buildProductMomentumContent(sectionPayload);
    }
    if (cellKey === "daily_sales_pace") {
      return buildDailyPaceContent(sectionPayload);
    }
    if (cellKey === "mtd_projection") {
      return buildProjectionContent(sectionPayload);
    }
    if (cellKey === "gross_net_returns") {
      return buildGrossNetReturnsContent(sectionPayload);
    }
    if (cellKey === "aov") {
      return buildAovContent(sectionPayload);
    }
    if (cellKey === "new_vs_returning") {
      return buildNewVsReturningContent(sectionPayload);
    }
    if (cellKey === "channel_split") {
      return buildChannelSplitContent(sectionPayload);
    }
    if (cellKey === "discount_impact") {
      return buildDiscountImpactContent(sectionPayload);
    }
    if (cellKey === "hourly_heatmap_today") {
      return buildHeatmapContent(sectionPayload);
    }
    if (cellKey === "refund_watchlist") {
      return buildRefundWatchlistContent(sectionPayload);
    }

    return {
      subtitle: "",
      body: '<p class="advanced-empty">Cell is not configured.</p>',
    };
  }

  function buildTopProductsUnitsContent(payload) {
    const products = payload && Array.isArray(payload.products) ? payload.products : [];
    const subtitle = payload && Number.isFinite(payload.total_units) ? `Total units: ${formatNumber(payload.total_units, 0)}` : "";
    return {
      subtitle,
      body: renderProductList(products, (item) => [
        `${formatNumberSafe(item.units, 0)} units`,
        formatCurrencySafe(item.revenue),
        formatPercentSafe(item.unit_share_pct),
      ]),
    };
  }

  function buildTopProductsRevenueContent(payload) {
    const products = payload && Array.isArray(payload.products) ? payload.products : [];
    const subtitle = payload && Number.isFinite(payload.total_revenue) ? `Total revenue: ${formatCurrencySafe(payload.total_revenue)}` : "";
    return {
      subtitle,
      body: renderProductList(products, (item) => [
        formatCurrencySafe(item.revenue),
        `${formatNumberSafe(item.units, 0)} units`,
        formatPercentSafe(item.revenue_share_pct),
      ]),
    };
  }

  function buildProductMomentumContent(payload) {
    const products = payload && Array.isArray(payload.products) ? payload.products : [];
    const metric = payload && payload.metric ? String(payload.metric) : "revenue";
    const body = !products.length
      ? '<p class="advanced-empty">No momentum data yet.</p>'
      : `<ul class="advanced-list">${products
          .slice(0, ADVANCED_CELL_LIST_LIMIT)
          .map((item) => {
            const delta = Number.isFinite(item.delta) ? item.delta : null;
            const chipClass = delta > 0 ? "up" : delta < 0 ? "down" : "";
            const deltaText =
              metric === "units"
                ? `${delta > 0 ? "+" : ""}${formatNumberSafe(delta, 0)}`
                : `${delta > 0 ? "+" : ""}${formatCurrencySafe(delta)}`;
            return `
              <li>
                <div class="advanced-list-head">
                  <span>${escapeHtml(item.title || "Unknown")}</span>
                  <span class="advanced-chip ${chipClass}">${escapeHtml(deltaText)}</span>
                </div>
                <div class="advanced-list-meta">
                  <span>Current: ${metric === "units" ? formatNumberSafe(item.current_value, 0) : formatCurrencySafe(item.current_value)}</span>
                  <span>Prev: ${metric === "units" ? formatNumberSafe(item.previous_value, 0) : formatCurrencySafe(item.previous_value)}</span>
                </div>
              </li>
            `;
          })
          .join("")}</ul>`;
    return {
      subtitle: `Metric: ${metric}`,
      body,
    };
  }

  function buildDailyPaceContent(payload) {
    const mtdSales = parseNumber(payload && payload.mtd_sales);
    const prevMonthTotal = parseNumber(payload && payload.previous_month_total_sales);
    return {
      subtitle: "",
      body: renderKeyValueList([
        ["Month Goal", formatCurrencySafe(payload && payload.month_goal)],
        ["MTD Total Sales", formatCurrencySafe(Number.isFinite(mtdSales) ? mtdSales : payload && payload.mtd_gross_sales)],
        ["Last Month Total", formatCurrencySafe(prevMonthTotal)],
        ["Today Sales", formatCurrencySafe(payload && payload.today_sales)],
        ["Required Daily Pace", formatCurrencySafe(payload && payload.required_daily_pace)],
        ["Days Remaining", formatNumberSafe(payload && payload.days_remaining, 0)],
      ]),
    };
  }

  function buildProjectionContent(payload) {
    const mtdSales = parseNumber(payload && payload.mtd_sales);
    return {
      subtitle: "",
      body: renderKeyValueList([
        ["MTD Total Sales", formatCurrencySafe(Number.isFinite(mtdSales) ? mtdSales : payload && payload.mtd_gross_sales)],
        ["Month Goal", formatCurrencySafe(payload && payload.month_goal)],
        ["Progress", formatPercentSafe(payload && payload.progress_pct_of_target)],
        ["Projected Month-End", formatCurrencySafe(payload && payload.projected_month_end_sales)],
        ["Projected vs Target", formatPercentSafe(payload && payload.projected_vs_target_pct)],
      ]),
    };
  }

  function buildGrossNetReturnsContent(payload) {
    return {
      subtitle: "",
      body: renderKeyValueList([
        ["Gross Sales", formatCurrencySafe(payload && payload.gross_sales)],
        ["Net Sales", formatCurrencySafe(payload && payload.net_sales)],
        ["Total Sales", formatCurrencySafe(payload && payload.total_sales)],
        ["Returns", formatCurrencySafe(payload && payload.returns_amount)],
        ["Returns % of Net", formatPercentSafe(payload && payload.returns_rate_pct_of_net)],
      ]),
    };
  }

  function buildAovContent(payload) {
    return {
      subtitle: "",
      body: renderKeyValueList([
        ["MTD AOV", formatCurrencySafe(payload && payload.mtd_aov)],
        ["Previous AOV", formatCurrencySafe(payload && payload.previous_period_aov)],
        ["AOV Change", formatPercentSafe(payload && payload.aov_change_pct)],
        ["MTD Orders", formatNumberSafe(payload && payload.mtd_orders, 0)],
      ]),
    };
  }

  function buildNewVsReturningContent(payload) {
    const revenue = payload && payload.revenue ? payload.revenue : {};
    const shares = payload && payload.shares_pct ? payload.shares_pct : {};
    return {
      subtitle: "",
      body: renderKeyValueList([
        ["New Revenue", `${formatCurrencySafe(revenue.new)} (${formatPercentSafe(shares.new)})`],
        ["Returning Revenue", `${formatCurrencySafe(revenue.returning)} (${formatPercentSafe(shares.returning)})`],
        ["Unknown Revenue", `${formatCurrencySafe(revenue.unknown)} (${formatPercentSafe(shares.unknown)})`],
      ]),
    };
  }

  function buildChannelSplitContent(payload) {
    const channels = payload && Array.isArray(payload.channels) ? payload.channels : [];
    const body = !channels.length
      ? '<p class="advanced-empty">No channel data yet.</p>'
      : `<ul class="advanced-list">${channels
          .slice(0, ADVANCED_CELL_LIST_LIMIT)
          .map(
            (item) => `
              <li>
                <div class="advanced-list-head">
                  <span>${escapeHtml(item.channel || "unknown")}</span>
                  <span>${formatCurrencySafe(item.revenue)}</span>
                </div>
                <div class="advanced-list-meta">
                  <span>Orders: ${formatNumberSafe(item.orders, 0)}</span>
                  <span>Share: ${formatPercentSafe(item.revenue_share_pct)}</span>
                </div>
              </li>
            `
          )
          .join("")}</ul>`;
    return {
      subtitle: "",
      body,
    };
  }

  function buildDiscountImpactContent(payload) {
    return {
      subtitle: "",
      body: renderKeyValueList([
        ["Discounted Orders %", formatPercentSafe(payload && payload.discounted_orders_pct)],
        ["Total Discounts", formatCurrencySafe(payload && payload.total_discounts)],
        ["Avg Discount / Order", formatCurrencySafe(payload && payload.avg_discount_per_order)],
        ["Discount % of Gross", formatPercentSafe(payload && payload.discount_rate_pct_of_gross)],
      ]),
    };
  }

  function buildHeatmapContent(payload) {
    const heatmap = payload && Array.isArray(payload.heatmap) ? payload.heatmap : [];
    const body = !heatmap.length
      ? '<p class="advanced-empty">No hourly data yet.</p>'
      : `<div class="advanced-heatmap">${heatmap
          .map(
            (item) => `
              <div class="advanced-heat">
                <span class="advanced-heat-hour">${escapeHtml(item.hour_utc)}:00</span>
                <span class="advanced-heat-value">${formatNumberSafe(item.orders, 0)} ord</span>
              </div>
            `
          )
          .join("")}</div>`;
    return {
      subtitle: "",
      body,
    };
  }

  function buildRefundWatchlistContent(payload) {
    const products = payload && Array.isArray(payload.products) ? payload.products : [];
    const body = !products.length
      ? '<p class="advanced-empty">No significant refunds this month.</p>'
      : `<ul class="advanced-list">${products
          .slice(0, ADVANCED_CELL_LIST_LIMIT)
          .map(
            (item) => `
              <li>
                <div class="advanced-list-head">
                  <span>${escapeHtml(item.title || "Unknown")}</span>
                  <span>${formatPercentSafe(item.return_rate_pct)}</span>
                </div>
                <div class="advanced-list-meta">
                  <span>Returned units: ${formatNumberSafe(item.returned_units, 0)}</span>
                  <span>Returned rev: ${formatCurrencySafe(item.returned_revenue)}</span>
                </div>
              </li>
            `
          )
          .join("")}</ul>`;
    return {
      subtitle: "",
      body,
    };
  }

  function renderProductList(products, metaBuilder) {
    if (!products.length) {
      return '<p class="advanced-empty">No product data yet.</p>';
    }

    return `<ul class="advanced-list">${products
      .slice(0, ADVANCED_CELL_LIST_LIMIT)
      .map((item) => {
        const meta = metaBuilder(item);
        return `
          <li>
            <div class="advanced-list-head">
              <span>#${escapeHtml(String(item.rank || ""))} ${escapeHtml(item.title || "Unknown")}</span>
            </div>
            <div class="advanced-list-meta">
              <span>${escapeHtml(meta[0])}</span>
              <span>${escapeHtml(meta[1])}</span>
              <span>${escapeHtml(meta[2])}</span>
            </div>
          </li>
        `;
      })
      .join("")}</ul>`;
  }

  function renderKeyValueList(entries) {
    return `<ul class="advanced-kv">${entries
      .map(
        ([label, value]) => `
          <li>
            <span class="advanced-kv-label">${escapeHtml(label)}</span>
            <span class="advanced-kv-value">${escapeHtml(value)}</span>
          </li>
        `
      )
      .join("")}</ul>`;
  }

  function formatCurrencySafe(value) {
    return Number.isFinite(value) ? formatCurrency(value, cfg.currencySymbol) : "--";
  }

  function formatNumberSafe(value, digits) {
    return Number.isFinite(value) ? formatNumber(value, digits) : "--";
  }

  function formatPercentSafe(value) {
    if (!Number.isFinite(value)) {
      return "--";
    }
    const sign = value > 0 ? "+" : "";
    return `${sign}${formatNumber(value, 0)}%`;
  }

  function prepareRows(rows) {
    if (!Array.isArray(rows) || !rows.length) {
      return [];
    }

    const enriched = rows.map((row) => ({
      ...row,
      __ts: getRowTimestamp(row),
    }));

    const timedRows = enriched.filter((row) => row.__ts instanceof Date && Number.isFinite(row.__ts.getTime()));
    if (!timedRows.length) {
      return enriched;
    }

    const mtdRows = timedRows
      .filter((row) => isCurrentMonth(row.__ts))
      .sort((a, b) => a.__ts.getTime() - b.__ts.getTime());

    return mtdRows.length ? mtdRows : timedRows.sort((a, b) => a.__ts.getTime() - b.__ts.getTime());
  }

  function rowHasData(row) {
    return Object.values(row).some((value) => value !== null && value !== "");
  }

  function updateKpis(rows) {
    syncMetricCatalog(rows);
    ensureValidCardLayout();
    renderWidgetBoard(rows);
    syncChartMetricOptions();
    renderCellLibrary();

    if (!rows.length) {
      clearSalesGoal();
      updateSelectedMetricTrend(null);
      return;
    }

    const latest = rows[rows.length - 1];
    const previous = rows.length > 1 ? rows[rows.length - 2] : null;
    updateSalesGoal(latest, previous, getSalesGoalMetric());
    updateSelectedMetricTrend(latest);
  }

  function syncMetricCatalog(rows) {
    const headers = collectRowHeaders(rows);
    const catalog = {};

    Object.values(metricSeed).forEach((seedMeta) => {
      const shouldAdd = headers.size ? headers.has(seedMeta.valueCol) : true;
      if (!shouldAdd) {
        return;
      }
      catalog[seedMeta.key] = { ...seedMeta };
    });

    const currentCols = Array.from(headers)
      .filter((header) => /\(Current\)\s*$/i.test(header))
      .sort((a, b) => a.localeCompare(b));

    currentCols.forEach((currentCol) => {
      const alreadyExists = Object.values(catalog).some((meta) => meta.valueCol === currentCol);
      if (alreadyExists) {
        return;
      }

      const baseLabel = currentCol.replace(/\s*\(Current\)\s*$/i, "").trim() || currentCol;
      const changeCol = currentCol.replace(/\(Current\)\s*$/i, "(% Change)");
      const directionCol = currentCol.replace(/\(Current\)\s*$/i, "(Direction)");
      const metricKey = toMetricKey(baseLabel, catalog);

      catalog[metricKey] = {
        key: metricKey,
        label: baseLabel,
        valueCol: currentCol,
        changeCol: headers.has(changeCol) ? changeCol : null,
        directionCol: headers.has(directionCol) ? directionCol : null,
        formatType: inferMetricFormat(baseLabel),
        accentClass: "metric-generic",
      };
    });

    state.metricCatalog = catalog;
  }

  function collectRowHeaders(rows) {
    const headers = new Set();
    rows.forEach((row) => {
      if (!row || typeof row !== "object") {
        return;
      }
      Object.keys(row).forEach((key) => {
        if (!key || key === "__ts") {
          return;
        }
        headers.add(key);
      });
    });
    return headers;
  }

  function toMetricKey(label, existingCatalog) {
    const base = String(label || "metric")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "") || "metric";

    let key = base;
    let index = 2;
    while (existingCatalog[key]) {
      key = `${base}_${index}`;
      index += 1;
    }
    return key;
  }

  function inferMetricFormat(label) {
    const text = String(label || "").toLowerCase();
    if (/%|percent|conversion|margin|returns|add to cart/.test(text)) {
      return "percent";
    }
    if (/roas|mer/.test(text)) {
      return "ratio";
    }
    if (/revenue|sales|spend|profit|aov|cpa/.test(text)) {
      return "currency";
    }
    if (/orders|users|sessions/.test(text)) {
      return "count";
    }
    return "number";
  }

  function toMetricWidgetKey(metricKey) {
    return `${METRIC_WIDGET_PREFIX}${metricKey}`;
  }

  function toAdvancedWidgetKey(cellKey) {
    return `${ADVANCED_WIDGET_PREFIX}${cellKey}`;
  }

  function isMetricWidgetKey(widgetKey) {
    return typeof widgetKey === "string" && widgetKey.startsWith(METRIC_WIDGET_PREFIX);
  }

  function isAdvancedWidgetKey(widgetKey) {
    return typeof widgetKey === "string" && widgetKey.startsWith(ADVANCED_WIDGET_PREFIX);
  }

  function metricKeyFromWidgetKey(widgetKey) {
    if (!isMetricWidgetKey(widgetKey)) {
      return "";
    }
    return widgetKey.slice(METRIC_WIDGET_PREFIX.length);
  }

  function advancedKeyFromWidgetKey(widgetKey) {
    if (!isAdvancedWidgetKey(widgetKey)) {
      return "";
    }
    return widgetKey.slice(ADVANCED_WIDGET_PREFIX.length);
  }

  function getAvailableWidgetDefinitions() {
    const fixedWidgets = Object.values(fixedWidgetSeed);
    const metricWidgets = Object.values(state.metricCatalog).map((meta) => ({
      key: toMetricWidgetKey(meta.key),
      label: meta.label,
      type: "metric",
      metricMeta: meta,
      colSpan: 3,
      rowSpan: 1,
      accentClass: meta.accentClass || "metric-generic",
    }));
    const advancedWidgets = advancedCellSeed.map((cell) => ({
      key: toAdvancedWidgetKey(cell.key),
      label: cell.label,
      type: "advanced",
      payloadKey: cell.key,
      colSpan: Number.isFinite(cell.colSpan) ? cell.colSpan : 3,
      rowSpan: Number.isFinite(cell.rowSpan) ? cell.rowSpan : 2,
      accentClass: "metric-generic",
    }));
    return [...fixedWidgets, ...metricWidgets, ...advancedWidgets];
  }

  function getAvailableWidgetKeys() {
    return getAvailableWidgetDefinitions().map((widget) => widget.key);
  }

  function getWidgetDefinition(widgetKey) {
    if (fixedWidgetSeed[widgetKey]) {
      return fixedWidgetSeed[widgetKey];
    }
    if (isMetricWidgetKey(widgetKey)) {
      const metricKey = metricKeyFromWidgetKey(widgetKey);
      const metricMeta = state.metricCatalog[metricKey];
      if (!metricMeta) {
        return null;
      }
      return {
        key: widgetKey,
        label: metricMeta.label,
        type: "metric",
        metricMeta,
        colSpan: 3,
        rowSpan: 1,
        accentClass: metricMeta.accentClass || "metric-generic",
      };
    }
    if (isAdvancedWidgetKey(widgetKey)) {
      const advancedKey = advancedKeyFromWidgetKey(widgetKey);
      const cellMeta = advancedCellSeed.find((cell) => cell.key === advancedKey);
      if (!cellMeta) {
        return null;
      }
      return {
        key: widgetKey,
        label: cellMeta.label,
        type: "advanced",
        payloadKey: cellMeta.key,
        colSpan: Number.isFinite(cellMeta.colSpan) ? cellMeta.colSpan : 3,
        rowSpan: Number.isFinite(cellMeta.rowSpan) ? cellMeta.rowSpan : 2,
        accentClass: "metric-generic",
      };
    }
    return null;
  }

  function isWidgetAvailable(widgetKey) {
    return Boolean(getWidgetDefinition(widgetKey));
  }

  function hasChartWidget() {
    return state.cardLayout.some((item) => item && item.key === "chart_panel");
  }

  function normalizeWidgetKey(key) {
    if (typeof key !== "string") {
      return "";
    }
    const trimmed = key.trim();
    if (!trimmed) {
      return "";
    }
    if (
      trimmed === "goal_panel" ||
      trimmed === "chart_panel" ||
      trimmed === "clock_panel" ||
      isMetricWidgetKey(trimmed) ||
      isAdvancedWidgetKey(trimmed)
    ) {
      return trimmed;
    }
    return toMetricWidgetKey(trimmed);
  }

  function normalizeLayoutEntry(entry) {
    if (typeof entry === "string") {
      const key = normalizeWidgetKey(entry);
      if (!key) {
        return null;
      }
      return { key, col: null, row: null };
    }
    if (!entry || typeof entry !== "object") {
      return null;
    }
    const key = normalizeWidgetKey(entry.key);
    if (!key) {
      return null;
    }
    const col = Number.parseInt(entry.col, 10);
    const row = Number.parseInt(entry.row, 10);
    return {
      key,
      col: Number.isFinite(col) ? col : null,
      row: Number.isFinite(row) ? row : null,
    };
  }

  function getLayoutGridConfig() {
    const geometry = getMetricGridGeometry();
    if (geometry) {
      return {
        slotCols: geometry.slotCols,
        metricCardSpan: geometry.metricCardSpan,
        gridCols: geometry.gridCols,
      };
    }
    return {
      slotCols: 4,
      metricCardSpan: 3,
      gridCols: 12,
    };
  }

  function getWidgetSlotSizeFromSpan(widgetKey, metricCardSpan) {
    const def = getWidgetDefinition(widgetKey);
    if (!def) {
      return { colSlots: 1, rowSlots: 1 };
    }
    const metricSpan = Math.max(1, metricCardSpan || 3);
    const colSlots = Math.max(1, Math.round((def.colSpan || metricSpan) / metricSpan));
    const rowSlots = Math.max(1, def.rowSpan || 1);
    return { colSlots, rowSlots };
  }

  function rectsOverlap(a, b) {
    return a.col < b.col + b.colSlots && a.col + a.colSlots > b.col && a.row < b.row + b.rowSlots && a.row + a.rowSlots > b.row;
  }

  function canPlaceWidgetAtSlot(widgetKey, col, row, slotCols, occupiedRects, metricCardSpan) {
    if (!Number.isFinite(col) || !Number.isFinite(row) || col < 1 || row < 1) {
      return false;
    }
    const size = getWidgetSlotSizeFromSpan(widgetKey, metricCardSpan);
    if (col + size.colSlots - 1 > slotCols) {
      return false;
    }
    const nextRect = {
      col,
      row,
      colSlots: size.colSlots,
      rowSlots: size.rowSlots,
    };
    return !occupiedRects.some((rect) => rectsOverlap(nextRect, rect));
  }

  function findFirstAvailableSlot(widgetKey, startCol, startRow, slotCols, occupiedRects, metricCardSpan) {
    const safeStartCol = clamp(Math.round(startCol || 1), 1, slotCols);
    const safeStartRow = Math.max(1, Math.round(startRow || 1));
    const scanLimit = Math.max(80, safeStartRow + 40);

    for (let row = safeStartRow; row <= scanLimit; row += 1) {
      const firstCol = row === safeStartRow ? safeStartCol : 1;
      for (let col = firstCol; col <= slotCols; col += 1) {
        if (canPlaceWidgetAtSlot(widgetKey, col, row, slotCols, occupiedRects, metricCardSpan)) {
          return { col, row };
        }
      }
    }

    for (let row = 1; row < safeStartRow; row += 1) {
      for (let col = 1; col <= slotCols; col += 1) {
        if (canPlaceWidgetAtSlot(widgetKey, col, row, slotCols, occupiedRects, metricCardSpan)) {
          return { col, row };
        }
      }
    }

    return { col: 1, row: scanLimit + 1 };
  }

  function buildOccupiedRects(layoutEntries, slotCols, metricCardSpan) {
    return layoutEntries
      .filter((entry) => entry && isWidgetAvailable(entry.key))
      .map((entry) => {
        const size = getWidgetSlotSizeFromSpan(entry.key, metricCardSpan);
        const col = clamp(Math.round(entry.col || 1), 1, slotCols);
        const maxCol = Math.max(1, slotCols - size.colSlots + 1);
        return {
          key: entry.key,
          col: clamp(col, 1, maxCol),
          row: Math.max(1, Math.round(entry.row || 1)),
          colSlots: size.colSlots,
          rowSlots: size.rowSlots,
        };
      });
  }

  function buildDefaultCardLayout(widgetKeys, slotCols, metricCardSpan) {
    const layout = [];
    const occupied = [];
    widgetKeys.forEach((key) => {
      const slot = findFirstAvailableSlot(key, 1, 1, slotCols, occupied, metricCardSpan);
      const size = getWidgetSlotSizeFromSpan(key, metricCardSpan);
      layout.push({ key, col: slot.col, row: slot.row });
      occupied.push({
        key,
        col: slot.col,
        row: slot.row,
        colSlots: size.colSlots,
        rowSlots: size.rowSlots,
      });
    });
    return layout;
  }

  function sortLayoutEntries(entries) {
    return [...entries].sort((a, b) => {
      if (a.row !== b.row) {
        return a.row - b.row;
      }
      return a.col - b.col;
    });
  }

  function isWidgetOnBoard(widgetKey) {
    return state.cardLayout.some((item) => item && item.key === widgetKey);
  }

  function buildWidgetPlacementStyle(entry, widget, gridConfig) {
    const metricSpan = Math.max(1, gridConfig.metricCardSpan || 3);
    const gridCols = Math.max(metricSpan, gridConfig.gridCols || metricSpan * gridConfig.slotCols || 12);
    const colSpan = Math.max(1, widget.colSpan || metricSpan);
    const maxColStart = Math.max(1, gridCols - colSpan + 1);
    const rawColStart = ((entry.col || 1) - 1) * metricSpan + 1;
    const colStart = clamp(Math.round(rawColStart), 1, maxColStart);
    const rowStart = Math.max(1, Math.round(entry.row || 1));
    return `--widget-col-start:${colStart};--widget-row-start:${rowStart};--widget-col-span:${colSpan};--widget-row-span:${Math.max(1, widget.rowSpan || 1)};`;
  }

  function ensureValidCardLayout() {
    const availableKeys = getAvailableWidgetKeys();
    const availableSet = new Set(availableKeys);
    const gridConfig = getLayoutGridConfig();
    const seen = new Set();
    let changed = false;
    const normalized = [];
    const occupied = [];

    state.cardLayout.forEach((rawEntry) => {
      const entry = normalizeLayoutEntry(rawEntry);
      if (!entry) {
        changed = true;
        return;
      }

      if (!availableSet.has(entry.key) || seen.has(entry.key)) {
        changed = true;
        return;
      }
      seen.add(entry.key);

      let col = Number.isFinite(entry.col) ? Math.max(1, Math.round(entry.col)) : null;
      let row = Number.isFinite(entry.row) ? Math.max(1, Math.round(entry.row)) : null;
      if (!Number.isFinite(col) || !Number.isFinite(row)) {
        changed = true;
        const fallback = findFirstAvailableSlot(entry.key, 1, 1, gridConfig.slotCols, occupied, gridConfig.metricCardSpan);
        col = fallback.col;
        row = fallback.row;
      }

      if (!canPlaceWidgetAtSlot(entry.key, col, row, gridConfig.slotCols, occupied, gridConfig.metricCardSpan)) {
        changed = true;
        const fallback = findFirstAvailableSlot(entry.key, col, row, gridConfig.slotCols, occupied, gridConfig.metricCardSpan);
        col = fallback.col;
        row = fallback.row;
      }

      const size = getWidgetSlotSizeFromSpan(entry.key, gridConfig.metricCardSpan);
      occupied.push({
        key: entry.key,
        col,
        row,
        colSlots: size.colSlots,
        rowSlots: size.rowSlots,
      });
      normalized.push({ key: entry.key, col, row });
    });

    state.cardLayout = normalized;

    if (!state.cardLayout.length && !state.hasCustomLayout) {
      const defaults = DEFAULT_CARD_KEYS.filter((key) => isWidgetAvailable(key));
      state.cardLayout = buildDefaultCardLayout(defaults.length ? defaults : availableKeys.slice(0, 6), gridConfig.slotCols, gridConfig.metricCardSpan);
      return;
    }

    if (changed) {
      saveCardLayout();
    }
  }

  function renderWidgetBoard(rows) {
    if (!dom.metricsBoard) {
      return;
    }

    const previousRects = captureMetricCardRects();
    if (state.chart) {
      state.chart.destroy();
      state.chart = null;
    }

    const latest = rows.length ? rows[rows.length - 1] : null;
    const previous = rows.length > 1 ? rows[rows.length - 2] : null;
    const gridConfig = getLayoutGridConfig();
    const cards = sortLayoutEntries(state.cardLayout.filter((entry) => entry && isWidgetAvailable(entry.key)));

    if (!cards.length) {
      dom.metricsBoard.innerHTML = '<p class="metric-empty">No cells on board. Open Edit Cells and drag metrics in.</p>';
      hideDropIndicator();
      refreshDynamicDomRefs();
      state.chartVisible = false;
      stopAutoCycle();
      return;
    }

    dom.metricsBoard.innerHTML = cards
      .map((entry) => {
        const widget = getWidgetDefinition(entry.key);
        if (!widget) {
          return "";
        }
        const placementStyle = buildWidgetPlacementStyle(entry, widget, gridConfig);

        if (widget.type === "fixed" && widget.view === "goal") {
          return renderGoalWidget(widget, placementStyle);
        }
        if (widget.type === "fixed" && widget.view === "chart") {
          return renderChartWidget(widget, placementStyle);
        }
        if (widget.type === "fixed" && widget.view === "clock") {
          return renderClockWidget(widget, placementStyle);
        }
        if (widget.type === "advanced") {
          return renderAdvancedWidget(widget, placementStyle);
        }

        const meta = widget.metricMeta;
        if (!meta) {
          return "";
        }

        const summarySnapshot = getSummaryMetricSnapshot(meta);
        const value = summarySnapshot && Number.isFinite(summarySnapshot.value)
          ? summarySnapshot.value
          : latest
            ? parseNumber(latest[meta.valueCol])
            : NaN;
        const change = summarySnapshot && Number.isFinite(summarySnapshot.change)
          ? summarySnapshot.change
          : latest && meta.changeCol
            ? parseNumber(latest[meta.changeCol])
            : NaN;
        const trend = summarySnapshot
          ? resolveTrend(change, "")
          : resolveTrend(change, latest && meta.directionCol ? latest[meta.directionCol] : "");
        const previousActual = summarySnapshot && Number.isFinite(summarySnapshot.previousValue)
          ? summarySnapshot.previousValue
          : getPreviousMetricValue(previous, meta.valueCol);
        const previousEstimated = estimatePreviousValue(value, change);
        const previousLabel = summarySnapshot ? "vs previous MTD" : "vs previous";
        const previousText = Number.isFinite(previousActual)
          ? `${previousLabel}: ${formatMetricValue(meta, previousActual)}`
          : Number.isFinite(previousEstimated)
            ? `${previousLabel} (estimated): ${formatMetricValue(meta, previousEstimated)}`
            : `${previousLabel}: --`;
        const sign = change > 0 ? "+" : "";
        const changeText = Number.isFinite(change) ? `${trendArrow(trend)} ${sign}${formatNumber(change, 0)}%` : "--";
        const changeClass = Number.isFinite(change) ? trendClass(trend) : "";

        return `
          <article
            class="layout-widget metric-detail ${escapeHtml(meta.accentClass || "metric-generic")}"
            draggable="true"
            data-widget-id="${escapeHtml(widget.key)}"
            style="${placementStyle}"
          >
            <div class="widget-head metric-head">
              <p class="metric-title">${escapeHtml(meta.label)}</p>
              <button type="button" class="metric-remove" data-remove-widget="${escapeHtml(widget.key)}">Remove</button>
            </div>
            <p class="metric-current">${Number.isFinite(value) ? formatMetricValue(meta, value) : "--"}</p>
            <p class="metric-change ${changeClass}">${changeText}</p>
            <p class="metric-context">${previousText}</p>
          </article>
        `;
      })
      .join("");

    refreshDynamicDomRefs();
    updateClockWidget();
    if (!hasChartWidget()) {
      state.chartVisible = false;
      stopAutoCycle();
    }
    syncChartWidgetUi();
    animateMetricCardFlow(previousRects);
    renderDropIndicator();
  }

  function renderGoalWidget(widget, placementStyle) {
    return `
      <section
        class="layout-widget goal-panel widget-fixed ${escapeHtml(widget.accentClass || "")}"
        draggable="true"
        data-widget-id="${escapeHtml(widget.key)}"
        style="${placementStyle}"
      >
        <div class="widget-head">
          <p class="widget-title">Target Progress</p>
          <div class="widget-head-actions">
            <button type="button" class="goal-target-action" data-set-monthly-target>Set Target</button>
            <button type="button" class="goal-target-action secondary" data-clear-monthly-target>Clear</button>
            <button type="button" class="metric-remove" data-remove-widget="${escapeHtml(widget.key)}">Remove</button>
          </div>
        </div>
        <div class="goal-progress-head">
          <div>
            <p class="goal-title">Sales Target</p>
            <p class="goal-subtitle">MTD total sales vs target and last month total</p>
          </div>
          <div class="goal-progress-right">
            <p id="sales-goal-status" class="goal-status">Status --</p>
            <p id="sales-goal-progress-label" class="goal-percent">--</p>
          </div>
        </div>
        <div class="goal-values">
          <p class="goal-value goal-value-current">
            <span class="goal-value-label">Current</span>
            <span id="sales-goal-current-value" class="goal-value-amount">--</span>
          </p>
          <p class="goal-value goal-value-target">
            <span class="goal-value-label">Target</span>
            <span id="sales-goal-target-value" class="goal-value-amount">--</span>
          </p>
          <p class="goal-value goal-value-prev">
            <span class="goal-value-label">Last Month</span>
            <span id="sales-goal-prev-value" class="goal-value-amount">--</span>
            <span id="sales-goal-prev-note" class="goal-value-note"></span>
          </p>
        </div>
        <div class="goal-bar-wrap">
          <div class="goal-bar-track">
            <div id="sales-goal-progress-fill" class="goal-bar-fill"></div>
          </div>
        </div>
        <div class="goal-milestones">
          <p id="sales-goal-milestone-prev">Pace checkpoint today: --</p>
          <p id="sales-goal-milestone-target">Last month vs target: --</p>
        </div>
        <div class="goal-gaps">
          <p id="sales-goal-gap-prev" class="goal-gap">Gap vs last month total: --</p>
          <p id="sales-goal-gap-target" class="goal-gap">Gap to target: --</p>
        </div>
      </section>
    `;
  }

  function renderChartWidget(widget, placementStyle) {
    const chartHiddenClass = state.chartVisible ? "" : "hidden";
    const chartBtnLabel = state.chartVisible ? "Hide Chart" : "Show Chart";
    return `
      <section
        class="layout-widget panel chart-panel widget-fixed ${escapeHtml(widget.accentClass || "")}"
        draggable="true"
        data-widget-id="${escapeHtml(widget.key)}"
        style="${placementStyle}"
      >
        <div class="widget-head">
          <p class="widget-title">Trend Chart</p>
          <div class="widget-head-actions">
            <button id="toggle-chart-btn" type="button">${chartBtnLabel}</button>
            <button type="button" class="metric-remove" data-remove-widget="${escapeHtml(widget.key)}">Remove</button>
          </div>
        </div>
        <div id="chart-content" class="chart-content ${chartHiddenClass}">
          <div class="panel-controls">
            <label>
              Metric
              <select id="metric-select">
                <option value="">Loading...</option>
              </select>
            </label>
            <label>
              Chart Style
              <select id="chart-type-select">
                <option value="line">Line</option>
                <option value="bar">Bar</option>
                <option value="area">Area</option>
                <option value="doughnut">Doughnut</option>
              </select>
            </label>
            <label class="toggle-wrap">
              <input id="cycle-toggle" type="checkbox" />
              Auto-cycle metrics
            </label>
            <p id="selected-metric-trend" class="selected-trend">Trend: --</p>
          </div>
          <div class="chart-wrap">
            <canvas id="main-chart"></canvas>
          </div>
        </div>
      </section>
    `;
  }

  function renderClockWidget(widget, placementStyle) {
    const clockParts = getClockDisplayParts(state.now);
    return `
      <article
        class="layout-widget metric-detail ${escapeHtml(widget.accentClass || "metric-generic")}"
        draggable="true"
        data-widget-id="${escapeHtml(widget.key)}"
        style="${placementStyle}"
      >
        <div class="widget-head metric-head">
          <p class="metric-title">Clock + Date/Day</p>
          <button type="button" class="metric-remove" data-remove-widget="${escapeHtml(widget.key)}">Remove</button>
        </div>
        <div id="clock-display" class="clock-display ${clockParts.isColonDim ? "is-colon-dim" : ""}">
          <p id="clock-time" class="clock-time" role="timer" aria-live="polite" aria-atomic="true" aria-label="Local time ${escapeHtml(clockParts.time)}">
            <span id="clock-hours" class="clock-digits">${escapeHtml(clockParts.hours)}</span>
            <span class="clock-separator">:</span>
            <span id="clock-minutes" class="clock-digits">${escapeHtml(clockParts.minutes)}</span>
            <span class="clock-separator">:</span>
            <span id="clock-seconds" class="clock-digits clock-digits-seconds">${escapeHtml(clockParts.seconds)}</span>
          </p>
          <span class="clock-badge">24H</span>
        </div>
        <p id="clock-date" class="clock-date">${escapeHtml(clockParts.dateDay)}</p>
      </article>
    `;
  }

  function renderAdvancedWidget(widget, placementStyle) {
    const content = buildAdvancedCellContent(widget.payloadKey, state.cellsPayload);
    const subtitleHtml = content.subtitle ? `<p class="advanced-cell-subtitle">${escapeHtml(content.subtitle)}</p>` : "";
    return `
      <article
        class="layout-widget advanced-cell-card advanced-widget"
        draggable="true"
        data-widget-id="${escapeHtml(widget.key)}"
        style="${placementStyle}"
      >
        <div class="widget-head advanced-widget-head">
          <div>
            <p class="advanced-cell-title">${escapeHtml(widget.label || "Advanced Cell")}</p>
            ${subtitleHtml}
          </div>
          <button type="button" class="metric-remove" data-remove-widget="${escapeHtml(widget.key)}">Remove</button>
        </div>
        ${content.body}
      </article>
    `;
  }

  function captureMetricCardRects() {
    const rects = new Map();
    if (!dom.metricsBoard) {
      return rects;
    }
    const cards = dom.metricsBoard.querySelectorAll(".layout-widget[data-widget-id]");
    cards.forEach((card) => {
      if (card.dataset.widgetId) {
        rects.set(card.dataset.widgetId, card.getBoundingClientRect());
      }
    });
    return rects;
  }

  function animateMetricCardFlow(previousRects) {
    if (!dom.metricsBoard || !previousRects || !previousRects.size) {
      return;
    }
    if (window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches) {
      return;
    }

    const cards = dom.metricsBoard.querySelectorAll(".layout-widget[data-widget-id]");
    cards.forEach((card) => {
      const key = card.dataset.widgetId;
      if (!key) {
        return;
      }
      const prev = previousRects.get(key);
      if (!prev) {
        return;
      }
      const next = card.getBoundingClientRect();
      const dx = prev.left - next.left;
      const dy = prev.top - next.top;
      if (Math.abs(dx) < 1 && Math.abs(dy) < 1) {
        return;
      }
      if (typeof card.animate !== "function") {
        return;
      }
      card.animate(
        [
          { transform: `translate(${dx}px, ${dy}px)` },
          { transform: "translate(0px, 0px)" },
        ],
        {
          duration: 260,
          easing: "cubic-bezier(0.2, 0.8, 0.2, 1)",
        }
      );
    });
  }

  function renderCellLibrary() {
    if (!dom.cellLibrary) {
      return;
    }

    const widgets = getAvailableWidgetDefinitions().sort((a, b) => a.label.localeCompare(b.label));
    if (!widgets.length) {
      dom.cellLibrary.innerHTML = '<p class="metric-empty">No cells found yet.</p>';
      return;
    }

    dom.cellLibrary.innerHTML = widgets
      .map((widget) => {
        const active = isWidgetOnBoard(widget.key);
        return `
          <button type="button" class="cell-item ${active ? "active" : ""}" draggable="true" data-cell-key="${escapeHtml(widget.key)}">
            <span>${escapeHtml(widget.label)}</span>
            <small>${active ? "On board" : "Drag to add"}</small>
          </button>
        `;
      })
      .join("");
  }

  function refreshDynamicDomRefs() {
    dom.metricSelect = document.getElementById("metric-select");
    dom.chartTypeSelect = document.getElementById("chart-type-select");
    dom.cycleToggle = document.getElementById("cycle-toggle");
    dom.toggleChartBtn = document.getElementById("toggle-chart-btn");
    dom.chartContent = document.getElementById("chart-content");
    dom.selectedMetricTrend = document.getElementById("selected-metric-trend");
    dom.chartCanvas = document.getElementById("main-chart");
    dom.clockDisplay = document.getElementById("clock-display");
    dom.clockTime = document.getElementById("clock-time");
    dom.clockHours = document.getElementById("clock-hours");
    dom.clockMinutes = document.getElementById("clock-minutes");
    dom.clockSeconds = document.getElementById("clock-seconds");
    dom.clockDate = document.getElementById("clock-date");

    dom.salesGoalStatus = document.getElementById("sales-goal-status");
    dom.salesGoalCurrentValue = document.getElementById("sales-goal-current-value");
    dom.salesGoalTargetValue = document.getElementById("sales-goal-target-value");
    dom.salesGoalPrevValue = document.getElementById("sales-goal-prev-value");
    dom.salesGoalPrevNote = document.getElementById("sales-goal-prev-note");
    dom.salesGoalMilestonePrev = document.getElementById("sales-goal-milestone-prev");
    dom.salesGoalMilestoneTarget = document.getElementById("sales-goal-milestone-target");
    dom.salesGoalProgressLabel = document.getElementById("sales-goal-progress-label");
    dom.salesGoalProgressFill = document.getElementById("sales-goal-progress-fill");
    dom.salesGoalGapPrev = document.getElementById("sales-goal-gap-prev");
    dom.salesGoalGapTarget = document.getElementById("sales-goal-gap-target");
  }

  function syncChartWidgetUi() {
    if (dom.chartTypeSelect) {
      dom.chartTypeSelect.value = state.chartType;
    }
    if (dom.cycleToggle) {
      dom.cycleToggle.checked = Boolean(state.cycleTimer);
    }
    if (dom.toggleChartBtn) {
      dom.toggleChartBtn.textContent = state.chartVisible ? "Hide Chart" : "Show Chart";
    }
    if (dom.chartContent) {
      dom.chartContent.classList.toggle("hidden", !state.chartVisible);
    }
  }

  function syncChartMetricOptions() {
    if (!dom.metricSelect) {
      if (state.chart) {
        state.chart.destroy();
        state.chart = null;
      }
      return;
    }

    const keys = getChartMetricKeys();
    if (!keys.length) {
      dom.metricSelect.innerHTML = '<option value="">No metrics</option>';
      dom.metricSelect.disabled = true;
      state.currentMetric = "";
      if (state.chart) {
        state.chart.destroy();
        state.chart = null;
      }
      return;
    }

    dom.metricSelect.disabled = false;
    dom.metricSelect.innerHTML = keys
      .map((key) => `<option value="${escapeHtml(key)}">${escapeHtml(state.metricCatalog[key].label)}</option>`)
      .join("");

    if (!keys.includes(state.currentMetric)) {
      state.currentMetric = keys[0];
    }
    dom.metricSelect.value = state.currentMetric;
  }

  function getChartMetricKeys() {
    const fromBoard = state.cardLayout
      .map((entry) => (entry && entry.key ? entry.key : ""))
      .filter((key) => isMetricWidgetKey(key))
      .map((key) => metricKeyFromWidgetKey(key))
      .filter((key) => state.metricCatalog[key]);
    if (fromBoard.length) {
      return fromBoard;
    }
    return Object.keys(state.metricCatalog);
  }

  function getSalesGoalMetric() {
    if (state.metricCatalog.sales) {
      return state.metricCatalog.sales;
    }
    return (
      Object.values(state.metricCatalog).find((meta) => meta.valueCol === "Order Revenue (Current)") ||
      getChartMetricKeys().map((key) => state.metricCatalog[key])[0] ||
      null
    );
  }

  function getSummaryMetricSnapshot(meta) {
    if (!meta || !meta.key || !state.cellsPayload || !state.cellsPayload.kpis) {
      return null;
    }

    const current = state.cellsPayload.kpis.current || {};
    const previous = state.cellsPayload.kpis.previous || {};
    const change = state.cellsPayload.kpis.change || {};

    if (meta.key === "sales") {
      return {
        value: parseNumber(current.sales_amount),
        previousValue: parseNumber(previous.sales_amount),
        change: parseNumber(change.sales_amount_pct),
      };
    }

    if (meta.key === "orders") {
      return {
        value: parseNumber(current.orders),
        previousValue: parseNumber(previous.orders),
        change: parseNumber(change.orders_pct),
      };
    }

    if (meta.key === "aov") {
      return {
        value: parseNumber(current.aov),
        previousValue: parseNumber(previous.aov),
        change: parseNumber(change.aov_pct),
      };
    }

    return null;
  }

  function formatMetricValue(meta, value) {
    if (!meta || !Number.isFinite(value)) {
      return "--";
    }
    if (meta.formatType === "currency") {
      return formatCurrency(value, cfg.currencySymbol);
    }
    if (meta.formatType === "ratio") {
      return `${formatNumber(value, 2)}x`;
    }
    if (meta.formatType === "percent") {
      return `${formatNumber(value, 0)}%`;
    }
    if (meta.formatType === "count") {
      return formatNumber(value, 0);
    }
    return formatNumber(value, 2);
  }

  function updateSalesGoal(row, previousRow, salesMeta) {
    if (
      !dom.salesGoalCurrentValue ||
      !dom.salesGoalTargetValue ||
      !dom.salesGoalPrevValue ||
      !dom.salesGoalPrevNote ||
      !dom.salesGoalMilestonePrev ||
      !dom.salesGoalMilestoneTarget ||
      !dom.salesGoalStatus ||
      !dom.salesGoalProgressLabel ||
      !dom.salesGoalProgressFill ||
      !dom.salesGoalGapPrev ||
      !dom.salesGoalGapTarget
    ) {
      return;
    }
    if (!salesMeta) {
      clearSalesGoal();
      return;
    }

    const backendPaceMtdTotalSales = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.daily_sales_pace &&
        state.cellsPayload.daily_sales_pace.mtd_sales
    );
    const backendProjectedMtdTotalSales = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.mtd_projection &&
        state.cellsPayload.mtd_projection.mtd_sales
    );
    const backendFallbackMtdGross = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.mtd_projection &&
        state.cellsPayload.mtd_projection.mtd_gross_sales
    );
    const current = Number.isFinite(backendPaceMtdTotalSales)
      ? backendPaceMtdTotalSales
      : Number.isFinite(backendProjectedMtdTotalSales)
        ? backendProjectedMtdTotalSales
      : Number.isFinite(backendFallbackMtdGross)
        ? backendFallbackMtdGross
        : parseNumber(row[salesMeta.valueCol]);
    const change = parseNumber(row[salesMeta.changeCol]);
    const backendPrevMonthTotal = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.daily_sales_pace &&
        state.cellsPayload.daily_sales_pace.previous_month_total_sales
    );
    const backendPrevMonthGross = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.daily_sales_pace &&
        state.cellsPayload.daily_sales_pace.previous_month_gross_sales
    );
    const backendPrevMonthNet = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.daily_sales_pace &&
        state.cellsPayload.daily_sales_pace.previous_month_net_sales
    );
    const backendPrevComparableSales = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.kpis &&
        state.cellsPayload.kpis.previous &&
        state.cellsPayload.kpis.previous.sales_amount
    );
    const previousFromRows = getPreviousMetricValue(previousRow, salesMeta.valueCol);
    const previousActual = Number.isFinite(backendPrevMonthTotal)
      ? backendPrevMonthTotal
      : Number.isFinite(backendPrevMonthGross)
        ? backendPrevMonthGross
      : Number.isFinite(backendPrevMonthNet)
        ? backendPrevMonthNet
      : Number.isFinite(backendPrevComparableSales)
        ? backendPrevComparableSales
        : previousFromRows;
    const previousEstimated = Number.isFinite(previousActual) ? NaN : estimatePreviousValue(current, change);
    const previous = Number.isFinite(previousActual) ? previousActual : previousEstimated;
    const previousSource = Number.isFinite(backendPrevMonthTotal)
      ? "shopify_last_month_total"
      : Number.isFinite(backendPrevMonthGross)
        ? "shopify_last_month_gross"
      : Number.isFinite(backendPrevMonthNet)
        ? "shopify_last_month_net"
      : Number.isFinite(backendPrevComparableSales)
        ? "shopify_previous_mtd"
      : Number.isFinite(previousFromRows)
        ? "prior_point"
        : "estimated";
    const monthKey = getMonthKey(new Date());

    const multiplier = Number.isFinite(cfg.salesTargetMultiplier) && cfg.salesTargetMultiplier > 0 ? cfg.salesTargetMultiplier : 1;
    const monthlyStoredTarget = getMonthlyTargetForKey(monthKey);
    const monthlyConfigTarget = getConfigTargetForKey(monthKey);
    const overrideTarget = parseNumber(cfg.salesTargetValue);
    const computedTarget = Number.isFinite(previous) ? previous * multiplier : NaN;
    const target = Number.isFinite(monthlyStoredTarget)
      ? monthlyStoredTarget
      : Number.isFinite(monthlyConfigTarget)
        ? monthlyConfigTarget
        : Number.isFinite(overrideTarget)
          ? overrideTarget
          : computedTarget;
    if (!Number.isFinite(current) || !Number.isFinite(target) || target <= 0) {
      showSalesGoalNeedsTarget(current, previous);
      return;
    }

    const progressPct = (current / target) * 100;
    const progressClamped = clampPercent(progressPct);
    const delta = current - target;
    const toTarget = target - current;
    const beatTarget = current >= target;
    const deltaVsLastMonth = Number.isFinite(previous) ? current - previous : NaN;
    const growthVsLastMonthPct = Number.isFinite(previous) && previous !== 0 ? ((current - previous) / Math.abs(previous)) * 100 : NaN;

    dom.salesGoalCurrentValue.textContent = formatCurrency(current, cfg.currencySymbol);
    dom.salesGoalTargetValue.textContent = formatCurrency(target, cfg.currencySymbol);
    dom.salesGoalPrevValue.textContent = Number.isFinite(previous) ? formatCurrency(previous, cfg.currencySymbol) : "--";
    dom.salesGoalPrevNote.textContent =
      previousSource === "shopify_last_month_total"
        ? "shopify last month total"
        : previousSource === "shopify_last_month_gross"
        ? "shopify last month gross"
        : previousSource === "shopify_last_month_net"
          ? "shopify last month net"
        : previousSource === "shopify_previous_mtd"
        ? "shopify prev MTD"
        : previousSource === "estimated"
          ? "estimated"
          : "";

    const paceDaysElapsed = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.daily_sales_pace &&
        state.cellsPayload.daily_sales_pace.days_elapsed
    );
    const paceDaysRemaining = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.daily_sales_pace &&
        state.cellsPayload.daily_sales_pace.days_remaining
    );
    const projectionDaysInMonth = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.mtd_projection &&
        state.cellsPayload.mtd_projection.days_in_month
    );
    const projectionDaysElapsed = parseNumber(
      state.cellsPayload &&
        state.cellsPayload.mtd_projection &&
        state.cellsPayload.mtd_projection.days_elapsed
    );
    const totalDays = Number.isFinite(paceDaysElapsed) && Number.isFinite(paceDaysRemaining)
      ? paceDaysElapsed + paceDaysRemaining
      : Number.isFinite(projectionDaysInMonth)
        ? projectionDaysInMonth
        : NaN;
    const elapsedDays = Number.isFinite(paceDaysElapsed)
      ? paceDaysElapsed
      : Number.isFinite(projectionDaysElapsed)
        ? projectionDaysElapsed
        : NaN;
    const expectedProgressPct =
      Number.isFinite(totalDays) && totalDays > 0 && Number.isFinite(elapsedDays)
        ? (elapsedDays / totalDays) * 100
        : NaN;

    dom.salesGoalMilestonePrev.textContent = Number.isFinite(expectedProgressPct) && Number.isFinite(elapsedDays) && Number.isFinite(totalDays)
      ? `Pace checkpoint today: ${formatNumber(expectedProgressPct, 1)}% (${formatNumber(elapsedDays, 0)}/${formatNumber(totalDays, 0)} days)`
      : "Pace checkpoint today: --";
    dom.salesGoalMilestoneTarget.textContent = Number.isFinite(previous) && target > 0
      ? `Last month vs target: ${formatNumber((previous / target) * 100, 1)}%`
      : `Last month vs target: --`;
    dom.salesGoalProgressLabel.textContent = `${formatNumber(progressPct, 1)}%`;
    dom.salesGoalProgressFill.style.width = `${progressClamped}%`;
    dom.salesGoalProgressFill.classList.remove("below-prev", "on-track", "complete");
    const onPace = Number.isFinite(expectedProgressPct) ? progressPct >= expectedProgressPct : current >= (target * 0.5);
    if (beatTarget) {
      dom.salesGoalProgressFill.classList.add("complete");
    } else if (onPace) {
      dom.salesGoalProgressFill.classList.add("on-track");
    } else {
      dom.salesGoalProgressFill.classList.add("below-prev");
    }

    if (beatTarget) {
      dom.salesGoalStatus.textContent = "Target achieved";
      dom.salesGoalStatus.className = "goal-status up";
      dom.salesGoalGapPrev.textContent = Number.isFinite(deltaVsLastMonth)
        ? `Gap vs last month total: +${formatCurrency(Math.abs(deltaVsLastMonth), cfg.currencySymbol)} (${formatPercentSafe(growthVsLastMonthPct)})`
        : "Gap vs last month total: --";
      dom.salesGoalGapPrev.className = "goal-gap up";
      dom.salesGoalGapTarget.textContent = `Gap to target: +${formatCurrency(Math.abs(delta), cfg.currencySymbol)}`;
      dom.salesGoalGapTarget.className = "goal-gap up";
      return;
    }

    if (!onPace) {
      dom.salesGoalStatus.textContent = "Behind pace";
      dom.salesGoalStatus.className = "goal-status down";
    } else {
      dom.salesGoalStatus.textContent = "On pace";
      dom.salesGoalStatus.className = "goal-status mid";
    }
    dom.salesGoalGapPrev.textContent = Number.isFinite(deltaVsLastMonth)
      ? `${deltaVsLastMonth >= 0 ? "Gap vs last month total: +" : "Gap vs last month total: -"}${formatCurrency(Math.abs(deltaVsLastMonth), cfg.currencySymbol)} (${formatPercentSafe(growthVsLastMonthPct)})`
      : "Gap vs last month total: --";
    dom.salesGoalGapPrev.className = deltaVsLastMonth >= 0 ? "goal-gap mid" : "goal-gap down";
    dom.salesGoalGapTarget.textContent = `Gap to target: -${formatCurrency(Math.abs(toTarget), cfg.currencySymbol)}`;
    dom.salesGoalGapTarget.className = onPace ? "goal-gap mid" : "goal-gap down";
  }

  function clearSalesGoal() {
    if (
      !dom.salesGoalCurrentValue ||
      !dom.salesGoalTargetValue ||
      !dom.salesGoalPrevValue ||
      !dom.salesGoalPrevNote ||
      !dom.salesGoalMilestonePrev ||
      !dom.salesGoalMilestoneTarget ||
      !dom.salesGoalStatus ||
      !dom.salesGoalProgressLabel ||
      !dom.salesGoalProgressFill ||
      !dom.salesGoalGapPrev ||
      !dom.salesGoalGapTarget
    ) {
      return;
    }
    dom.salesGoalCurrentValue.textContent = "--";
    dom.salesGoalTargetValue.textContent = "--";
    dom.salesGoalPrevValue.textContent = "--";
    dom.salesGoalPrevNote.textContent = "";
    dom.salesGoalStatus.textContent = "Status --";
    dom.salesGoalStatus.className = "goal-status";
    dom.salesGoalMilestonePrev.textContent = "Pace checkpoint today: --";
    dom.salesGoalMilestoneTarget.textContent = "Last month vs target: --";
    dom.salesGoalProgressLabel.textContent = "--";
    dom.salesGoalProgressFill.style.width = "0%";
    dom.salesGoalProgressFill.classList.remove("below-prev", "on-track", "complete");
    dom.salesGoalGapPrev.textContent = "Gap vs last month total: --";
    dom.salesGoalGapPrev.className = "goal-gap";
    dom.salesGoalGapTarget.textContent = "Gap to target: --";
    dom.salesGoalGapTarget.className = "goal-gap";
  }

  function showSalesGoalNeedsTarget(current, previous) {
    if (
      !dom.salesGoalCurrentValue ||
      !dom.salesGoalTargetValue ||
      !dom.salesGoalPrevValue ||
      !dom.salesGoalPrevNote ||
      !dom.salesGoalMilestonePrev ||
      !dom.salesGoalMilestoneTarget ||
      !dom.salesGoalStatus ||
      !dom.salesGoalProgressLabel ||
      !dom.salesGoalProgressFill ||
      !dom.salesGoalGapPrev ||
      !dom.salesGoalGapTarget
    ) {
      return;
    }

    dom.salesGoalCurrentValue.textContent = Number.isFinite(current) ? formatCurrency(current, cfg.currencySymbol) : "--";
    dom.salesGoalTargetValue.textContent = "--";
    dom.salesGoalPrevValue.textContent = Number.isFinite(previous) ? formatCurrency(previous, cfg.currencySymbol) : "--";
    dom.salesGoalPrevNote.textContent = "";
    dom.salesGoalStatus.textContent = "Set monthly target";
    dom.salesGoalStatus.className = "goal-status mid";
    dom.salesGoalMilestonePrev.textContent = "Pace checkpoint today: --";
    dom.salesGoalMilestoneTarget.textContent = "Last month vs target: --";
    dom.salesGoalProgressLabel.textContent = "--";
    dom.salesGoalProgressFill.style.width = "0%";
    dom.salesGoalProgressFill.classList.remove("below-prev", "on-track", "complete");
    dom.salesGoalGapPrev.textContent = "Gap vs last month total: --";
    dom.salesGoalGapPrev.className = "goal-gap";
    dom.salesGoalGapTarget.textContent = "Gap to target: set target";
    dom.salesGoalGapTarget.className = "goal-gap mid";
  }

  function onSetMonthlyTarget() {
    const monthKey = getCurrentGoalMonthKey();
    const monthLabel = formatMonthLabel(monthKey);
    const existing = getMonthlyTargetForKey(monthKey) || getConfigTargetForKey(monthKey) || parseNumber(cfg.salesTargetValue) || "";
    const response = window.prompt(
      `Set monthly sales target for ${monthLabel} (${cfg.currencySymbol}).\nEnter numbers only, e.g. 120000`,
      existing ? String(existing) : ""
    );
    if (response === null) {
      return;
    }

    const parsed = parseNumber(response);
    if (!Number.isFinite(parsed) || parsed <= 0) {
      window.alert("Please enter a valid positive number.");
      return;
    }

    state.monthlyTargets[monthKey] = roundTo(parsed, 2);
    saveMonthlyTargets();
    updateKpis(state.rows);
    if (state.chartVisible) {
      renderChart();
    }
  }

  function onClearMonthlyTarget() {
    const monthKey = getCurrentGoalMonthKey();
    if (Object.prototype.hasOwnProperty.call(state.monthlyTargets, monthKey)) {
      delete state.monthlyTargets[monthKey];
      saveMonthlyTargets();
      updateKpis(state.rows);
      if (state.chartVisible) {
        renderChart();
      }
      return;
    }
    window.alert(`No saved monthly target found for ${formatMonthLabel(monthKey)}.`);
  }

  function getCurrentGoalMonthKey() {
    return getMonthKey(new Date());
  }

  function getMonthKey(date) {
    const safeDate = date instanceof Date && Number.isFinite(date.getTime()) ? date : new Date();
    const year = safeDate.getFullYear();
    const month = String(safeDate.getMonth() + 1).padStart(2, "0");
    return `${year}-${month}`;
  }

  function formatMonthLabel(monthKey) {
    if (typeof monthKey !== "string" || !/^\d{4}-\d{2}$/.test(monthKey)) {
      return "this month";
    }
    const [yearText, monthText] = monthKey.split("-");
    const year = Number.parseInt(yearText, 10);
    const month = Number.parseInt(monthText, 10);
    if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) {
      return monthKey;
    }
    return new Date(year, month - 1, 1).toLocaleString("en-GB", {
      month: "long",
      year: "numeric",
    });
  }

  function loadMonthlyTargets() {
    try {
      const raw = window.localStorage.getItem(MONTHLY_TARGETS_STORAGE_KEY);
      if (!raw) {
        return {};
      }
      const parsed = JSON.parse(raw);
      if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
        return {};
      }

      const result = {};
      Object.entries(parsed).forEach(([key, value]) => {
        if (!/^\d{4}-\d{2}$/.test(key)) {
          return;
        }
        const parsedValue = parseNumber(value);
        if (!Number.isFinite(parsedValue) || parsedValue <= 0) {
          return;
        }
        result[key] = roundTo(parsedValue, 2);
      });
      return result;
    } catch (_error) {
      return {};
    }
  }

  function saveMonthlyTargets() {
    try {
      window.localStorage.setItem(MONTHLY_TARGETS_STORAGE_KEY, JSON.stringify(state.monthlyTargets));
    } catch (_error) {
      // Ignore storage write failures.
    }
  }

  function getMonthlyTargetForKey(monthKey) {
    if (!monthKey || !state.monthlyTargets || typeof state.monthlyTargets !== "object") {
      return NaN;
    }
    return parseNumber(state.monthlyTargets[monthKey]);
  }

  function getConfigTargetForKey(monthKey) {
    if (!monthKey || !cfg.salesTargetsByMonth || typeof cfg.salesTargetsByMonth !== "object") {
      return NaN;
    }
    return parseNumber(cfg.salesTargetsByMonth[monthKey]);
  }

  function toggleCellMenu(forceState) {
    if (!dom.cellMenu || !dom.cellMenuToggle) {
      return;
    }
    const nextOpen = typeof forceState === "boolean" ? forceState : dom.cellMenu.classList.contains("hidden");
    dom.cellMenu.classList.toggle("hidden", !nextOpen);
    dom.cellMenuToggle.classList.toggle("is-open", nextOpen);
  }

  function loadCardLayout() {
    try {
      const raw = window.localStorage.getItem(LAYOUT_STORAGE_KEY) || window.localStorage.getItem("elave_dash_metric_layout_v1");
      if (!raw) {
        return null;
      }
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        return null;
      }
      return parsed
        .map((item) => {
          if (typeof item === "string") {
            return normalizeWidgetKey(item);
          }
          if (!item || typeof item !== "object") {
            return null;
          }
          const key = normalizeWidgetKey(item.key);
          if (!key) {
            return null;
          }
          return {
            key,
            col: Number.parseInt(item.col, 10),
            row: Number.parseInt(item.row, 10),
          };
        })
        .filter(Boolean);
    } catch (_error) {
      return null;
    }
  }

  function saveCardLayout() {
    try {
      window.localStorage.setItem(LAYOUT_STORAGE_KEY, JSON.stringify(state.cardLayout));
    } catch (_error) {
      // Ignore storage write failures.
    }
  }

  function onLibraryClickAdd(event) {
    const item = event.target.closest("[data-cell-key]");
    if (!item) {
      return;
    }
    addOrMoveCard(item.dataset.cellKey, null);
  }

  function onMetricsBoardClick(event) {
    const toggleBtn = event.target.closest("#toggle-chart-btn");
    if (toggleBtn) {
      onToggleChartVisibility();
      return;
    }

    const setTargetBtn = event.target.closest("[data-set-monthly-target]");
    if (setTargetBtn) {
      onSetMonthlyTarget();
      return;
    }

    const clearTargetBtn = event.target.closest("[data-clear-monthly-target]");
    if (clearTargetBtn) {
      onClearMonthlyTarget();
      return;
    }

    const removeBtn = event.target.closest("[data-remove-widget]");
    if (!removeBtn || !removeBtn.dataset.removeWidget) {
      return;
    }
    const key = removeBtn.dataset.removeWidget;
    state.cardLayout = state.cardLayout.filter((entry) => entry.key !== key);
    state.hasCustomLayout = true;
    if (key === "chart_panel") {
      state.chartVisible = false;
      stopAutoCycle();
      if (state.chart) {
        state.chart.destroy();
        state.chart = null;
      }
    }
    saveCardLayout();
    renderWidgetBoard(state.rows);
    syncChartMetricOptions();
    if (dom.cycleToggle && dom.cycleToggle.checked) {
      startAutoCycle();
    }
    renderCellLibrary();
    updateSelectedMetricTrend(state.rows[state.rows.length - 1] || null);
    if (state.chartVisible) {
      renderChart();
    }
  }

  function onMetricsBoardChange(event) {
    if (!event || !event.target) {
      return;
    }
    const target = event.target;
    if (target.id === "metric-select") {
      state.currentMetric = target.value;
      renderChart();
      updateSelectedMetricTrend(state.rows[state.rows.length - 1] || null);
      return;
    }
    if (target.id === "chart-type-select") {
      state.chartType = target.value;
      renderChart();
      return;
    }
    if (target.id === "cycle-toggle") {
      if (target.checked) {
        startAutoCycle();
        return;
      }
      stopAutoCycle();
    }
  }

  function onLibraryDragStart(event) {
    const cell = event.target.closest("[data-cell-key]");
    if (!cell || !cell.dataset.cellKey) {
      return;
    }
    state.activeDragMetric = cell.dataset.cellKey;
    state.activeDragSource = "library";
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "copyMove";
      event.dataTransfer.setData("text/plain", state.activeDragMetric);
    }
  }

  function onMetricCardDragStart(event) {
    const card = event.target.closest("[data-widget-id]");
    if (!card || !card.dataset.widgetId) {
      return;
    }
    state.activeDragMetric = card.dataset.widgetId;
    state.activeDragSource = "board";
    state.activeDragCardEl = card;
    card.classList.add("dragging");
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", state.activeDragMetric);
    }
  }

  function onMetricsBoardDragOver(event) {
    if (!state.activeDragMetric || !dom.metricsBoard) {
      return;
    }
    event.preventDefault();
    if (event.dataTransfer) {
      event.dataTransfer.dropEffect = state.activeDragSource === "library" ? "copy" : "move";
    }
    dom.metricsBoard.classList.add("drag-target");
    const pointerSlot = getDropSlot(event);
    state.dragDropSlot = pointerSlot ? resolvePreferredDropSlot(state.activeDragMetric, pointerSlot) : null;
    renderDropIndicator();
  }

  function onMetricsBoardDrop(event) {
    if (!dom.metricsBoard) {
      return;
    }
    event.preventDefault();
    dom.metricsBoard.classList.remove("drag-target");
    if (!state.activeDragMetric) {
      return;
    }
    const dropSlot = state.dragDropSlot || getDropSlot(event);
    addOrMoveCard(state.activeDragMetric, dropSlot);
    onAnyDragEnd();
  }

  function onMetricsBoardDragLeave(event) {
    if (!dom.metricsBoard || dom.metricsBoard.contains(event.relatedTarget)) {
      return;
    }
    dom.metricsBoard.classList.remove("drag-target");
    hideDropIndicator();
  }

  function onAnyDragEnd() {
    if (dom.metricsBoard) {
      dom.metricsBoard.classList.remove("drag-target");
    }
    if (state.activeDragCardEl) {
      state.activeDragCardEl.classList.remove("dragging");
    }
    state.activeDragMetric = "";
    state.activeDragSource = "";
    state.activeDragCardEl = null;
    state.dragDropSlot = null;
    hideDropIndicator();
  }

  function getDropSlot(event) {
    if (!dom.metricsBoard) {
      return null;
    }
    const geometry = getMetricGridGeometry();
    if (!geometry) {
      return null;
    }

    const x = clamp(event.clientX - geometry.rect.left, 0, Math.max(0, geometry.rect.width - 1));
    const y = clamp(event.clientY - geometry.rect.top, 0, Math.max(0, geometry.rect.height + geometry.rowStep));
    const widgetSize = getWidgetSlotSize(state.activeDragMetric, geometry);
    const maxColStart = Math.max(1, geometry.slotCols - widgetSize.colSlots + 1);
    const col = clamp(Math.floor(x / geometry.slotWidth) + 1, 1, maxColStart);
    const row = Math.max(1, Math.floor(y / geometry.rowStep) + 1);
    return { col, row };
  }

  function resolvePreferredDropSlot(widgetKey, desiredSlot) {
    const gridConfig = getLayoutGridConfig();
    const remaining = state.cardLayout.filter((entry) => entry.key !== widgetKey);
    const occupied = buildOccupiedRects(remaining, gridConfig.slotCols, gridConfig.metricCardSpan);
    return findFirstAvailableSlot(
      widgetKey,
      desiredSlot.col,
      desiredSlot.row,
      gridConfig.slotCols,
      occupied,
      gridConfig.metricCardSpan
    );
  }

  function renderDropIndicator() {
    if (!dom.metricsBoard || !state.dragDropSlot || !state.activeDragMetric) {
      hideDropIndicator();
      return;
    }
    const geometry = getMetricGridGeometry();
    if (!geometry) {
      hideDropIndicator();
      return;
    }

    const col = state.dragDropSlot.col;
    const row = state.dragDropSlot.row;
    const widgetSize = getWidgetSlotSize(state.activeDragMetric, geometry);
    const x = (col - 1) * geometry.slotWidth;
    const y = (row - 1) * geometry.rowStep;
    const width = Math.max(72, geometry.slotWidth * widgetSize.colSlots - geometry.gap);
    const height = Math.max(72, geometry.rowHeight * widgetSize.rowSlots + geometry.gap * (widgetSize.rowSlots - 1));

    const indicator = ensureDropIndicatorEl();
    indicator.style.width = `${width}px`;
    indicator.style.height = `${height}px`;
    indicator.style.transform = `translate(${x}px, ${y}px)`;
    indicator.classList.add("is-visible");
  }

  function hideDropIndicator() {
    if (!dom.metricsBoard) {
      return;
    }
    const indicator = dom.metricsBoard.querySelector(".metrics-drop-indicator");
    if (indicator) {
      indicator.classList.remove("is-visible");
    }
  }

  function ensureDropIndicatorEl() {
    let indicator = dom.metricsBoard.querySelector(".metrics-drop-indicator");
    if (indicator) {
      return indicator;
    }
    indicator = document.createElement("div");
    indicator.className = "metrics-drop-indicator";
    dom.metricsBoard.appendChild(indicator);
    return indicator;
  }

  function getMetricGridGeometry() {
    if (!dom.metricsBoard) {
      return null;
    }
    const rect = dom.metricsBoard.getBoundingClientRect();
    if (!rect.width) {
      return null;
    }

    const style = window.getComputedStyle(dom.metricsBoard);
    const gridCols = Number.parseInt(style.getPropertyValue("--metric-grid-cols"), 10) || 12;
    const cardSpan = Number.parseInt(style.getPropertyValue("--metric-card-span"), 10) || 3;
    const rowHeight = Number.parseFloat(style.getPropertyValue("--metric-row-height")) || 198;
    const gap = Number.parseFloat(style.rowGap || "0") || 0;
    const slotCols = Math.max(1, Math.floor(gridCols / Math.max(1, cardSpan)));
    const slotWidth = rect.width / slotCols;
    const rowStep = rowHeight + gap;

    return {
      rect,
      gridCols,
      slotCols,
      slotWidth,
      rowHeight,
      rowStep,
      gap,
      metricCardSpan: cardSpan,
    };
  }

  function getWidgetSlotSize(widgetKey, geometry) {
    const def = getWidgetDefinition(widgetKey);
    if (!def) {
      return { colSlots: 1, rowSlots: 1 };
    }
    const metricSpan = Math.max(1, geometry.metricCardSpan || 3);
    const colSlots = Math.max(1, Math.round((def.colSpan || metricSpan) / metricSpan));
    const rowSlots = Math.max(1, def.rowSpan || 1);
    return { colSlots, rowSlots };
  }

  function addOrMoveCard(widgetKey, targetSlot) {
    if (!widgetKey || !isWidgetAvailable(widgetKey)) {
      return;
    }

    const normalizedKey = normalizeWidgetKey(widgetKey);
    const gridConfig = getLayoutGridConfig();
    const existingEntry = state.cardLayout.find((entry) => entry.key === normalizedKey);
    if (existingEntry && !targetSlot) {
      return;
    }

    const remaining = state.cardLayout.filter((entry) => entry.key !== normalizedKey);
    const occupied = buildOccupiedRects(remaining, gridConfig.slotCols, gridConfig.metricCardSpan);
    const desiredCol = targetSlot && Number.isFinite(targetSlot.col) ? targetSlot.col : 1;
    const desiredRow = targetSlot && Number.isFinite(targetSlot.row) ? targetSlot.row : 1;
    const nextSlot = findFirstAvailableSlot(
      normalizedKey,
      desiredCol,
      desiredRow,
      gridConfig.slotCols,
      occupied,
      gridConfig.metricCardSpan
    );

    state.cardLayout = [...remaining, { key: normalizedKey, col: nextSlot.col, row: nextSlot.row }];
    state.hasCustomLayout = true;
    saveCardLayout();
    renderWidgetBoard(state.rows);
    syncChartMetricOptions();
    if (dom.cycleToggle && dom.cycleToggle.checked) {
      startAutoCycle();
    }
    renderCellLibrary();
    updateSelectedMetricTrend(state.rows[state.rows.length - 1] || null);
    if (state.chartVisible) {
      renderChart();
    }
  }

  function renderChart() {
    if (!state.chartVisible || !dom.chartContent || dom.chartContent.classList.contains("hidden") || !dom.chartCanvas) {
      return;
    }
    if (typeof window.Chart === "undefined") {
      return;
    }
    const rows = state.rows;
    if (!rows.length) {
      if (state.chart) {
        state.chart.destroy();
        state.chart = null;
      }
      return;
    }

    const metric = state.metricCatalog[state.currentMetric];
    if (!metric) {
      if (state.chart) {
        state.chart.destroy();
        state.chart = null;
      }
      return;
    }
    const points = rows.map((row) => {
      const value = parseNumber(row[metric.valueCol]);
      const change = metric.changeCol ? parseNumber(row[metric.changeCol]) : NaN;
      const trend = resolveTrend(change, metric.directionCol ? row[metric.directionCol] : "");
      return {
        label: `${getLabel(row)} ${trendArrow(trend)}`,
        value,
        trend,
      };
    });

    const chartKind = state.chartType === "area" ? "line" : state.chartType;
    const isArea = state.chartType === "area";
    const validLabels = [];
    const validValues = [];
    const validTrends = [];
    for (let i = 0; i < points.length; i += 1) {
      if (!Number.isFinite(points[i].value)) {
        continue;
      }
      validLabels.push(points[i].label);
      validValues.push(points[i].value);
      validTrends.push(points[i].trend);
    }

    const pointColors = validTrends.map((trend) => trendColor(trend));
    updateSelectedMetricTrend(rows[rows.length - 1]);

    if (state.chart) {
      state.chart.destroy();
    }

    state.chart = new Chart(dom.chartCanvas, {
      type: chartKind,
      data: {
        labels: validLabels,
        datasets: [
          {
            label: metric.label,
            data: validValues,
            borderColor: chartKind === "bar" ? pointColors : "#25364f",
            backgroundColor: chartKind === "doughnut" || chartKind === "bar" ? pointColors : "rgba(37, 54, 79, 0.22)",
            pointBackgroundColor: pointColors,
            fill: isArea,
            borderWidth: 2,
            tension: 0.25,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: {
              color: "#101113",
              font: {
                family: "IBM Plex Sans",
              },
            },
          },
        },
        scales:
          chartKind === "doughnut"
            ? {}
            : {
                x: {
                  ticks: {
                    color: "#61656d",
                    maxRotation: 0,
                  },
                  grid: {
                    color: "rgba(16, 17, 19, 0.08)",
                  },
                },
                y: {
                  ticks: {
                    color: "#61656d",
                  },
                  grid: {
                    color: "rgba(16, 17, 19, 0.08)",
                  },
                },
              },
      },
    });
  }

  function getLabel(row) {
    const ts = row && row.__ts instanceof Date ? row.__ts : getRowTimestamp(row);
    if (ts instanceof Date && Number.isFinite(ts.getTime())) {
      return formatTimestampLabel(ts);
    }

    return (
      row["Current Date Range"] ||
      row["MONTH"] ||
      row["Month"] ||
      row["Date"] ||
      row["Timestamp"] ||
      "Point"
    );
  }

  function getRowTimestamp(row) {
    if (!row || typeof row !== "object") {
      return null;
    }
    const candidate = row["Logged At"] || row["Timestamp"] || row["Date"];
    return parseTimestamp(candidate);
  }

  function parseTimestamp(value) {
    if (value instanceof Date) {
      return Number.isFinite(value.getTime()) ? value : null;
    }
    if (typeof value === "number") {
      const d = new Date(value);
      return Number.isFinite(d.getTime()) ? d : null;
    }
    if (typeof value !== "string") {
      return null;
    }

    const cleaned = value.replace(/[\u200B-\u200D\uFEFF]/g, "").trim();
    if (!cleaned) {
      return null;
    }

    const parsedIso = new Date(cleaned.replace(" at ", " "));
    if (Number.isFinite(parsedIso.getTime())) {
      return parsedIso;
    }

    const dayFirstMatch = cleaned.match(
      /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:[ T](\d{1,2})(?::(\d{2}))?(?::(\d{2}))?)?$/
    );
    if (!dayFirstMatch) {
      return null;
    }

    const day = Number.parseInt(dayFirstMatch[1], 10);
    const month = Number.parseInt(dayFirstMatch[2], 10);
    const year = Number.parseInt(dayFirstMatch[3], 10);
    const hour = Number.parseInt(dayFirstMatch[4] || "0", 10);
    const minute = Number.parseInt(dayFirstMatch[5] || "0", 10);
    const second = Number.parseInt(dayFirstMatch[6] || "0", 10);

    const parsed = new Date(year, month - 1, day, hour, minute, second);
    return Number.isFinite(parsed.getTime()) ? parsed : null;
  }

  function isCurrentMonth(date) {
    if (!(date instanceof Date) || !Number.isFinite(date.getTime())) {
      return false;
    }
    const now = new Date();
    return date.getFullYear() === now.getFullYear() && date.getMonth() === now.getMonth();
  }

  function formatTimestampLabel(date) {
    return new Intl.DateTimeFormat("en-GB", {
      day: "2-digit",
      month: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      hour12: false,
    }).format(date);
  }

  function getClockDisplayParts(date) {
    const safeDate = date instanceof Date && Number.isFinite(date.getTime()) ? date : new Date();
    const timeParts = clockTimeFormatter.formatToParts(safeDate);
    const hours = getClockTimePart(timeParts, "hour");
    const minutes = getClockTimePart(timeParts, "minute");
    const seconds = getClockTimePart(timeParts, "second");
    const secondValue = Number.parseInt(seconds, 10);
    return {
      time: clockTimeFormatter.format(safeDate),
      dateDay: clockDateFormatter.format(safeDate),
      hours,
      minutes,
      seconds,
      isColonDim: Number.isFinite(secondValue) ? secondValue % 2 === 1 : false,
    };
  }

  function getClockTimePart(parts, type) {
    if (!Array.isArray(parts)) {
      return "00";
    }
    const match = parts.find((part) => part && part.type === type);
    const value = match && typeof match.value === "string" ? match.value.trim() : "";
    return value ? value.padStart(2, "0") : "00";
  }

  function updateSelectedMetricTrend(row) {
    if (!dom.selectedMetricTrend) {
      return;
    }
    if (!row) {
      dom.selectedMetricTrend.textContent = "Trend: --";
      dom.selectedMetricTrend.className = "selected-trend";
      return;
    }
    const metric = state.metricCatalog[state.currentMetric];
    if (!metric) {
      dom.selectedMetricTrend.textContent = "Trend: --";
      dom.selectedMetricTrend.className = "selected-trend";
      return;
    }
    const change = metric.changeCol ? parseNumber(row[metric.changeCol]) : NaN;
    const trend = resolveTrend(change, metric.directionCol ? row[metric.directionCol] : "");
    const sign = Number.isFinite(change) && change > 0 ? "+" : "";
    const pct = Number.isFinite(change) ? `${sign}${formatNumber(change, 0)}%` : "--";
    dom.selectedMetricTrend.textContent = `Trend (${metric.label}): ${trendArrow(trend)} ${trendLabel(trend)} ${pct}`;
    dom.selectedMetricTrend.className = `selected-trend ${trendClass(trend)}`.trim();
  }

  function parseNumber(value) {
    if (typeof value === "number") {
      return Number.isFinite(value) ? value : NaN;
    }
    if (typeof value !== "string") {
      return NaN;
    }

    const cleaned = value
      .replace(/[\u200B-\u200D\uFEFF]/g, "")
      .replace(/[,%£$€]/g, "")
      .trim();

    if (!cleaned) {
      return NaN;
    }

    const num = Number.parseFloat(cleaned);
    return Number.isFinite(num) ? num : NaN;
  }

  function calcPercentChange(current, previous) {
    if (!Number.isFinite(current) || !Number.isFinite(previous) || previous === 0) {
      return NaN;
    }
    return ((current - previous) / Math.abs(previous)) * 100;
  }

  function trendFromNumber(value) {
    if (!Number.isFinite(value)) {
      return "flat";
    }
    if (value > 0) {
      return "up";
    }
    if (value < 0) {
      return "down";
    }
    return "flat";
  }

  function roundTo(value, digits) {
    if (!Number.isFinite(value)) {
      return NaN;
    }
    const factor = 10 ** digits;
    return Math.round(value * factor) / factor;
  }

  function formatCurrency(value, currency) {
    return new Intl.NumberFormat("en-GB", {
      style: "currency",
      currency,
      maximumFractionDigits: 0,
    }).format(value);
  }

  function formatNumber(value, digits) {
    return new Intl.NumberFormat("en-GB", {
      minimumFractionDigits: digits,
      maximumFractionDigits: digits,
    }).format(value);
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function normalizeConfig(input) {
    const passcodeValue = typeof input.passcode === "undefined" || input.passcode === null ? "" : String(input.passcode).trim();
    const targetMultiplierRaw = parseNumber(input.salesTargetMultiplier);
    const targetMultiplier = Number.isFinite(targetMultiplierRaw) && targetMultiplierRaw > 0 ? targetMultiplierRaw : 1;
    const targetValueRaw = parseNumber(input.salesTargetValue);
    const targetsByMonth = normalizeSalesTargetsByMonth(input.salesTargetsByMonth);
    return {
      ...input,
      currencySymbol: DASHBOARD_CURRENCY,
      passcode: passcodeValue,
      salesTargetMultiplier: targetMultiplier,
      salesTargetValue: Number.isFinite(targetValueRaw) ? targetValueRaw : null,
      salesTargetsByMonth: targetsByMonth,
    };
  }

  function normalizeSalesTargetsByMonth(input) {
    if (!input || typeof input !== "object" || Array.isArray(input)) {
      return {};
    }
    const clean = {};
    Object.entries(input).forEach(([monthKey, rawValue]) => {
      if (!/^\d{4}-\d{2}$/.test(monthKey)) {
        return;
      }
      const parsed = parseNumber(rawValue);
      if (!Number.isFinite(parsed) || parsed <= 0) {
        return;
      }
      clean[monthKey] = roundTo(parsed, 2);
    });
    return clean;
  }

  function normalizePasscodeForCompare(value) {
    const raw = String(value || "");
    let normalized = raw;
    if (typeof raw.normalize === "function") {
      try {
        normalized = raw.normalize("NFKC");
      } catch (_error) {
        normalized = raw;
      }
    }
    normalized = normalized.replace(/\s+/g, "").trim();
    const digitsOnly = normalized.replace(/[^\d]/g, "");
    return digitsOnly || normalized;
  }

  function resolveTrend(change, directionRaw) {
    const direction = String(directionRaw || "").toLowerCase().trim();
    if (direction === "up" || direction === "down") {
      return direction;
    }
    if (Number.isFinite(change)) {
      if (change > 0) {
        return "up";
      }
      if (change < 0) {
        return "down";
      }
    }
    return "flat";
  }

  function trendArrow(trend) {
    if (trend === "up") {
      return "↑";
    }
    if (trend === "down") {
      return "↓";
    }
    return "→";
  }

  function trendLabel(trend) {
    if (trend === "up") {
      return "Up";
    }
    if (trend === "down") {
      return "Down";
    }
    return "Flat";
  }

  function trendClass(trend) {
    if (trend === "up") {
      return "up";
    }
    if (trend === "down") {
      return "down";
    }
    return "";
  }

  function trendColor(trend) {
    if (trend === "up") {
      return "rgba(11, 122, 58, 0.75)";
    }
    if (trend === "down") {
      return "rgba(162, 41, 41, 0.75)";
    }
    return "rgba(37, 54, 79, 0.55)";
  }

  function estimatePreviousValue(current, pctChange) {
    if (!Number.isFinite(current) || !Number.isFinite(pctChange)) {
      return NaN;
    }
    const ratio = 1 + pctChange / 100;
    if (ratio === 0) {
      return NaN;
    }
    return current / ratio;
  }

  function getPreviousMetricValue(previousRow, valueCol) {
    if (!previousRow) {
      return NaN;
    }
    return parseNumber(previousRow[valueCol]);
  }

  function clamp(value, min, max) {
    return Math.max(min, Math.min(value, max));
  }

  function clampPercent(value, min = 0, max = 100) {
    if (!Number.isFinite(value)) {
      return min;
    }
    return clamp(value, min, max);
  }

  function getQueryPasscode() {
    try {
      const params = new URLSearchParams(window.location.search);
      return normalizePasscodeForCompare(params.get("unlock") || "");
    } catch (_error) {
      return "";
    }
  }

  function setBootStatus(message, mode) {
    if (!dom.bootStatus) {
      return;
    }
    dom.bootStatus.classList.remove("hidden");
    dom.bootStatus.textContent = message;
    dom.bootStatus.classList.remove("ok", "err");
    if (mode === "ok" || mode === "err") {
      dom.bootStatus.classList.add(mode);
    }
    if (state.bootHideTimer) {
      window.clearTimeout(state.bootHideTimer);
      state.bootHideTimer = null;
    }
    if (mode === "ok") {
      state.bootHideTimer = window.setTimeout(() => {
        if (dom.bootStatus) {
          dom.bootStatus.classList.add("hidden");
        }
      }, 3500);
    }
  }

  function startAutoCycle() {
    stopAutoCycle();
    const keys = getChartMetricKeys();
    if (!keys.length) {
      return;
    }
    state.cycleTimer = window.setInterval(() => {
      const index = keys.indexOf(state.currentMetric);
      const nextIndex = (index + 1) % keys.length;
      const nextMetric = keys[nextIndex];
      state.currentMetric = nextMetric;
      if (dom.metricSelect) {
        dom.metricSelect.value = nextMetric;
      }
      renderChart();
      if (state.rows.length) {
        updateSelectedMetricTrend(state.rows[state.rows.length - 1]);
      }
    }, cfg.cycleIntervalMs);
  }

  function stopAutoCycle() {
    if (state.cycleTimer) {
      window.clearInterval(state.cycleTimer);
      state.cycleTimer = null;
    }
  }
})();
