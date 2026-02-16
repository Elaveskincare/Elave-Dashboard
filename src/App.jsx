import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import Chart from "chart.js/auto";
import {
  CalendarDays,
  ChartColumnIncreasing,
  Clock3,
  FileSpreadsheet,
  Lock,
  Maximize2,
  Minimize2,
  PencilLine,
  RefreshCcw,
  Search,
  Video,
} from "lucide-react";
import sampleData from "../sample-data.json";

import { Badge } from "./components/ui/badge";
import { Button } from "./components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "./components/ui/card";
import { Checkbox } from "./components/ui/checkbox";
import { Input } from "./components/ui/input";
import { Label } from "./components/ui/label";
import { ScrollArea } from "./components/ui/scroll-area";
import { Sheet, SheetContent, SheetDescription, SheetHeader, SheetTitle } from "./components/ui/sheet";
import { Switch } from "./components/ui/switch";
import { cn } from "./lib/utils";

const DASHBOARD_CURRENCY = "EUR";

const fallbackConfig = {
  appsScriptUrl: "",
  backendApiUrl: "",
  sheetName: "Triple Whale Hourly",
  refreshIntervalMs: 15 * 60 * 1000,
  cycleIntervalMs: 15 * 1000,
  passcode: "",
  salesTargetMultiplier: 1,
  salesTargetValue: null,
  salesTargetsByMonth: {},
  currencySymbol: DASHBOARD_CURRENCY,
};

const LAYOUT_STORAGE_KEY = "elave_dash_widget_layout_v2";
const MONTHLY_TARGETS_STORAGE_KEY = "elave_dash_monthly_targets_v1";
const METRIC_WIDGET_PREFIX = "metric:";
const ADVANCED_WIDGET_PREFIX = "advanced:";
const ADVANCED_CELL_LIST_LIMIT = 5;
const CALENDAR_PREVIEW_LIMIT = 4;
const CALENDAR_REFRESH_MS = 60 * 1000;
const VALUE_TWEEN_DURATION_MS = 650;
const VALUE_TWEEN_MIN_DELTA = 0.01;
const CARD_REVEAL_STAGGER_MS = 675;
const CARD_REVEAL_DURATION_MS = 2625;

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
  quick_links_panel: {
    key: "quick_links_panel",
    label: "Quick Links",
    type: "fixed",
    view: "quick_links",
    colSpan: 3,
    rowSpan: 1,
    accentClass: "metric-generic",
  },
  calendar_panel: {
    key: "calendar_panel",
    label: "Google Calendar Upcoming",
    type: "fixed",
    view: "calendar",
    colSpan: 3,
    rowSpan: 2,
    accentClass: "metric-calendar",
  },
};

const DEFAULT_CARD_KEYS = [
  "goal_panel",
  `${METRIC_WIDGET_PREFIX}sales`,
  `${METRIC_WIDGET_PREFIX}orders`,
  `${METRIC_WIDGET_PREFIX}aov`,
  `${METRIC_WIDGET_PREFIX}roas`,
  "clock_panel",
  "quick_links_panel",
  "calendar_panel",
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
  { key: "website_sessions_mtd", label: "Website Sessions (MTD)", colSpan: 3, rowSpan: 1 },
  { key: "new_vs_returning", label: "New vs Returning Revenue", colSpan: 3, rowSpan: 2 },
  { key: "channel_split", label: "Channel Split", colSpan: 3, rowSpan: 2 },
  { key: "discount_impact", label: "Discount Impact", colSpan: 3, rowSpan: 2 },
  { key: "ytd_orders", label: "Orders (YTD vs LY YTD)", colSpan: 3, rowSpan: 1 },
  { key: "ytd_total_sales", label: "Total Sales (YTD vs LY YTD)", colSpan: 3, rowSpan: 1 },
  { key: "ytd_growth_rate", label: "Growth Rate (YTD vs LY YTD)", colSpan: 3, rowSpan: 1 },
  { key: "hourly_heatmap_today", label: "Hourly Heatmap (Today)", colSpan: 3, rowSpan: 2 },
  { key: "refund_watchlist", label: "Refund Watchlist", colSpan: 3, rowSpan: 2 },
];

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

const dateTileWeekdayFormatter = new Intl.DateTimeFormat(undefined, { weekday: "short" });
const dateTileMonthFormatter = new Intl.DateTimeFormat(undefined, { month: "short" });
const dateTileDayFormatter = new Intl.DateTimeFormat(undefined, { day: "numeric" });
const calendarEventTimeFormatter = new Intl.DateTimeFormat(undefined, {
  hour: "numeric",
  minute: "2-digit",
});
const calendarEventDayTimeFormatter = new Intl.DateTimeFormat(undefined, {
  weekday: "short",
  hour: "numeric",
  minute: "2-digit",
});

const quickLinkItems = [
  {
    key: "meet",
    label: "Meet",
    href: "https://meet.google.com",
    ariaLabel: "Open Google Meet",
    toneClass: "quick-link-tile-meet",
    icon: Video,
  },
  {
    key: "google",
    label: "Google",
    href: "https://www.google.com",
    ariaLabel: "Open Google Search",
    toneClass: "quick-link-tile-google",
    icon: Search,
  },
  {
    key: "calendar",
    label: "Calendar",
    href: "https://calendar.google.com",
    ariaLabel: "Open Google Calendar",
    toneClass: "quick-link-tile-calendar",
    icon: CalendarDays,
  },
  {
    key: "sheets",
    label: "Sheets",
    href: "https://sheets.google.com",
    ariaLabel: "Open Google Sheets",
    toneClass: "quick-link-tile-sheets",
    icon: FileSpreadsheet,
  },
];

function App() {
  const cfg = useMemo(() => normalizeConfig({ ...fallbackConfig, ...(window.ELAVE_DASH_CONFIG || {}) }), []);

  const [bootStatus, setBootStatusState] = useState({ message: "Booting dashboard...", mode: "", visible: true });
  const [locked, setLocked] = useState(Boolean(cfg.passcode));
  const [passcodeInput, setPasscodeInput] = useState("");
  const [showPasscode, setShowPasscode] = useState(false);
  const [lockError, setLockError] = useState("");

  const [rows, setRows] = useState([]);
  const [cellsPayload, setCellsPayload] = useState(null);
  const [metricCatalog, setMetricCatalog] = useState(() => syncMetricCatalog([]));
  const [cardLayout, setCardLayout] = useState([]);
  const [hasCustomLayout, setHasCustomLayout] = useState(false);
  const [layoutHydrated, setLayoutHydrated] = useState(false);
  const [metricsHydrated, setMetricsHydrated] = useState(false);
  const [monthlyTargets, setMonthlyTargets] = useState({});
  const [lastUpdated, setLastUpdated] = useState("--");

  const [chartVisible, setChartVisible] = useState(false);
  const [chartType, setChartType] = useState("line");
  const [currentMetric, setCurrentMetric] = useState("sales");
  const [cycleEnabled, setCycleEnabled] = useState(false);

  const [cellMenuOpen, setCellMenuOpen] = useState(false);
  const [now, setNow] = useState(new Date());
  const [calendarPreview, setCalendarPreview] = useState({
    status: "idle",
    events: [],
    error: "",
    authUrl: "",
    updatedAt: "",
  });
  const [isFullscreen, setIsFullscreen] = useState(false);
  const [fullscreenSupported, setFullscreenSupported] = useState(true);
  const [viewportWidth, setViewportWidth] = useState(() =>
    typeof window === "undefined" ? 0 : window.innerWidth
  );

  const [activeDragMetric, setActiveDragMetric] = useState("");
  const [activeDragSource, setActiveDragSource] = useState("");
  const [dragDropSlot, setDragDropSlot] = useState(null);
  const [draggingWidgetId, setDraggingWidgetId] = useState("");
  const [dragTargetActive, setDragTargetActive] = useState(false);
  const [cardRevealActive, setCardRevealActive] = useState(true);

  const metricsBoardRef = useRef(null);
  const fullscreenTargetRef = useRef(null);
  const chartCanvasRef = useRef(null);
  const dropIndicatorRef = useRef(null);
  const chartRef = useRef(null);
  const refreshTimerRef = useRef(null);
  const clockTimerRef = useRef(null);
  const calendarTimerRef = useRef(null);
  const cycleTimerRef = useRef(null);
  const bootHideTimerRef = useRef(null);
  const wasFullscreenRef = useRef(false);

  const setBootStatus = useCallback((message, mode = "") => {
    setBootStatusState({ message, mode, visible: true });
    if (bootHideTimerRef.current) {
      window.clearTimeout(bootHideTimerRef.current);
      bootHideTimerRef.current = null;
    }
    if (mode === "ok") {
      bootHideTimerRef.current = window.setTimeout(() => {
        setBootStatusState((prev) => ({ ...prev, visible: false }));
      }, 3500);
    }
  }, []);

  useEffect(() => {
    setBootStatus("Binding state");
    const savedLayout = loadCardLayout();
    if (Array.isArray(savedLayout)) {
      setCardLayout(savedLayout);
      setHasCustomLayout(true);
    }
    setLayoutHydrated(true);
    setMonthlyTargets(loadMonthlyTargets());

    const queryPasscode = getQueryPasscode();
    const expected = normalizePasscodeForCompare(cfg.passcode);
    if (!expected || (queryPasscode && queryPasscode === expected)) {
      setLocked(false);
    }

    setBootStatus("Init complete", "ok");
  }, [cfg.passcode, setBootStatus]);

  const fetchAndRender = useCallback(async () => {
    setBootStatus("Loading data");
    try {
      const [loadedRows, loadedCells] = await Promise.all([loadRows(cfg), loadAdvancedCells(cfg)]);
      const catalog = syncMetricCatalog(loadedRows);
      setRows(loadedRows);
      setCellsPayload(loadedCells);
      setMetricCatalog(catalog);
      setLastUpdated(new Date().toLocaleString());
      if (cfg.backendApiUrl && !loadedCells) {
        setBootStatus("Backend unavailable: no live Shopify dashboard data", "err");
      } else {
        setBootStatus(`Loaded ${loadedRows.length} rows${loadedCells ? " + advanced data" : ""}`, "ok");
      }
    } catch (error) {
      console.error(error);
      setLastUpdated("error");
      setBootStatus("Data/chart error", "err");
    } finally {
      setMetricsHydrated(true);
    }
  }, [cfg, setBootStatus]);

  useEffect(() => {
    if (locked) {
      return undefined;
    }

    fetchAndRender();

    if (!clockTimerRef.current) {
      setNow(new Date());
      clockTimerRef.current = window.setInterval(() => setNow(new Date()), 1000);
    }

    refreshTimerRef.current = window.setInterval(fetchAndRender, cfg.refreshIntervalMs);

    return () => {
      if (refreshTimerRef.current) {
        window.clearInterval(refreshTimerRef.current);
        refreshTimerRef.current = null;
      }
      if (clockTimerRef.current) {
        window.clearInterval(clockTimerRef.current);
        clockTimerRef.current = null;
      }
    };
  }, [cfg.refreshIntervalMs, fetchAndRender, locked]);

  useEffect(() => {
    const syncViewportWidth = () => {
      setViewportWidth(window.innerWidth || 0);
    };

    syncViewportWidth();
    window.addEventListener("resize", syncViewportWidth);
    return () => window.removeEventListener("resize", syncViewportWidth);
  }, []);

  useEffect(() => {
    if (locked) {
      return undefined;
    }

    if (!cfg.backendApiUrl) {
      setCalendarPreview({
        status: "backend_missing",
        events: [],
        error: "Backend API is required for Google Calendar widgets.",
        authUrl: "",
        updatedAt: "",
      });
      return undefined;
    }

    let cancelled = false;
    const refreshCalendar = async () => {
      setCalendarPreview((prev) =>
        prev.status === "ready" && Array.isArray(prev.events) && prev.events.length
          ? prev
          : { ...prev, status: "loading", error: "" }
      );
      const payload = await loadGoogleCalendarPreview(cfg, { limit: CALENDAR_PREVIEW_LIMIT });
      if (cancelled) {
        return;
      }
      setCalendarPreview(payload);
    };

    refreshCalendar();
    calendarTimerRef.current = window.setInterval(refreshCalendar, CALENDAR_REFRESH_MS);
    const handleVisibility = () => {
      if (!document.hidden) {
        refreshCalendar();
      }
    };
    document.addEventListener("visibilitychange", handleVisibility);

    return () => {
      cancelled = true;
      document.removeEventListener("visibilitychange", handleVisibility);
      if (calendarTimerRef.current) {
        window.clearInterval(calendarTimerRef.current);
        calendarTimerRef.current = null;
      }
    };
  }, [cfg.backendApiUrl, locked]);

  useEffect(() => {
    const supportsFullscreen =
      Boolean(document.fullscreenEnabled) ||
      Boolean(document.webkitFullscreenEnabled) ||
      Boolean(document.documentElement?.requestFullscreen) ||
      Boolean(document.documentElement?.webkitRequestFullscreen);

    setFullscreenSupported(supportsFullscreen);

    const syncFullscreenState = () => {
      setIsFullscreen(Boolean(document.fullscreenElement || document.webkitFullscreenElement));
    };

    syncFullscreenState();
    document.addEventListener("fullscreenchange", syncFullscreenState);
    document.addEventListener("webkitfullscreenchange", syncFullscreenState);

    return () => {
      document.removeEventListener("fullscreenchange", syncFullscreenState);
      document.removeEventListener("webkitfullscreenchange", syncFullscreenState);
    };
  }, []);

  const requestFullscreen = useCallback(async () => {
    try {
      const targetElement = fullscreenTargetRef.current || document.documentElement;
      if (targetElement.requestFullscreen) {
        await targetElement.requestFullscreen();
      } else if (targetElement.webkitRequestFullscreen) {
        targetElement.webkitRequestFullscreen();
      } else {
        setBootStatus("Fullscreen is not supported in this browser", "err");
      }
    } catch (error) {
      console.error(error);
      setBootStatus("Fullscreen request blocked by browser", "err");
    }
  }, [setBootStatus]);

  const handleToggleFullscreen = useCallback(async () => {
    const activeElement = document.fullscreenElement || document.webkitFullscreenElement;

    if (activeElement) {
      try {
        if (document.exitFullscreen) {
          await document.exitFullscreen();
        } else if (document.webkitExitFullscreen) {
          document.webkitExitFullscreen();
        }
      } catch (error) {
        console.error(error);
        setBootStatus("Fullscreen request blocked by browser", "err");
      }
      return;
    }

    await requestFullscreen();
  }, [requestFullscreen, setBootStatus]);

  const availableWidgetDefinitions = useMemo(() => getAvailableWidgetDefinitions(metricCatalog), [metricCatalog]);
  const availableWidgetKeys = useMemo(() => availableWidgetDefinitions.map((widget) => widget.key), [availableWidgetDefinitions]);

  useEffect(() => {
    if (!layoutHydrated || !metricsHydrated) {
      return;
    }

    setCardLayout((prevLayout) => {
      const gridConfig = getLayoutGridConfig(metricsBoardRef.current);
      const result = ensureValidCardLayout({
        layout: prevLayout,
        hasCustomLayout,
        availableWidgetKeys,
        metricCatalog,
        slotCols: gridConfig.slotCols,
        metricCardSpan: gridConfig.metricCardSpan,
      });

      if (result.changed) {
        saveCardLayout(result.layout);
      }
      if (!hasCustomLayout && result.layout.length && result.fromDefault) {
        saveCardLayout(result.layout);
      }

      return result.layout;
    });
  }, [availableWidgetKeys, hasCustomLayout, layoutHydrated, metricCatalog, metricsHydrated]);

  const chartMetricKeys = useMemo(() => getChartMetricKeys(cardLayout, metricCatalog), [cardLayout, metricCatalog]);

  useEffect(() => {
    if (!chartMetricKeys.length) {
      setCurrentMetric("");
      return;
    }
    if (!chartMetricKeys.includes(currentMetric)) {
      setCurrentMetric(chartMetricKeys[0]);
    }
  }, [chartMetricKeys, currentMetric]);

  useEffect(() => {
    if (!cycleEnabled || !chartVisible || !chartMetricKeys.length || !cfg.cycleIntervalMs) {
      if (cycleTimerRef.current) {
        window.clearInterval(cycleTimerRef.current);
        cycleTimerRef.current = null;
      }
      return undefined;
    }

    cycleTimerRef.current = window.setInterval(() => {
      setCurrentMetric((prev) => {
        const index = chartMetricKeys.indexOf(prev);
        const nextIndex = (index + 1) % chartMetricKeys.length;
        return chartMetricKeys[nextIndex];
      });
    }, cfg.cycleIntervalMs);

    return () => {
      if (cycleTimerRef.current) {
        window.clearInterval(cycleTimerRef.current);
        cycleTimerRef.current = null;
      }
    };
  }, [cfg.cycleIntervalMs, chartMetricKeys, chartVisible, cycleEnabled]);

  const hasChartWidget = useMemo(
    () => cardLayout.some((entry) => entry && entry.key === "chart_panel"),
    [cardLayout]
  );

  useEffect(() => {
    if (!hasChartWidget) {
      setChartVisible(false);
      setCycleEnabled(false);
    }
  }, [hasChartWidget]);

  const sortedLayout = useMemo(
    () => sortLayoutEntries(cardLayout.filter((entry) => entry && isWidgetAvailable(entry.key, metricCatalog))),
    [cardLayout, metricCatalog]
  );

  const latestRow = rows.length ? rows[rows.length - 1] : null;
  const previousRow = rows.length > 1 ? rows[rows.length - 2] : null;

  const salesMeta = useMemo(() => getSalesGoalMetric(metricCatalog, cardLayout), [metricCatalog, cardLayout]);

  const goalData = useMemo(
    () =>
      buildGoalData({
        cfg,
        row: latestRow,
        previousRow,
        salesMeta,
        cellsPayload,
        monthlyTargets,
      }),
    [cfg, latestRow, previousRow, salesMeta, cellsPayload, monthlyTargets]
  );

  const selectedTrend = useMemo(() => {
    if (!latestRow || !currentMetric || !metricCatalog[currentMetric]) {
      return { text: "Trend: --", className: "" };
    }
    const metric = metricCatalog[currentMetric];
    const change = metric.changeCol ? parseNumber(latestRow[metric.changeCol]) : NaN;
    const trend = resolveTrend(change, metric.directionCol ? latestRow[metric.directionCol] : "");
    const sign = Number.isFinite(change) && change > 0 ? "+" : "";
    const pct = Number.isFinite(change) ? `${sign}${formatNumber(change, 0)}%` : "--";
    return {
      text: `Trend (${metric.label}): ${trendArrow(trend)} ${trendLabel(trend)} ${pct}`,
      className: trendClass(trend),
    };
  }, [currentMetric, latestRow, metricCatalog]);

  useEffect(() => {
    if (locked || !cardRevealActive) {
      return undefined;
    }

    const maxDelay = sortedLayout.length * CARD_REVEAL_STAGGER_MS;

    const timeoutId = window.setTimeout(() => {
      setCardRevealActive(false);
    }, maxDelay + CARD_REVEAL_DURATION_MS + 80);

    return () => {
      window.clearTimeout(timeoutId);
    };
  }, [cardRevealActive, locked, sortedLayout.length]);

  useEffect(() => {
    if (locked) {
      wasFullscreenRef.current = isFullscreen;
      return undefined;
    }

    const enteredFullscreen = isFullscreen && !wasFullscreenRef.current;
    wasFullscreenRef.current = isFullscreen;

    if (!enteredFullscreen) {
      return undefined;
    }

    setCardRevealActive(false);
    const rafId = window.requestAnimationFrame(() => {
      setCardRevealActive(true);
    });

    return () => {
      window.cancelAnimationFrame(rafId);
    };
  }, [isFullscreen, locked]);

  const handleUnlock = useCallback(
    (event) => {
      event.preventDefault();
      const entered = normalizePasscodeForCompare(passcodeInput);
      const expected = normalizePasscodeForCompare(cfg.passcode);

      if (!expected || entered === expected) {
        setLockError("");
        setPasscodeInput("");
        const activeElement = document.fullscreenElement || document.webkitFullscreenElement;
        if (!activeElement) {
          void requestFullscreen();
        }
        setLocked(false);
        return;
      }
      setLockError(`Incorrect passcode (expected ${expected.length} digits)`);
    },
    [cfg.passcode, passcodeInput, requestFullscreen]
  );

  const saveTargets = useCallback(
    (nextTargets) => {
      setMonthlyTargets(nextTargets);
      try {
        window.localStorage.setItem(MONTHLY_TARGETS_STORAGE_KEY, JSON.stringify(nextTargets));
      } catch (_error) {
        // ignore storage errors
      }
    },
    [setMonthlyTargets]
  );

  const handleSetMonthlyTarget = useCallback(() => {
    const monthKey = getCurrentGoalMonthKey();
    const monthLabel = formatMonthLabel(monthKey);
    const existing =
      getMonthlyTargetForKey(monthlyTargets, monthKey) ||
      getConfigTargetForKey(cfg.salesTargetsByMonth, monthKey) ||
      parseNumber(cfg.salesTargetValue) ||
      "";

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

    saveTargets({ ...monthlyTargets, [monthKey]: roundTo(parsed, 2) });
  }, [cfg.currencySymbol, cfg.salesTargetValue, cfg.salesTargetsByMonth, monthlyTargets, saveTargets]);

  const handleClearMonthlyTarget = useCallback(() => {
    const monthKey = getCurrentGoalMonthKey();
    if (!Object.prototype.hasOwnProperty.call(monthlyTargets, monthKey)) {
      window.alert(`No saved monthly target found for ${formatMonthLabel(monthKey)}.`);
      return;
    }

    const next = { ...monthlyTargets };
    delete next[monthKey];
    saveTargets(next);
  }, [monthlyTargets, saveTargets]);

  const addOrMoveCard = useCallback(
    (widgetKey, targetSlot) => {
      if (!widgetKey || !isWidgetAvailable(widgetKey, metricCatalog)) {
        return;
      }

      const normalizedKey = normalizeWidgetKey(widgetKey);
      const gridConfig = getLayoutGridConfig(metricsBoardRef.current);

      setCardLayout((prev) => {
        const existingEntry = prev.find((entry) => entry.key === normalizedKey);
        if (existingEntry && !targetSlot) {
          return prev;
        }

        const remaining = prev.filter((entry) => entry.key !== normalizedKey);
        const occupied = buildOccupiedRects(remaining, gridConfig.slotCols, gridConfig.metricCardSpan, metricCatalog);

        const desiredCol = targetSlot && Number.isFinite(targetSlot.col) ? targetSlot.col : 1;
        const desiredRow = targetSlot && Number.isFinite(targetSlot.row) ? targetSlot.row : 1;
        const nextSlot = findFirstAvailableSlot(
          normalizedKey,
          desiredCol,
          desiredRow,
          gridConfig.slotCols,
          occupied,
          gridConfig.metricCardSpan,
          metricCatalog
        );

        const nextLayout = [...remaining, { key: normalizedKey, col: nextSlot.col, row: nextSlot.row }];
        saveCardLayout(nextLayout);
        return nextLayout;
      });

      setHasCustomLayout(true);
    },
    [metricCatalog]
  );

  const handleRemoveWidget = useCallback(
    (widgetKey) => {
      setCardLayout((prev) => {
        const next = prev.filter((entry) => entry.key !== widgetKey);
        saveCardLayout(next);
        return next;
      });
      setHasCustomLayout(true);
      if (widgetKey === "chart_panel") {
        setChartVisible(false);
        setCycleEnabled(false);
      }
    },
    [setCardLayout]
  );

  const onLibraryDragStart = useCallback((event, widgetKey) => {
    setActiveDragMetric(widgetKey);
    setActiveDragSource("library");
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "copyMove";
      event.dataTransfer.setData("text/plain", widgetKey);
    }
  }, []);

  const onWidgetDragStart = useCallback((event, widgetKey) => {
    setActiveDragMetric(widgetKey);
    setActiveDragSource("board");
    setDraggingWidgetId(widgetKey);
    if (event.dataTransfer) {
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", widgetKey);
    }
  }, []);

  const resolvePreferredDropSlot = useCallback(
    (widgetKey, desiredSlot) => {
      const gridConfig = getLayoutGridConfig(metricsBoardRef.current);
      const remaining = cardLayout.filter((entry) => entry.key !== widgetKey);
      const occupied = buildOccupiedRects(remaining, gridConfig.slotCols, gridConfig.metricCardSpan, metricCatalog);
      return findFirstAvailableSlot(
        widgetKey,
        desiredSlot.col,
        desiredSlot.row,
        gridConfig.slotCols,
        occupied,
        gridConfig.metricCardSpan,
        metricCatalog
      );
    },
    [cardLayout, metricCatalog]
  );

  const getDropSlot = useCallback(
    (event, widgetKey) => {
      const geometry = getMetricGridGeometry(metricsBoardRef.current);
      if (!geometry) {
        return null;
      }

      const x = clamp(event.clientX - geometry.rect.left, 0, Math.max(0, geometry.rect.width - 1));
      const y = clamp(event.clientY - geometry.rect.top, 0, Math.max(0, geometry.rect.height + geometry.rowStep));
      const widgetSize = getWidgetSlotSize(widgetKey, geometry, metricCatalog);
      const maxColStart = Math.max(1, geometry.slotCols - widgetSize.colSlots + 1);
      const col = clamp(Math.floor(x / geometry.slotWidth) + 1, 1, maxColStart);
      const row = Math.max(1, Math.floor(y / geometry.rowStep) + 1);
      return { col, row };
    },
    [metricCatalog]
  );

  const onMetricsBoardDragOver = useCallback(
    (event) => {
      if (!activeDragMetric) {
        return;
      }
      event.preventDefault();
      setDragTargetActive(true);

      if (event.dataTransfer) {
        event.dataTransfer.dropEffect = activeDragSource === "library" ? "copy" : "move";
      }

      const pointerSlot = getDropSlot(event, activeDragMetric);
      if (!pointerSlot) {
        setDragDropSlot(null);
        return;
      }
      setDragDropSlot(resolvePreferredDropSlot(activeDragMetric, pointerSlot));
    },
    [activeDragMetric, activeDragSource, getDropSlot, resolvePreferredDropSlot]
  );

  const onMetricsBoardDrop = useCallback(
    (event) => {
      event.preventDefault();
      setDragTargetActive(false);
      if (!activeDragMetric) {
        return;
      }

      const dropSlot = dragDropSlot || getDropSlot(event, activeDragMetric);
      addOrMoveCard(activeDragMetric, dropSlot);

      setActiveDragMetric("");
      setActiveDragSource("");
      setDragDropSlot(null);
      setDraggingWidgetId("");
    },
    [activeDragMetric, addOrMoveCard, dragDropSlot, getDropSlot]
  );

  const onMetricsBoardDragLeave = useCallback((event) => {
    if (!metricsBoardRef.current || metricsBoardRef.current.contains(event.relatedTarget)) {
      return;
    }
    setDragTargetActive(false);
    setDragDropSlot(null);
  }, []);

  useEffect(() => {
    const onAnyDragEnd = () => {
      setDragTargetActive(false);
      setActiveDragMetric("");
      setActiveDragSource("");
      setDragDropSlot(null);
      setDraggingWidgetId("");
    };

    window.addEventListener("dragend", onAnyDragEnd);
    return () => window.removeEventListener("dragend", onAnyDragEnd);
  }, []);

  useEffect(() => {
    const indicator = dropIndicatorRef.current;
    if (!indicator || !dragDropSlot || !activeDragMetric) {
      if (indicator) {
        indicator.classList.remove("is-visible");
      }
      return;
    }

    const geometry = getMetricGridGeometry(metricsBoardRef.current);
    if (!geometry) {
      indicator.classList.remove("is-visible");
      return;
    }

    const widgetSize = getWidgetSlotSize(activeDragMetric, geometry, metricCatalog);
    const x = (dragDropSlot.col - 1) * geometry.slotWidth;
    const y = (dragDropSlot.row - 1) * geometry.rowStep;
    const width = Math.max(72, geometry.slotWidth * widgetSize.colSlots - geometry.gap);
    const height = Math.max(72, geometry.rowHeight * widgetSize.rowSlots + geometry.gap * (widgetSize.rowSlots - 1));

    indicator.style.width = `${width}px`;
    indicator.style.height = `${height}px`;
    indicator.style.transform = `translate(${x}px, ${y}px)`;
    indicator.classList.add("is-visible");
  }, [activeDragMetric, dragDropSlot, metricCatalog]);

  useEffect(() => {
    if (!chartVisible || !chartCanvasRef.current || !rows.length) {
      if (chartRef.current) {
        chartRef.current.destroy();
        chartRef.current = null;
      }
      return;
    }

    const metric = metricCatalog[currentMetric];
    if (!metric) {
      if (chartRef.current) {
        chartRef.current.destroy();
        chartRef.current = null;
      }
      return;
    }

    const points = rows
      .map((row) => {
        const value = parseNumber(row[metric.valueCol]);
        const change = metric.changeCol ? parseNumber(row[metric.changeCol]) : NaN;
        const trend = resolveTrend(change, metric.directionCol ? row[metric.directionCol] : "");
        return {
          label: `${getLabel(row)} ${trendArrow(trend)}`,
          value,
          trend,
        };
      })
      .filter((point) => Number.isFinite(point.value));

    const chartKind = chartType === "area" ? "line" : chartType;
    const isArea = chartType === "area";

    const labels = points.map((point) => point.label);
    const values = points.map((point) => point.value);
    const chartTypography = getChartTypography(viewportWidth);
    if (chartRef.current) {
      chartRef.current.destroy();
    }

    chartRef.current = new Chart(chartCanvasRef.current, {
      type: chartKind,
      data: {
        labels,
        datasets: [
          {
            label: metric.label,
            data: values,
            borderColor: chartKind === "bar" ? "rgba(229, 229, 229, 0.65)" : "rgba(245, 245, 245, 0.88)",
            backgroundColor:
              chartKind === "doughnut" || chartKind === "bar" ? "rgba(229, 229, 229, 0.35)" : "rgba(245, 245, 245, 0.16)",
            pointBackgroundColor: "rgba(245, 245, 245, 0.9)",
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
              color: "#e5e5e5",
              font: {
                family: "IBM Plex Sans",
                size: chartTypography.legendSize,
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
                    color: "#a3a3a3",
                    maxRotation: 0,
                    font: {
                      size: chartTypography.tickSize,
                    },
                  },
                  grid: {
                    color: "rgba(163, 163, 163, 0.16)",
                  },
                },
                y: {
                  ticks: {
                    color: "#a3a3a3",
                    font: {
                      size: chartTypography.tickSize,
                    },
                  },
                  grid: {
                    color: "rgba(163, 163, 163, 0.16)",
                  },
                },
              },
      },
    });

    return () => {
      if (chartRef.current) {
        chartRef.current.destroy();
        chartRef.current = null;
      }
    };
  }, [chartType, chartVisible, currentMetric, metricCatalog, rows, viewportWidth]);

  useEffect(() => {
    return () => {
      if (bootHideTimerRef.current) {
        window.clearTimeout(bootHideTimerRef.current);
      }
      if (cycleTimerRef.current) {
        window.clearInterval(cycleTimerRef.current);
      }
      if (refreshTimerRef.current) {
        window.clearInterval(refreshTimerRef.current);
      }
      if (clockTimerRef.current) {
        window.clearInterval(clockTimerRef.current);
      }
      if (calendarTimerRef.current) {
        window.clearInterval(calendarTimerRef.current);
      }
      if (chartRef.current) {
        chartRef.current.destroy();
      }
    };
  }, []);

  const chartMetricOptions = chartMetricKeys.map((key) => ({ key, label: metricCatalog[key]?.label || key }));

  const cellLibrary = useMemo(() => {
    return [...availableWidgetDefinitions].sort((a, b) => a.label.localeCompare(b.label));
  }, [availableWidgetDefinitions]);

  const boardWidgets = sortedLayout.map((entry, index) => {
    const widget = getWidgetDefinition(entry.key, metricCatalog);
    if (!widget) {
      return null;
    }
    const styleObject = styleVarsFromPlacement(buildWidgetPlacementStyle(entry, widget, getLayoutGridConfig(metricsBoardRef.current)));
    const revealDelay = (index + 1) * CARD_REVEAL_STAGGER_MS;
    const showReveal = cardRevealActive;
    if (showReveal) {
      styleObject["--card-reveal-delay"] = `${revealDelay}ms`;
    }
    return {
      widget,
      showReveal,
      styleObject,
    };
  });

  const topBarRevealStyle = cardRevealActive ? { "--card-reveal-delay": "0ms" } : undefined;

  const clockParts = getClockDisplayParts(now);
  const dateTileParts = getDateTileDisplayParts(now);

  return (
    <div ref={fullscreenTargetRef} className="dashboard-shell min-h-screen px-4 py-4 md:px-6 md:py-6">
      <div
        className={cn(
          "pointer-events-none fixed left-1/2 top-3 z-50 -translate-x-1/2 rounded-md border bg-card/95 px-3 py-1.5 text-xs shadow-sm transition-opacity",
          bootStatus.mode === "ok" && "border-positive/40 text-positive",
          bootStatus.mode === "err" && "border-destructive/50 text-destructive",
          bootStatus.visible ? "opacity-100" : "opacity-0"
        )}
      >
        {bootStatus.message}
      </div>

      {locked ? (
        <main className="mx-auto flex min-h-[calc(100svh-2rem)] w-full max-w-md items-center justify-center md:min-h-[calc(100svh-3rem)]">
          <Card className="w-full rounded-2xl border-white/10 bg-[#0f0f0f]/95 shadow-lg shadow-black/50">
            <CardHeader>
              <CardDescription className="flex justify-center text-zinc-400">
                <img src="/elave-apothecary-white.svg" alt="Elave Apothecary" className="h-[112px] w-auto md:h-[126px]" />
              </CardDescription>
            </CardHeader>
            <CardContent>
              <form className="space-y-4" onSubmit={handleUnlock}>
                <div className="space-y-2">
                  <Label htmlFor="passcode-input">Passcode</Label>
                  <Input
                    id="passcode-input"
                    type={showPasscode ? "text" : "password"}
                    inputMode="numeric"
                    autoComplete="off"
                    placeholder="Passcode"
                    value={passcodeInput}
                    onChange={(event) => setPasscodeInput(event.target.value)}
                  />
                </div>
                <div className="flex items-center gap-2">
                  <Checkbox
                    id="show-pass-toggle"
                    checked={showPasscode}
                    onCheckedChange={(value) => setShowPasscode(Boolean(value))}
                  />
                  <Label htmlFor="show-pass-toggle">Show passcode</Label>
                </div>
                <div className="grid gap-2 sm:grid-cols-2">
                  <Button type="submit" className="w-full gap-2">
                    <Lock className="h-4 w-4" />
                    Unlock
                  </Button>
                  <Button
                    type="button"
                    variant="outline"
                    className="w-full"
                    onClick={handleToggleFullscreen}
                    disabled={!fullscreenSupported}
                  >
                    {isFullscreen ? "Exit Full Screen" : "Full Screen"}
                  </Button>
                </div>
                <p className="text-xs text-muted-foreground">Press Enter or click Unlock.</p>
                {lockError ? <p className="text-sm text-zinc-300">{lockError}</p> : null}
              </form>
            </CardContent>
          </Card>
        </main>
      ) : (
        <main className="dashboard-main mx-auto flex w-full max-w-[1880px] flex-col gap-4">
          <Card
            className={cn(
              "relative overflow-hidden rounded-2xl border-white/10 bg-[#0d0d0d]/95 shadow-[0_20px_44px_-30px_rgba(0,0,0,0.95)] backdrop-blur-md",
              cardRevealActive && "card-reveal-in"
            )}
            style={topBarRevealStyle}
          >
            <CardContent className="p-4 md:p-5">
              <div className="relative flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
                <div className="space-y-1">
                  <p className="font-mono text-[0.7rem] uppercase tracking-[0.34em] text-zinc-500">Elave Skincare</p>
                  <h1 className="text-2xl font-semibold tracking-tight text-zinc-50 md:text-[2.05rem]">Marketing Office</h1>
                </div>
                <div className="flex flex-col gap-2.5 sm:items-end">
                  <Badge
                    variant="secondary"
                    className="inline-flex w-fit items-center gap-1.5 border border-white/15 bg-white/10 px-3 py-1 font-medium text-zinc-200"
                  >
                    <Clock3 className="h-3.5 w-3.5 text-zinc-300" />
                    Last updated: {lastUpdated}
                  </Badge>
                  <div className="flex flex-wrap items-center gap-2 sm:justify-end">
                    <Button
                      onClick={fetchAndRender}
                      size="sm"
                      variant="outline"
                      className="h-10 rounded-xl border-white/20 bg-black/30 px-4 text-zinc-100 shadow-[inset_0_1px_0_rgba(255,255,255,0.08)] hover:bg-white/10 hover:text-white"
                    >
                      <RefreshCcw className="h-4 w-4" />
                      Refresh
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={handleToggleFullscreen}
                      disabled={!fullscreenSupported}
                      className="h-10 rounded-xl border-white/20 bg-black/30 px-4 text-zinc-100 shadow-[inset_0_1px_0_rgba(255,255,255,0.08)] hover:bg-white/10 hover:text-white"
                    >
                      {isFullscreen ? <Minimize2 className="h-4 w-4" /> : <Maximize2 className="h-4 w-4" />}
                      {isFullscreen ? "Exit Full Screen" : "Full Screen"}
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => setCellMenuOpen(true)}
                      className="h-10 rounded-xl border-white/20 bg-black/30 px-4 text-zinc-100 shadow-[inset_0_1px_0_rgba(255,255,255,0.08)] hover:bg-white/10 hover:text-white"
                    >
                      <PencilLine className="h-4 w-4" />
                      Edit Cells
                    </Button>
                  </div>
                </div>
              </div>
            </CardContent>
          </Card>

          <section className="space-y-3">
            <div
              ref={metricsBoardRef}
              className={cn("metrics-board rounded-xl", dragTargetActive && "drag-target")}
              onDragOver={onMetricsBoardDragOver}
              onDrop={onMetricsBoardDrop}
              onDragLeave={onMetricsBoardDragLeave}
            >
              {boardWidgets.length ? null : (
                <Card className="layout-widget col-span-full row-span-1 rounded-2xl border-dashed bg-[#101010]">
                  <CardContent className="p-6 text-sm text-muted-foreground">
                    No cells on board. Open Edit Cells and drag metrics in.
                  </CardContent>
                </Card>
              )}

              {boardWidgets.map((item) => {
                if (!item) {
                  return null;
                }

                const { widget, showReveal, styleObject } = item;

                if (widget.type === "fixed" && widget.view === "goal") {
                  return (
                    <Card
                      key={widget.key}
                      className={cn(
                        "layout-widget flex h-full flex-col rounded-2xl border-white/10 bg-[#121212]/95 shadow-[0_0_0_1px_rgba(255,255,255,0.02)]",
                        showReveal && "card-reveal-in",
                        draggingWidgetId === widget.key && "card-dragging"
                      )}
                      draggable
                      onDragStart={(event) => onWidgetDragStart(event, widget.key)}
                      style={styleObject}
                    >
                      <CardHeader className="pb-4">
                        <div className="flex items-center justify-between gap-2">
                          <div>
                            <CardTitle className="text-lg">Target Progress</CardTitle>
                            <CardDescription>MTD total sales vs target and last month total</CardDescription>
                          </div>
                          <div className="flex items-center gap-2">
                            <Button size="sm" variant="outline" onClick={handleSetMonthlyTarget}>
                              Set Target
                            </Button>
                            <Button size="sm" variant="outline" onClick={handleClearMonthlyTarget}>
                              Clear
                            </Button>
                            <Button size="sm" variant="ghost" onClick={() => handleRemoveWidget(widget.key)}>
                              Remove
                            </Button>
                          </div>
                        </div>
                      </CardHeader>
                      <CardContent className="flex flex-1 flex-col gap-5">
                        <div className="grid grid-cols-1 gap-3.5 text-sm md:grid-cols-3 md:text-base">
                          <StatPill
                            label="Current"
                            value={goalData.currentText}
                            numericValue={goalData.currentValue}
                            formatValue={(num) => formatCurrency(num, cfg.currencySymbol)}
                            tone="neutral"
                            note={goalData.currentNote}
                          />
                          <StatPill
                            label="Target"
                            value={goalData.targetText}
                            numericValue={goalData.targetValue}
                            formatValue={(num) => formatCurrency(num, cfg.currencySymbol)}
                            tone="warning"
                            note={goalData.targetNote}
                          />
                          <StatPill
                            label="Last Month"
                            value={goalData.previousText}
                            numericValue={goalData.previousValue}
                            formatValue={(num) => formatCurrency(num, cfg.currencySymbol)}
                            tone="neutral"
                            note={goalData.previousNote}
                          />
                        </div>

                        <div className="space-y-2">
                          <div className="flex items-center justify-between text-base md:text-lg">
                            <span className={cn("font-semibold", goalData.statusClass)}>{goalData.statusText}</span>
                            <AnimatedNumber
                              as="span"
                              className="font-mono text-base md:text-lg"
                              value={goalData.progressValue}
                              formatValue={(num) => `${formatNumber(num, 1)}%`}
                              fallback={goalData.progressText}
                            />
                          </div>
                          <div className="h-[1.2rem] overflow-hidden rounded-full bg-zinc-800/80">
                            <div
                              className={cn("goal-fill h-full rounded-full", goalData.fillClass)}
                              style={{ width: `${goalData.progressClamped}%` }}
                            />
                          </div>
                        </div>

                        <div className="mt-auto grid grid-cols-1 gap-x-7 gap-y-2 text-[clamp(0.96rem,0.95vw,1.2rem)] leading-[1.32] text-zinc-300 md:grid-cols-2">
                          <p>{goalData.milestonePrev}</p>
                          <p>{goalData.milestoneTarget}</p>
                          <p className={goalData.gapPrevClass}>{goalData.gapPrev}</p>
                          <p className={goalData.gapTargetClass}>{goalData.gapTarget}</p>
                        </div>
                      </CardContent>
                    </Card>
                  );
                }

                if (widget.type === "fixed" && widget.view === "chart") {
                  return (
                    <Card
                      key={widget.key}
                      className={cn(
                        "layout-widget rounded-2xl border-white/10 bg-[#121212]/95 shadow-[0_0_0_1px_rgba(255,255,255,0.02)]",
                        showReveal && "card-reveal-in",
                        draggingWidgetId === widget.key && "card-dragging"
                      )}
                      draggable
                      onDragStart={(event) => onWidgetDragStart(event, widget.key)}
                      style={styleObject}
                    >
                      <CardHeader className="pb-3">
                        <div className="flex items-center justify-between gap-2">
                          <div>
                            <CardTitle className="text-lg">Trend Chart</CardTitle>
                            <CardDescription>Switch metric, style, and cycle behavior</CardDescription>
                          </div>
                          <div className="flex items-center gap-2">
                            <Button
                              size="sm"
                              variant="outline"
                              onClick={() => setChartVisible((prev) => !prev)}
                              className="gap-2"
                            >
                              <ChartColumnIncreasing className="h-4 w-4" />
                              {chartVisible ? "Hide Chart" : "Show Chart"}
                            </Button>
                            <Button size="sm" variant="ghost" onClick={() => handleRemoveWidget(widget.key)}>
                              Remove
                            </Button>
                          </div>
                        </div>
                      </CardHeader>
                      <CardContent className={cn("space-y-4", !chartVisible && "hidden")}>
                        <div className="grid grid-cols-1 gap-3 md:grid-cols-4">
                          <label className="space-y-1 text-sm">
                            <span className="font-medium">Metric</span>
                            <select
                              className="h-10 w-full rounded-md border border-input bg-background px-3 text-sm"
                              value={currentMetric}
                              onChange={(event) => setCurrentMetric(event.target.value)}
                              disabled={!chartMetricOptions.length}
                            >
                              {chartMetricOptions.length ? (
                                chartMetricOptions.map((item) => (
                                  <option key={item.key} value={item.key}>
                                    {item.label}
                                  </option>
                                ))
                              ) : (
                                <option value="">No metrics</option>
                              )}
                            </select>
                          </label>

                          <label className="space-y-1 text-sm">
                            <span className="font-medium">Chart Style</span>
                            <select
                              className="h-10 w-full rounded-md border border-input bg-background px-3 text-sm"
                              value={chartType}
                              onChange={(event) => setChartType(event.target.value)}
                            >
                              <option value="line">Line</option>
                              <option value="bar">Bar</option>
                              <option value="area">Area</option>
                              <option value="doughnut">Doughnut</option>
                            </select>
                          </label>

                          <div className="flex items-center gap-2 pt-6 text-sm">
                            <Switch checked={cycleEnabled} onCheckedChange={setCycleEnabled} id="cycle-toggle" />
                            <Label htmlFor="cycle-toggle">Auto-cycle metrics</Label>
                          </div>

                          <div className="pt-6 text-sm">
                            <p className={cn("font-medium", selectedTrend.className && `trend-${selectedTrend.className}`)}>
                              {selectedTrend.text}
                            </p>
                          </div>
                        </div>

                        <div className="chart-container h-[280px] rounded-lg border border-white/10 bg-[#101010]/70 p-2">
                          <canvas ref={chartCanvasRef} />
                        </div>
                      </CardContent>
                    </Card>
                  );
                }

                if (widget.type === "fixed" && widget.view === "clock") {
                  return (
                    <Card
                      key={widget.key}
                      className={cn(
                        "layout-widget flex flex-col overflow-hidden rounded-2xl border-white/10 bg-[#121212]/95 shadow-[0_0_0_1px_rgba(255,255,255,0.02)]",
                        showReveal && "card-reveal-in",
                        draggingWidgetId === widget.key && "card-dragging"
                      )}
                      draggable
                      onDragStart={(event) => onWidgetDragStart(event, widget.key)}
                      style={styleObject}
                    >
                      <CardHeader className="px-4 pb-2 pt-4">
                        <div className="flex items-center justify-between">
                          <CardTitle className="text-base">Clock + Date/Day</CardTitle>
                          <Button size="sm" variant="ghost" onClick={() => handleRemoveWidget(widget.key)}>
                            Remove
                          </Button>
                        </div>
                      </CardHeader>
                      <CardContent className="flex flex-1 items-center justify-center px-4 pb-4 pt-0">
                        <div className="clock-tile-layout">
                          <div className={cn("clock-time-stack", clockParts.isColonDim && "clock-colon-dim")}>
                            <p className="clock-time-main" role="timer" aria-live="polite">
                              <span>{clockParts.hours}</span>
                              <span className="clock-separator">:</span>
                              <span>{clockParts.minutes}</span>
                            </p>
                          </div>
                          <div className="date-tile">
                            <p className="date-tile-top">
                              <span className="font-semibold">{dateTileParts.weekday}</span> {dateTileParts.month}
                            </p>
                            <p className="date-tile-day">{dateTileParts.day}</p>
                          </div>
                        </div>
                      </CardContent>
                    </Card>
                  );
                }

                if (widget.type === "fixed" && widget.view === "quick_links") {
                  return (
                    <Card
                      key={widget.key}
                      className={cn(
                        "layout-widget flex flex-col overflow-hidden rounded-2xl border-white/10 bg-[#121212]/95 shadow-[0_0_0_1px_rgba(255,255,255,0.02)]",
                        showReveal && "card-reveal-in",
                        draggingWidgetId === widget.key && "card-dragging"
                      )}
                      draggable
                      onDragStart={(event) => onWidgetDragStart(event, widget.key)}
                      style={styleObject}
                    >
                      <CardHeader className="px-4 pb-1 pt-3">
                        <div className="flex items-center justify-between gap-2">
                          <div>
                            <CardTitle className="text-base">Quick Links</CardTitle>
                            <CardDescription className="text-xs">Google shortcuts</CardDescription>
                          </div>
                          <Button size="sm" variant="ghost" onClick={() => handleRemoveWidget(widget.key)}>
                            Remove
                          </Button>
                        </div>
                      </CardHeader>
                      <CardContent className="flex flex-1 items-center px-4 pb-2 pt-1">
                        <div className="quick-links-grid">
                          {quickLinkItems.map((item) => {
                            const Icon = item.icon;
                            return (
                              <a
                                key={item.key}
                                href={item.href}
                                target="_blank"
                                rel="noreferrer"
                                className={cn("quick-link-tile", item.toneClass)}
                                aria-label={item.ariaLabel}
                              >
                                <span className="quick-link-icon-shell">
                                  <Icon className="h-5 w-5" strokeWidth={2.2} />
                                </span>
                                <span className="quick-link-label">{item.label}</span>
                              </a>
                            );
                          })}
                        </div>
                      </CardContent>
                    </Card>
                  );
                }

                if (widget.type === "fixed" && widget.view === "calendar") {
                  const showConnect = calendarPreview.status === "auth_required" && Boolean(calendarPreview.authUrl);
                  const showError = calendarPreview.status === "error";
                  const showBackendMissing = calendarPreview.status === "backend_missing";
                  const showLoading = calendarPreview.status === "loading";
                  const events = Array.isArray(calendarPreview.events) ? calendarPreview.events : [];

                  return (
                    <Card
                      key={widget.key}
                      className={cn(
                        "layout-widget rounded-2xl border-white/10 bg-[#121212]/95 shadow-[0_0_0_1px_rgba(255,255,255,0.02)]",
                        showReveal && "card-reveal-in",
                        draggingWidgetId === widget.key && "card-dragging"
                      )}
                      draggable
                      onDragStart={(event) => onWidgetDragStart(event, widget.key)}
                      style={styleObject}
                    >
                      <CardHeader className="pb-2">
                        <div className="flex items-start justify-between gap-2">
                          <div>
                            <CardTitle className="text-base">Upcoming</CardTitle>
                            <CardDescription className="text-xs">
                              {calendarPreview.updatedAt ? `Updated ${calendarPreview.updatedAt}` : "Google Calendar"}
                            </CardDescription>
                          </div>
                          <div className="flex items-center gap-1">
                            {showConnect ? (
                              <Button size="sm" variant="outline" asChild>
                                <a href={calendarPreview.authUrl} target="_blank" rel="noreferrer">
                                  Connect
                                </a>
                              </Button>
                            ) : null}
                            <Button size="sm" variant="ghost" onClick={() => handleRemoveWidget(widget.key)}>
                              Remove
                            </Button>
                          </div>
                        </div>
                      </CardHeader>
                      <CardContent>
                        {showBackendMissing ? <p className="text-sm text-zinc-300">{calendarPreview.error}</p> : null}
                        {showError ? <p className="text-sm text-zinc-300">{calendarPreview.error || "Unable to load Google Calendar."}</p> : null}
                        {showLoading ? <p className="text-sm text-zinc-300">Loading upcoming events...</p> : null}
                        {showConnect ? <p className="mb-2 text-sm text-zinc-300">Authorize Google Calendar to load meetings.</p> : null}
                        {!showBackendMissing && !showError && !showLoading && !events.length ? (
                          <p className="text-sm text-zinc-300">No upcoming meetings.</p>
                        ) : null}
                        {events.length ? (
                          <ul className="calendar-upcoming-list">
                            {events.map((event) => (
                              <li key={event.id} className="calendar-upcoming-item">
                                <span
                                  className={cn(
                                    "calendar-upcoming-marker",
                                    event.markerType === "bar" ? "is-bar" : "is-ring",
                                    event.colorToken === "green" && "is-green",
                                    event.colorToken === "teal" && "is-teal",
                                    event.colorToken === "blue" && "is-blue",
                                    event.colorToken === "gray" && "is-gray"
                                  )}
                                />
                                <div className="min-w-0">
                                  <p className="truncate text-lg font-semibold leading-tight text-zinc-100">{event.title}</p>
                                  <p className="text-sm text-zinc-400">{event.timeText}</p>
                                </div>
                              </li>
                            ))}
                          </ul>
                        ) : null}
                      </CardContent>
                    </Card>
                  );
                }

                if (widget.type === "advanced") {
                  const content = buildAdvancedCellContent(widget.payloadKey, cellsPayload, cfg);
                  const isMetricLikeContent = content && content.mode === "metric_like";
                  return (
                    <Card
                      key={widget.key}
                      className={cn(
                        "layout-widget rounded-2xl border-white/10 bg-[#121212]/95 shadow-[0_0_0_1px_rgba(255,255,255,0.02)]",
                        showReveal && "card-reveal-in",
                        draggingWidgetId === widget.key && "card-dragging"
                      )}
                      draggable
                      onDragStart={(event) => onWidgetDragStart(event, widget.key)}
                      style={styleObject}
                    >
                      <CardHeader className="pb-2">
                        {isMetricLikeContent ? (
                          <div className="flex items-center justify-between">
                            <CardTitle className="text-base font-medium text-zinc-300">{content.title || widget.label || "Advanced Cell"}</CardTitle>
                            <div className="flex items-center gap-2">
                              <span className={cn("rounded-full border px-2.5 py-1 text-xs font-medium", content.trendPillClass || "trend-pill-flat")}>
                                {content.changeText || "--"}
                              </span>
                              <Button
                                size="sm"
                                variant="ghost"
                                className="text-zinc-500 hover:text-zinc-200"
                                onClick={() => handleRemoveWidget(widget.key)}
                              >
                                Remove
                              </Button>
                            </div>
                          </div>
                        ) : (
                          <div className="flex items-center justify-between gap-3">
                            <div>
                              <CardTitle className="text-base leading-tight">{widget.label || "Advanced Cell"}</CardTitle>
                              {content.subtitle ? (
                                <CardDescription className="mt-1 text-xs">{content.subtitle}</CardDescription>
                              ) : null}
                            </div>
                            <Button size="sm" variant="ghost" onClick={() => handleRemoveWidget(widget.key)}>
                              Remove
                            </Button>
                          </div>
                        )}
                      </CardHeader>
                      {isMetricLikeContent ? (
                        <CardContent className="space-y-3">
                          <AnimatedNumber
                            as="p"
                            className="text-[3rem] font-semibold leading-none tracking-tight text-zinc-50"
                            value={content.value}
                            formatValue={(num) => formatMetricLikeValue(num, content, cfg)}
                            fallback={content.valueText || "--"}
                          />
                          <p className={cn("text-xl font-medium", content.trendSummaryClass && `trend-${content.trendSummaryClass}`)}>
                            {content.trendSummary || "Stable this period"}
                          </p>
                          <p className="text-sm text-zinc-400">{content.previousText || "vs previous MTD: --"}</p>
                        </CardContent>
                      ) : (
                        <CardContent>
                          <div className="text-sm text-zinc-200" dangerouslySetInnerHTML={{ __html: content.body }} />
                        </CardContent>
                      )}
                    </Card>
                  );
                }

                const meta = widget.metricMeta;
                if (!meta) {
                  return null;
                }

                const summarySnapshot = getSummaryMetricSnapshot(meta, cellsPayload);
                const value = summarySnapshot && Number.isFinite(summarySnapshot.value)
                  ? summarySnapshot.value
                  : latestRow
                    ? parseNumber(latestRow[meta.valueCol])
                    : NaN;
                const change = summarySnapshot && Number.isFinite(summarySnapshot.change)
                  ? summarySnapshot.change
                  : latestRow && meta.changeCol
                    ? parseNumber(latestRow[meta.changeCol])
                    : NaN;
                const trend = summarySnapshot
                  ? resolveTrend(change, "")
                  : resolveTrend(change, latestRow && meta.directionCol ? latestRow[meta.directionCol] : "");

                const previousActual = summarySnapshot && Number.isFinite(summarySnapshot.previousValue)
                  ? summarySnapshot.previousValue
                  : getPreviousMetricValue(previousRow, meta.valueCol);
                const previousEstimated = estimatePreviousValue(value, change);
                const previousLabel = summarySnapshot ? "vs previous MTD" : "vs previous";
                const previousText = Number.isFinite(previousActual)
                  ? `${previousLabel}: ${formatMetricValue(meta, valueOrDash(previousActual), cfg.currencySymbol)}`
                  : Number.isFinite(previousEstimated)
                    ? `${previousLabel} (estimated): ${formatMetricValue(meta, valueOrDash(previousEstimated), cfg.currencySymbol)}`
                    : `${previousLabel}: --`;

                const sign = change > 0 ? "+" : "";
                const changeText = Number.isFinite(change)
                  ? `${trendArrow(trend)} ${sign}${formatNumber(change, 0)}%`
                  : "--";
                const trendPillClass = trend === "up"
                  ? "trend-pill-up"
                  : trend === "down"
                    ? "trend-pill-down"
                    : "trend-pill-flat";
                const trendSummary = trend === "up"
                  ? "Trending up this month"
                  : trend === "down"
                    ? "Down this period"
                    : "Stable this period";
                const trendSummaryClass = trendClass(trend);

                return (
                  <Card
                    key={widget.key}
                    className={cn(
                      "layout-widget rounded-2xl border-white/10 bg-[#121212]/95 shadow-[0_0_0_1px_rgba(255,255,255,0.02)]",
                      showReveal && "card-reveal-in",
                      draggingWidgetId === widget.key && "card-dragging"
                    )}
                    draggable
                    onDragStart={(event) => onWidgetDragStart(event, widget.key)}
                    style={styleObject}
                  >
                    <CardHeader className="pb-2">
                      <div className="flex items-center justify-between">
                        <CardTitle className="text-base font-medium text-zinc-300">{meta.label}</CardTitle>
                        <div className="flex items-center gap-2">
                          <span className={cn("rounded-full border px-2.5 py-1 text-xs font-medium", trendPillClass)}>
                            {changeText}
                          </span>
                          <Button size="sm" variant="ghost" className="text-zinc-500 hover:text-zinc-200" onClick={() => handleRemoveWidget(widget.key)}>
                            Remove
                          </Button>
                        </div>
                      </div>
                    </CardHeader>
                    <CardContent className="space-y-3">
                      <AnimatedNumber
                        as="p"
                        className="text-[3rem] font-semibold leading-none tracking-tight text-zinc-50"
                        value={value}
                        formatValue={(num) => formatMetricValue(meta, num, cfg.currencySymbol)}
                        fallback="--"
                      />
                      <p className={cn("text-xl font-medium", trendSummaryClass && `trend-${trendSummaryClass}`)}>
                        {trendSummary}
                      </p>
                      <p className="text-sm text-zinc-400">{previousText}</p>
                    </CardContent>
                  </Card>
                );
              })}

              <div ref={dropIndicatorRef} className="metrics-drop-indicator" />
            </div>
          </section>

          <Sheet open={cellMenuOpen} onOpenChange={setCellMenuOpen}>
            <SheetContent side="right" className="sm:max-w-md" container={fullscreenTargetRef.current || undefined}>
              <SheetHeader>
                <SheetTitle>Cell Library</SheetTitle>
                <SheetDescription>Drag any cell into the board. All cells can be moved or removed.</SheetDescription>
              </SheetHeader>
              <ScrollArea className="mt-4 h-[82vh] pr-3">
                <div className="space-y-2">
                  {cellLibrary.length ? (
                    cellLibrary.map((widget) => {
                      const active = isWidgetOnBoard(cardLayout, widget.key);
                      return (
                        <button
                          key={widget.key}
                          type="button"
                          draggable
                          onDragStart={(event) => onLibraryDragStart(event, widget.key)}
                          onClick={() => addOrMoveCard(widget.key, null)}
                          className={cn(
                            "w-full rounded-lg border px-3 py-2 text-left transition hover:border-primary/45 hover:bg-zinc-900/60",
                            active ? "border-primary/45 bg-primary/5" : "border-white/10"
                          )}
                        >
                          <p className="text-sm font-medium">{widget.label}</p>
                          <p className="text-xs text-zinc-400">{active ? "On board" : "Drag to add"}</p>
                        </button>
                      );
                    })
                  ) : (
                    <p className="text-sm text-muted-foreground">No cells found yet.</p>
                  )}
                </div>
              </ScrollArea>
            </SheetContent>
          </Sheet>
        </main>
      )}
    </div>
  );
}

function StatPill({ label, value, numericValue = NaN, formatValue = null, note = "", tone = "neutral" }) {
  const canTween = Number.isFinite(numericValue) && typeof formatValue === "function";
  return (
    <div
      className={cn(
        "flex min-h-[7.5rem] flex-col justify-center rounded-xl border px-4 py-3.5 md:min-h-[8.4rem]",
        tone === "warning" && "border-white/15 bg-white/[0.04]",
        tone === "neutral" && "border-white/10 bg-[#111111]"
      )}
    >
      <p className="text-sm uppercase tracking-[0.08em] text-zinc-400">{label}</p>
      {canTween ? (
        <AnimatedNumber
          as="p"
          className="mt-1 whitespace-nowrap text-[clamp(1.65rem,1.95vw,2.15rem)] font-semibold leading-[1.04] tracking-tight text-zinc-100"
          value={numericValue}
          formatValue={formatValue}
          fallback={value}
        />
      ) : (
        <p className="mt-1 whitespace-nowrap text-[clamp(1.65rem,1.95vw,2.15rem)] font-semibold leading-[1.04] tracking-tight text-zinc-100">{value}</p>
      )}
      {note ? <p className="mt-2 text-sm text-zinc-400">{note}</p> : null}
    </div>
  );
}

function AnimatedNumber({
  value,
  formatValue,
  fallback = "--",
  className = "",
  as: Element = "span",
  durationMs = VALUE_TWEEN_DURATION_MS,
}) {
  const [displayValue, setDisplayValue] = useState(() => (Number.isFinite(value) ? value : NaN));
  const previousValueRef = useRef(Number.isFinite(value) ? value : NaN);
  const hasMountedRef = useRef(false);
  const reduceMotionRef = useRef(false);

  useEffect(() => {
    if (typeof window === "undefined" || typeof window.matchMedia !== "function") {
      return undefined;
    }

    const query = window.matchMedia("(prefers-reduced-motion: reduce)");
    const handleChange = () => {
      reduceMotionRef.current = query.matches;
    };

    handleChange();

    if (typeof query.addEventListener === "function") {
      query.addEventListener("change", handleChange);
      return () => {
        query.removeEventListener("change", handleChange);
      };
    }

    query.addListener(handleChange);
    return () => {
      query.removeListener(handleChange);
    };
  }, []);

  useEffect(() => {
    const nextValue = Number.isFinite(value) ? value : NaN;
    const previousValue = previousValueRef.current;

    if (!Number.isFinite(nextValue)) {
      previousValueRef.current = NaN;
      setDisplayValue(NaN);
      return undefined;
    }

    const skipTween =
      !hasMountedRef.current ||
      reduceMotionRef.current ||
      !Number.isFinite(previousValue) ||
      durationMs <= 0 ||
      Math.abs(nextValue - previousValue) < VALUE_TWEEN_MIN_DELTA;

    if (skipTween) {
      previousValueRef.current = nextValue;
      setDisplayValue(nextValue);
      hasMountedRef.current = true;
      return undefined;
    }

    hasMountedRef.current = true;
    let rafId = 0;
    const startValue = previousValue;
    const delta = nextValue - startValue;
    const startTime = performance.now();

    const animate = (timestamp) => {
      const progress = Math.min(1, (timestamp - startTime) / durationMs);
      const eased = 1 - (1 - progress) ** 3;
      setDisplayValue(startValue + delta * eased);

      if (progress < 1) {
        rafId = window.requestAnimationFrame(animate);
        return;
      }

      previousValueRef.current = nextValue;
      setDisplayValue(nextValue);
    };

    rafId = window.requestAnimationFrame(animate);
    return () => {
      window.cancelAnimationFrame(rafId);
    };
  }, [durationMs, value]);

  const text =
    Number.isFinite(displayValue) && typeof formatValue === "function"
      ? formatValue(displayValue)
      : fallback;

  return <Element className={className}>{text}</Element>;
}

function styleVarsFromPlacement(placementStyle) {
  const obj = {};
  placementStyle
    .split(";")
    .map((part) => part.trim())
    .filter(Boolean)
    .forEach((entry) => {
      const [key, value] = entry.split(":");
      obj[key] = value;
    });
  return obj;
}

function valueOrDash(value) {
  return Number.isFinite(value) ? value : NaN;
}

async function loadAdvancedCells(cfg) {
  if (!cfg.backendApiUrl) {
    return null;
  }

  try {
    const payload = await fetchBackendPayload(cfg, "/api/cells");
    if (payload && typeof payload === "object") {
      if (!payload.kpis || !payload.summary) {
        try {
          const summaryPayload = await fetchBackendPayload(cfg, "/api/summary");
          if (summaryPayload && typeof summaryPayload === "object") {
            payload.summary = payload.summary || summaryPayload.summary || null;
            payload.kpis = payload.kpis || summaryPayload.kpis || null;
            payload.ytd_comparison = payload.ytd_comparison || summaryPayload.ytd || null;
          }
        } catch (_error) {
          // Keep payload as-is.
        }
      }

      if (!payload.website_sessions_mtd) {
        try {
          const sessionsPayload = await fetchBackendPayload(cfg, "/api/sessions/mtd");
          if (sessionsPayload && typeof sessionsPayload === "object") {
            payload.website_sessions_mtd = sessionsPayload;
          }
        } catch (error) {
          const message = error && error.message ? String(error.message) : "";
          const routeMissing = /\(404\)/.test(message);
          payload.website_sessions_mtd = {
            status: "unavailable",
            unavailable_reason: routeMissing
              ? "Backend route /api/sessions/mtd is missing. Restart backend with latest code."
              : error && error.message
                ? `Unable to load /api/sessions/mtd: ${error.message}`
                : "Unable to load /api/sessions/mtd from backend.",
          };
        }
      }
      if (!payload.aov || typeof payload.aov !== "object") {
        try {
          const aovPayload = await fetchBackendPayload(cfg, "/api/aov");
          if (aovPayload && typeof aovPayload === "object") {
            payload.aov = aovPayload;
          }
        } catch (error) {
          const message = error && error.message ? String(error.message) : "";
          const routeMissing = /\(404\)/.test(message);
          payload.aov = {
            status: "unavailable",
            unavailable_reason: routeMissing
              ? "Backend route /api/aov is missing. Restart backend with latest code."
              : error && error.message
                ? `Unable to load /api/aov: ${error.message}`
                : "Unable to load /api/aov from backend.",
          };
        }
      }
      if (!payload.ytd_comparison || typeof payload.ytd_comparison !== "object") {
        try {
          const ytdPayload = await fetchBackendPayload(cfg, "/api/ytd");
          if (ytdPayload && typeof ytdPayload === "object") {
            payload.ytd_comparison = ytdPayload;
          }
        } catch (_error) {
          payload.ytd_comparison = payload.ytd_comparison || null;
        }
      }
      return payload;
    }
  } catch (_error) {
    // Fall back to granular endpoints below so key cells still work.
  }

  const fallbackEndpoints = [
    ["summary_bundle", "/api/summary"],
    ["top_products_units", `/api/products/top-units?limit=${ADVANCED_CELL_LIST_LIMIT}`],
    ["top_products_revenue", `/api/products/top-revenue?limit=${ADVANCED_CELL_LIST_LIMIT}`],
    ["product_momentum", `/api/products/momentum?metric=revenue&limit=${ADVANCED_CELL_LIST_LIMIT}`],
    ["daily_sales_pace", "/api/pace"],
    ["mtd_projection", "/api/projection"],
    ["gross_net_returns", "/api/finance/gross-net-returns"],
    ["aov", "/api/aov"],
    ["ytd_comparison", "/api/ytd"],
    ["website_sessions_mtd", "/api/sessions/mtd"],
    ["new_vs_returning", "/api/customers/new-vs-returning"],
    ["channel_split", "/api/channels"],
    ["discount_impact", "/api/discount-impact"],
    ["hourly_heatmap_today", "/api/heatmap/today"],
    ["refund_watchlist", `/api/refund-watchlist?limit=${ADVANCED_CELL_LIST_LIMIT}`],
  ];

  const settled = await Promise.allSettled(
    fallbackEndpoints.map(([, path]) => fetchBackendPayload(cfg, path))
  );

  const merged = {
    updatedAt: new Date().toISOString(),
  };

  fallbackEndpoints.forEach(([key], index) => {
    const result = settled[index];
    if (result.status !== "fulfilled" || !result.value || typeof result.value !== "object") {
      return;
    }

    if (key === "summary_bundle") {
      merged.summary = result.value.summary || null;
      merged.kpis = result.value.kpis || null;
      merged.ytd_comparison = result.value.ytd || null;
      return;
    }

    merged[key] = result.value;
  });

  const hasUsefulData =
    Boolean(merged.kpis) ||
    Boolean(merged.daily_sales_pace) ||
    Boolean(merged.mtd_projection) ||
    Boolean(merged.aov) ||
    Boolean(merged.top_products_units) ||
    Boolean(merged.top_products_revenue) ||
    Boolean(merged.product_momentum) ||
    Boolean(merged.website_sessions_mtd) ||
    Boolean(merged.ytd_comparison);

  return hasUsefulData ? merged : null;
}

function pickLastMonthShopifyTotal(cellsPayload) {
  const monthTotal = parseNumber(cellsPayload?.daily_sales_pace?.previous_month_total_sales);
  if (Number.isFinite(monthTotal)) {
    return { value: monthTotal, source: "shopify last month total" };
  }

  const monthGross = parseNumber(cellsPayload?.daily_sales_pace?.previous_month_gross_sales);
  if (Number.isFinite(monthGross)) {
    return { value: monthGross, source: "shopify last month gross" };
  }

  const monthNet = parseNumber(cellsPayload?.daily_sales_pace?.previous_month_net_sales);
  if (Number.isFinite(monthNet)) {
    return { value: monthNet, source: "shopify last month net" };
  }

  return { value: NaN, source: "shopify last month unavailable" };
}

async function loadRows(cfg) {
  if (cfg.backendApiUrl) {
    return loadRowsFromPipeline(cfg, { backendOnly: true });
  }

  if (cfg.appsScriptUrl) {
    const pipelineRows = await loadRowsFromPipeline(cfg, { backendOnly: false });
    if (pipelineRows.length) {
      return pipelineRows;
    }

    const payload = await fetchAppsScriptPayload(cfg, `sheet=${encodeURIComponent(cfg.sheetName)}`);
    const data = extractPayloadRows(payload);
    return prepareRows(data.filter(rowHasData));
  }

  const fallback = Array.isArray(sampleData?.data) ? sampleData.data : [];
  return prepareRows(fallback.filter(rowHasData));
}

async function loadRowsFromPipeline(cfg, options = {}) {
  const backendOnly = Boolean(options.backendOnly);

  if (cfg.backendApiUrl) {
    try {
      const payload = await fetchBackendPayload(cfg, "/api/clean");
      const cleanRows = extractPayloadRows(payload);
      if (cleanRows.length) {
        return prepareMtdRowsFromClean(cleanRows);
      }
      if (backendOnly) {
        return [];
      }
    } catch (_error) {
      if (backendOnly) {
        return [];
      }
    }
  }

  if (backendOnly) {
    return [];
  }

  try {
    const payload = await fetchAppsScriptPayload(cfg, "mode=clean");
    const cleanRows = extractPayloadRows(payload);
    if (!cleanRows.length) {
      return [];
    }
    return prepareMtdRowsFromClean(cleanRows);
  } catch (_error) {
    return [];
  }
}

async function fetchBackendPayload(cfg, path) {
  const base = String(cfg.backendApiUrl || "").replace(/\/+$/, "");
  const safePath = path.startsWith("/") ? path : `/${path}`;

  if (safePath !== "/api/sessions/mtd") {
    return fetchBackendPayloadFromBase(base, safePath);
  }

  const tried = new Set();
  const candidates = [base, ...buildSessionEndpointFallbackBases(base)].filter(Boolean);
  let lastError = null;

  for (const candidateBase of candidates) {
    const normalizedBase = String(candidateBase || "").replace(/\/+$/, "");
    if (!normalizedBase || tried.has(normalizedBase)) {
      continue;
    }
    tried.add(normalizedBase);
    try {
      return await fetchBackendPayloadFromBase(normalizedBase, safePath);
    } catch (error) {
      lastError = error;
      const status = Number(error && error.status);
      if (status && status !== 404) {
        break;
      }
    }
  }

  if (lastError) {
    throw lastError;
  }
  throw new Error("Backend request failed");
}

function buildSessionEndpointFallbackBases(base) {
  const defaults = [
    "http://localhost:8787",
    "http://127.0.0.1:8787",
    "http://localhost:3000",
    "http://127.0.0.1:3000",
  ];

  try {
    const parsed = new URL(base);
    if (parsed.protocol.startsWith("http") && (parsed.hostname === "localhost" || parsed.hostname === "127.0.0.1")) {
      const alt = new URL(parsed.toString());
      alt.port = "8787";
      defaults.unshift(alt.origin);
    }
  } catch (_error) {
    // Keep default fallbacks.
  }

  return defaults;
}

async function fetchBackendPayloadFromBase(base, safePath) {
  const cacheBuster = `_=${Date.now()}`;
  const separator = safePath.includes("?") ? "&" : "?";
  const url = `${base}${safePath}${separator}${cacheBuster}`;
  const res = await fetch(url);
  if (!res.ok) {
    const error = new Error(`Backend request failed (${res.status})`);
    error.status = res.status;
    error.url = url;
    throw error;
  }
  return res.json();
}

async function fetchAppsScriptPayload(cfg, query) {
  const separator = cfg.appsScriptUrl.includes("?") ? "&" : "?";
  const url = `${cfg.appsScriptUrl}${separator}${query}&_=${Date.now()}`;
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(`Apps Script request failed (${res.status})`);
  }
  return res.json();
}

async function safeParseJsonResponse(res) {
  const text = await res.text();
  if (!text) {
    return null;
  }
  try {
    return JSON.parse(text);
  } catch (_error) {
    return null;
  }
}

async function loadGoogleCalendarPreview(cfg, options = {}) {
  const limit = Number.isFinite(options.limit) ? Math.max(1, Math.min(10, Math.floor(options.limit))) : CALENDAR_PREVIEW_LIMIT;
  if (!cfg.backendApiUrl) {
    return {
      status: "backend_missing",
      events: [],
      error: "Backend API is required for Google Calendar widgets.",
      authUrl: "",
      updatedAt: "",
    };
  }

  const base = String(cfg.backendApiUrl || "").replace(/\/+$/, "");
  const path = `/api/google/calendar/upcoming?max=${limit}`;
  const url = `${base}${path}&_=${Date.now()}`;

  try {
    const res = await fetch(url);
    const payload = await safeParseJsonResponse(res);
    if (!res.ok) {
      const authRequired = payload?.status === "google_auth_required";
      return {
        status: authRequired ? "auth_required" : "error",
        events: [],
        error: String(payload?.error || `Google Calendar request failed (${res.status})`),
        authUrl: authRequired ? String(payload?.auth_url || "") : "",
        updatedAt: "",
      };
    }

    const events = Array.isArray(payload?.events)
      ? payload.events
          .map((item, index) => normalizeUpcomingMeeting(item, index))
          .filter(Boolean)
      : [];

    return {
      status: "ready",
      events,
      error: "",
      authUrl: "",
      updatedAt: formatLastUpdatedLabel(payload?.updatedAt),
    };
  } catch (error) {
    return {
      status: "error",
      events: [],
      error: String(error && error.message ? error.message : error),
      authUrl: "",
      updatedAt: "",
    };
  }
}

function normalizeUpcomingMeeting(rawEvent, index) {
  if (!rawEvent || typeof rawEvent !== "object") {
    return null;
  }

  const title = String(rawEvent.title || "").trim() || "(Untitled)";
  const start = String(rawEvent.start || "").trim();
  const end = String(rawEvent.end || "").trim();
  const isAllDay = Boolean(rawEvent.is_all_day);
  if (!start) {
    return null;
  }

  const marker = getCalendarMarker(index);
  return {
    id: String(rawEvent.id || `meeting-${index}-${start}`),
    title,
    timeText: formatUpcomingEventTime({ start, end, isAllDay }),
    markerType: marker.type,
    colorToken: marker.colorToken,
  };
}

function getCalendarMarker(index) {
  const palette = ["blue", "green", "teal", "gray"];
  return {
    type: index % 2 === 0 ? "bar" : "ring",
    colorToken: palette[index % palette.length],
  };
}

function toDateSafe(value) {
  const d = new Date(value);
  return Number.isFinite(d.getTime()) ? d : null;
}

function formatUpcomingEventTime({ start, end, isAllDay }) {
  if (isAllDay) {
    return "All day";
  }

  const startDate = toDateSafe(start);
  if (!startDate) {
    return "--";
  }

  const startLabel = calendarEventTimeFormatter.format(startDate);
  const endDate = toDateSafe(end);
  if (!endDate) {
    return startLabel;
  }

  const sameDay =
    startDate.getFullYear() === endDate.getFullYear() &&
    startDate.getMonth() === endDate.getMonth() &&
    startDate.getDate() === endDate.getDate();
  const endLabel = sameDay ? calendarEventTimeFormatter.format(endDate) : calendarEventDayTimeFormatter.format(endDate);
  return `${startLabel} - ${endLabel}`;
}

function formatLastUpdatedLabel(isoText) {
  const ts = toDateSafe(isoText);
  if (!ts) {
    return "";
  }
  return ts.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
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

function rowHasData(row) {
  return Object.values(row).some((value) => value !== null && value !== "");
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

  return catalog;
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
  const base =
    String(label || "metric")
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
  return isMetricWidgetKey(widgetKey) ? widgetKey.slice(METRIC_WIDGET_PREFIX.length) : "";
}

function advancedKeyFromWidgetKey(widgetKey) {
  return isAdvancedWidgetKey(widgetKey) ? widgetKey.slice(ADVANCED_WIDGET_PREFIX.length) : "";
}

function getAvailableWidgetDefinitions(metricCatalog) {
  const fixedWidgets = Object.values(fixedWidgetSeed);
  const metricWidgets = Object.values(metricCatalog).map((meta) => ({
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

function getWidgetDefinition(widgetKey, metricCatalog) {
  if (fixedWidgetSeed[widgetKey]) {
    return fixedWidgetSeed[widgetKey];
  }

  if (isMetricWidgetKey(widgetKey)) {
    const metricKey = metricKeyFromWidgetKey(widgetKey);
    const metricMeta = metricCatalog[metricKey];
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

function isWidgetAvailable(widgetKey, metricCatalog) {
  return Boolean(getWidgetDefinition(widgetKey, metricCatalog));
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
    trimmed === "quick_links_panel" ||
    trimmed === "calendar_panel" ||
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

function getLayoutGridConfig(boardEl) {
  const geometry = getMetricGridGeometry(boardEl);
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

function getWidgetSlotSizeFromSpan(widgetKey, metricCardSpan, metricCatalog) {
  const def = getWidgetDefinition(widgetKey, metricCatalog);
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

function canPlaceWidgetAtSlot(widgetKey, col, row, slotCols, occupiedRects, metricCardSpan, metricCatalog) {
  if (!Number.isFinite(col) || !Number.isFinite(row) || col < 1 || row < 1) {
    return false;
  }
  const size = getWidgetSlotSizeFromSpan(widgetKey, metricCardSpan, metricCatalog);
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

function findFirstAvailableSlot(widgetKey, startCol, startRow, slotCols, occupiedRects, metricCardSpan, metricCatalog) {
  const safeStartCol = clamp(Math.round(startCol || 1), 1, slotCols);
  const safeStartRow = Math.max(1, Math.round(startRow || 1));
  const scanLimit = Math.max(80, safeStartRow + 40);

  for (let row = safeStartRow; row <= scanLimit; row += 1) {
    const firstCol = row === safeStartRow ? safeStartCol : 1;
    for (let col = firstCol; col <= slotCols; col += 1) {
      if (canPlaceWidgetAtSlot(widgetKey, col, row, slotCols, occupiedRects, metricCardSpan, metricCatalog)) {
        return { col, row };
      }
    }
  }

  for (let row = 1; row < safeStartRow; row += 1) {
    for (let col = 1; col <= slotCols; col += 1) {
      if (canPlaceWidgetAtSlot(widgetKey, col, row, slotCols, occupiedRects, metricCardSpan, metricCatalog)) {
        return { col, row };
      }
    }
  }

  return { col: 1, row: scanLimit + 1 };
}

function buildOccupiedRects(layoutEntries, slotCols, metricCardSpan, metricCatalog) {
  return layoutEntries
    .filter((entry) => entry && isWidgetAvailable(entry.key, metricCatalog))
    .map((entry) => {
      const size = getWidgetSlotSizeFromSpan(entry.key, metricCardSpan, metricCatalog);
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

function buildDefaultCardLayout(widgetKeys, slotCols, metricCardSpan, metricCatalog) {
  const layout = [];
  const occupied = [];
  widgetKeys.forEach((key) => {
    const slot = findFirstAvailableSlot(key, 1, 1, slotCols, occupied, metricCardSpan, metricCatalog);
    const size = getWidgetSlotSizeFromSpan(key, metricCardSpan, metricCatalog);
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

function ensureValidCardLayout({ layout, hasCustomLayout, availableWidgetKeys, metricCatalog, slotCols, metricCardSpan }) {
  const availableSet = new Set(availableWidgetKeys);
  const seen = new Set();
  let changed = false;
  const normalized = [];
  const occupied = [];

  layout.forEach((rawEntry) => {
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
      const fallback = findFirstAvailableSlot(entry.key, 1, 1, slotCols, occupied, metricCardSpan, metricCatalog);
      col = fallback.col;
      row = fallback.row;
    }

    if (!canPlaceWidgetAtSlot(entry.key, col, row, slotCols, occupied, metricCardSpan, metricCatalog)) {
      changed = true;
      const fallback = findFirstAvailableSlot(entry.key, col, row, slotCols, occupied, metricCardSpan, metricCatalog);
      col = fallback.col;
      row = fallback.row;
    }

    const size = getWidgetSlotSizeFromSpan(entry.key, metricCardSpan, metricCatalog);
    occupied.push({
      key: entry.key,
      col,
      row,
      colSlots: size.colSlots,
      rowSlots: size.rowSlots,
    });

    normalized.push({ key: entry.key, col, row });
  });

  if (!normalized.length && !hasCustomLayout) {
    const defaults = DEFAULT_CARD_KEYS.filter((key) => isWidgetAvailable(key, metricCatalog));
    const nextLayout = buildDefaultCardLayout(
      defaults.length ? defaults : availableWidgetKeys.slice(0, 6),
      slotCols,
      metricCardSpan,
      metricCatalog
    );
    return { layout: nextLayout, changed: false, fromDefault: true };
  }

  return { layout: normalized, changed, fromDefault: false };
}

function isWidgetOnBoard(layout, widgetKey) {
  return layout.some((item) => item && item.key === widgetKey);
}

function getMetricGridGeometry(boardEl) {
  if (!boardEl) {
    return null;
  }

  const rect = boardEl.getBoundingClientRect();
  if (!rect.width) {
    return null;
  }

  const style = window.getComputedStyle(boardEl);
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

function getWidgetSlotSize(widgetKey, geometry, metricCatalog) {
  const def = getWidgetDefinition(widgetKey, metricCatalog);
  if (!def) {
    return { colSlots: 1, rowSlots: 1 };
  }
  const metricSpan = Math.max(1, geometry.metricCardSpan || 3);
  const colSlots = Math.max(1, Math.round((def.colSpan || metricSpan) / metricSpan));
  const rowSlots = Math.max(1, def.rowSpan || 1);
  return { colSlots, rowSlots };
}

function getChartMetricKeys(cardLayout, metricCatalog) {
  const fromBoard = cardLayout
    .map((entry) => (entry && entry.key ? entry.key : ""))
    .filter((key) => isMetricWidgetKey(key))
    .map((key) => metricKeyFromWidgetKey(key))
    .filter((key) => metricCatalog[key]);

  if (fromBoard.length) {
    return fromBoard;
  }

  return Object.keys(metricCatalog);
}

function getSalesGoalMetric(metricCatalog, cardLayout) {
  if (metricCatalog.sales) {
    return metricCatalog.sales;
  }
  return (
    Object.values(metricCatalog).find((meta) => meta.valueCol === "Order Revenue (Current)") ||
    getChartMetricKeys(cardLayout, metricCatalog)
      .map((key) => metricCatalog[key])
      .find(Boolean) ||
    null
  );
}

function getSummaryMetricSnapshot(meta, cellsPayload) {
  if (!meta || !meta.key || !cellsPayload || !cellsPayload.kpis) {
    return null;
  }

  const current = cellsPayload.kpis.current || {};
  const previous = cellsPayload.kpis.previous || {};
  const change = cellsPayload.kpis.change || {};

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

function formatMetricValue(meta, value, currencySymbol) {
  if (!meta || !Number.isFinite(value)) {
    return "--";
  }

  if (meta.formatType === "currency") {
    return formatCurrency(value, currencySymbol);
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

function buildGoalData({ cfg, row, previousRow, salesMeta, cellsPayload, monthlyTargets }) {
  const empty = {
    currentText: "--",
    currentValue: NaN,
    currentNote: "MTD previous: --",
    targetText: "--",
    targetValue: NaN,
    targetNote: salesMeta ? "gap to target: set target" : "gap to target: --",
    previousText: "--",
    previousValue: NaN,
    previousNote: "",
    statusText: salesMeta ? "Set monthly target" : "Status --",
    statusClass: "text-zinc-400",
    progressText: "--",
    progressValue: NaN,
    progressClamped: 0,
    fillClass: "",
    milestonePrev: "Pace checkpoint today: --",
    milestoneTarget: "Last month vs target: --",
    gapPrev: "Gap vs last month total: --",
    gapTarget: salesMeta ? "Gap to target: set target" : "Gap to target: --",
    gapPrevClass: "text-zinc-300",
    gapTargetClass: salesMeta ? "text-zinc-300" : "text-zinc-300",
  };

  if (!row || !salesMeta) {
    return empty;
  }

  const backendConfigured = Boolean(cfg.backendApiUrl);
  const backendPaceMtdTotalSales = parseNumber(cellsPayload?.daily_sales_pace?.mtd_sales);
  const backendProjectedMtdTotalSales = parseNumber(cellsPayload?.mtd_projection?.mtd_sales);
  const backendFallbackMtdGross = parseNumber(cellsPayload?.mtd_projection?.mtd_gross_sales);
  const rowCurrent = parseNumber(row[salesMeta.valueCol]);

  const current = Number.isFinite(backendPaceMtdTotalSales)
    ? backendPaceMtdTotalSales
    : Number.isFinite(backendProjectedMtdTotalSales)
      ? backendProjectedMtdTotalSales
      : Number.isFinite(backendFallbackMtdGross)
        ? backendFallbackMtdGross
        : backendConfigured
          ? NaN
          : rowCurrent;

  const change = parseNumber(row[salesMeta.changeCol]);
  const previousFromRows = getPreviousMetricValue(previousRow, salesMeta.valueCol);
  const previousFromRowsEstimated = estimatePreviousValue(current, change);
  const shopifyPrevious = pickLastMonthShopifyTotal(cellsPayload);

  const previous = backendConfigured
    ? shopifyPrevious.value
    : Number.isFinite(previousFromRows)
      ? previousFromRows
      : previousFromRowsEstimated;

  const previousComparableMtd = parseNumber(cellsPayload?.kpis?.previous?.sales_amount);

  const previousSource = backendConfigured
    ? shopifyPrevious.source
    : Number.isFinite(previousFromRows)
      ? "prior point"
      : "estimated";

  const monthKey = getMonthKey(new Date());
  const multiplier = Number.isFinite(cfg.salesTargetMultiplier) && cfg.salesTargetMultiplier > 0 ? cfg.salesTargetMultiplier : 1;
  const monthlyStoredTarget = getMonthlyTargetForKey(monthlyTargets, monthKey);
  const monthlyConfigTarget = getConfigTargetForKey(cfg.salesTargetsByMonth, monthKey);
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
    return {
      ...empty,
      currentText: Number.isFinite(current) ? formatCurrency(current, cfg.currencySymbol) : "--",
      currentValue: current,
      currentNote: Number.isFinite(previousComparableMtd)
        ? `MTD previous: ${formatCurrency(previousComparableMtd, cfg.currencySymbol)}`
        : "MTD previous: --",
      previousText: Number.isFinite(previous) ? formatCurrency(previous, cfg.currencySymbol) : "--",
      targetValue: target,
      previousValue: previous,
      previousNote: previousSource,
    };
  }

  const progressPct = (current / target) * 100;
  const progressClamped = clampPercent(progressPct);
  const delta = current - target;
  const toTarget = target - current;
  const beatTarget = current >= target;
  const deltaVsLastMonth = Number.isFinite(previous) ? current - previous : NaN;
  const growthVsLastMonthPct = Number.isFinite(previous) && previous !== 0 ? ((current - previous) / Math.abs(previous)) * 100 : NaN;

  const paceDaysElapsed = parseNumber(cellsPayload?.daily_sales_pace?.days_elapsed);
  const paceDaysRemaining = parseNumber(cellsPayload?.daily_sales_pace?.days_remaining);
  const projectionDaysInMonth = parseNumber(cellsPayload?.mtd_projection?.days_in_month);
  const projectionDaysElapsed = parseNumber(cellsPayload?.mtd_projection?.days_elapsed);

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

  const expectedProgressPct = Number.isFinite(totalDays) && totalDays > 0 && Number.isFinite(elapsedDays)
    ? (elapsedDays / totalDays) * 100
    : NaN;

  const onPace = Number.isFinite(expectedProgressPct) ? progressPct >= expectedProgressPct : current >= target * 0.5;

  let statusText = "On pace";
  let statusClass = "text-zinc-200";
  let fillClass = "on-track";
  let gapPrevClass = "text-zinc-300";
  let gapTargetClass = "text-zinc-300";

  if (beatTarget) {
    statusText = "Target achieved";
    statusClass = "text-zinc-100";
    fillClass = "complete";
    gapPrevClass = "text-zinc-200";
    gapTargetClass = "text-zinc-200";
  } else if (!onPace) {
    statusText = "Behind pace";
    statusClass = "text-zinc-200";
    fillClass = "below-prev";
    gapTargetClass = "text-zinc-300";
    gapPrevClass = "text-zinc-300";
  }

  const gapPrev = Number.isFinite(deltaVsLastMonth)
    ? `${deltaVsLastMonth >= 0 ? "Gap vs last month total: +" : "Gap vs last month total: -"}${formatCurrency(
        Math.abs(deltaVsLastMonth),
        cfg.currencySymbol
      )} (${formatPercentSafe(growthVsLastMonthPct)})`
    : "Gap vs last month total: --";

  const gapTarget = beatTarget
    ? `Gap to target: +${formatCurrency(Math.abs(delta), cfg.currencySymbol)}`
    : `Gap to target: -${formatCurrency(Math.abs(toTarget), cfg.currencySymbol)}`;

  const currentNote = Number.isFinite(previousComparableMtd)
    ? `MTD previous: ${formatCurrency(previousComparableMtd, cfg.currencySymbol)}`
    : "MTD previous: --";

  const targetNote = beatTarget
    ? `gap to target: +${formatCurrency(Math.abs(delta), cfg.currencySymbol)}`
    : `gap to target: -${formatCurrency(Math.abs(toTarget), cfg.currencySymbol)}`;

  return {
    currentText: formatCurrency(current, cfg.currencySymbol),
    currentValue: current,
    currentNote,
    targetText: formatCurrency(target, cfg.currencySymbol),
    targetValue: target,
    targetNote,
    previousText: Number.isFinite(previous) ? formatCurrency(previous, cfg.currencySymbol) : "--",
    previousValue: previous,
    previousNote: previousSource,
    statusText,
    statusClass,
    progressText: `${formatNumber(progressPct, 1)}%`,
    progressValue: progressPct,
    progressClamped,
    fillClass,
    milestonePrev:
      Number.isFinite(expectedProgressPct) && Number.isFinite(elapsedDays) && Number.isFinite(totalDays)
        ? `Pace checkpoint today: ${formatNumber(expectedProgressPct, 1)}% (${formatNumber(elapsedDays, 0)}/${formatNumber(totalDays, 0)} days)`
        : "Pace checkpoint today: --",
    milestoneTarget: Number.isFinite(previous)
      ? `Last month vs target: ${formatNumber((previous / target) * 100, 1)}%`
      : "Last month vs target: --",
    gapPrev,
    gapTarget,
    gapPrevClass,
    gapTargetClass,
  };
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
  return "flat";
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

function getChartTypography(viewportWidth) {
  if (viewportWidth >= 3840) {
    return { legendSize: 18, tickSize: 16 };
  }
  if (viewportWidth >= 2560) {
    return { legendSize: 15, tickSize: 14 };
  }
  return { legendSize: 12, tickSize: 11 };
}

function getQueryPasscode() {
  try {
    const params = new URLSearchParams(window.location.search);
    return normalizePasscodeForCompare(params.get("unlock") || "");
  } catch (_error) {
    return "";
  }
}

function normalizeConfig(input) {
  const passcodeValue = typeof input.passcode === "undefined" || input.passcode === null ? "" : String(input.passcode).trim();
  const backendApiUrl = normalizeBackendApiUrl(input.backendApiUrl);
  const targetMultiplierRaw = parseNumber(input.salesTargetMultiplier);
  const targetMultiplier = Number.isFinite(targetMultiplierRaw) && targetMultiplierRaw > 0 ? targetMultiplierRaw : 1;
  const targetValueRaw = parseNumber(input.salesTargetValue);
  const targetsByMonth = normalizeSalesTargetsByMonth(input.salesTargetsByMonth);
  return {
    ...input,
    backendApiUrl,
    currencySymbol: DASHBOARD_CURRENCY,
    passcode: passcodeValue,
    salesTargetMultiplier: targetMultiplier,
    salesTargetValue: Number.isFinite(targetValueRaw) ? targetValueRaw : null,
    salesTargetsByMonth: targetsByMonth,
  };
}

function normalizeBackendApiUrl(value) {
  const configured = typeof value === "string" ? value.trim() : String(value || "").trim();
  const browserOrigin = getBrowserOrigin();

  if (!configured) {
    if (!browserOrigin) {
      return "";
    }
    try {
      const current = new URL(browserOrigin);
      return isLocalHostname(current.hostname) ? "" : browserOrigin;
    } catch (_error) {
      return "";
    }
  }

  if (!/^https?:\/\//i.test(configured)) {
    return configured;
  }

  if (!browserOrigin) {
    return configured;
  }

  try {
    const configuredUrl = new URL(configured);
    const currentUrl = new URL(browserOrigin);
    if (!isLocalHostname(currentUrl.hostname) && isLocalHostname(configuredUrl.hostname)) {
      return browserOrigin;
    }
    return configured;
  } catch (_error) {
    return configured;
  }
}

function getBrowserOrigin() {
  if (typeof window === "undefined" || !window.location || !window.location.origin) {
    return "";
  }
  return String(window.location.origin || "").trim();
}

function isLocalHostname(hostname) {
  const value = String(hostname || "").trim().toLowerCase();
  return value === "localhost" || value === "127.0.0.1" || value === "[::1]" || value === "::1";
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

function getLabel(row) {
  const ts = row && row.__ts instanceof Date ? row.__ts : getRowTimestamp(row);
  if (ts instanceof Date && Number.isFinite(ts.getTime())) {
    return formatTimestampLabel(ts);
  }
  return row["Current Date Range"] || row.MONTH || row.Month || row.Date || row.Timestamp || "Point";
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

function getDateTileDisplayParts(date) {
  const safeDate = date instanceof Date && Number.isFinite(date.getTime()) ? date : new Date();
  return {
    weekday: dateTileWeekdayFormatter.format(safeDate),
    month: dateTileMonthFormatter.format(safeDate),
    day: dateTileDayFormatter.format(safeDate),
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

function saveCardLayout(layout) {
  try {
    window.localStorage.setItem(LAYOUT_STORAGE_KEY, JSON.stringify(layout));
  } catch (_error) {
    // ignore storage errors
  }
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

function getMonthlyTargetForKey(monthlyTargets, monthKey) {
  if (!monthKey || !monthlyTargets || typeof monthlyTargets !== "object") {
    return NaN;
  }
  return parseNumber(monthlyTargets[monthKey]);
}

function getConfigTargetForKey(configTargetsByMonth, monthKey) {
  if (!monthKey || !configTargetsByMonth || typeof configTargetsByMonth !== "object") {
    return NaN;
  }
  return parseNumber(configTargetsByMonth[monthKey]);
}

function formatCurrencySafe(value, cfg) {
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

function formatMetricLikeValue(value, content, cfg) {
  const digits = Number.isFinite(content?.valueDigits) ? content.valueDigits : 0;
  if (content?.valueFormat === "currency") {
    return formatCurrency(value, cfg.currencySymbol);
  }
  if (content?.valueFormat === "percent") {
    const sign = content?.signed && value > 0 ? "+" : "";
    return `${sign}${formatNumber(value, digits)}%`;
  }
  const sign = content?.signed && value > 0 ? "+" : "";
  return `${sign}${formatNumber(value, digits)}`;
}

function escapeHtml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildAdvancedCellContent(cellKey, payload, cfg) {
  if (!payload || typeof payload !== "object") {
    return {
      subtitle: "",
      body: '<p class="text-sm text-zinc-400">Advanced data unavailable. Start backend API and run Supabase sync.</p>',
    };
  }

  const sectionPayload = payload[cellKey];

  if (cellKey === "top_products_units") {
    return buildTopProductsUnitsContent(sectionPayload, cfg);
  }
  if (cellKey === "top_products_revenue") {
    return buildTopProductsRevenueContent(sectionPayload, cfg);
  }
  if (cellKey === "product_momentum") {
    return buildProductMomentumContent(sectionPayload, cfg);
  }
  if (cellKey === "daily_sales_pace") {
    return buildDailyPaceContent(sectionPayload, cfg);
  }
  if (cellKey === "mtd_projection") {
    return buildProjectionContent(sectionPayload, cfg);
  }
  if (cellKey === "gross_net_returns") {
    return buildGrossNetReturnsContent(sectionPayload, cfg);
  }
  if (cellKey === "aov") {
    return buildAovContent(sectionPayload, cfg, payload);
  }
  if (cellKey === "website_sessions_mtd") {
    return buildWebsiteSessionsMtdContent(sectionPayload, payload);
  }
  if (cellKey === "ytd_orders") {
    return buildYtdOrdersContent(payload, cfg);
  }
  if (cellKey === "ytd_total_sales") {
    return buildYtdTotalSalesContent(payload, cfg);
  }
  if (cellKey === "ytd_growth_rate") {
    return buildYtdGrowthRateContent(payload, cfg);
  }
  if (cellKey === "new_vs_returning") {
    return buildNewVsReturningContent(sectionPayload, cfg);
  }
  if (cellKey === "channel_split") {
    return buildChannelSplitContent(sectionPayload, cfg);
  }
  if (cellKey === "discount_impact") {
    return buildDiscountImpactContent(sectionPayload, cfg);
  }
  if (cellKey === "hourly_heatmap_today") {
    return buildHeatmapContent(sectionPayload, cfg);
  }
  if (cellKey === "refund_watchlist") {
    return buildRefundWatchlistContent(sectionPayload, cfg);
  }

  return {
    subtitle: "",
    body: '<p class="text-sm text-zinc-400">Cell is not configured.</p>',
  };
}

function buildTopProductsUnitsContent(payload, cfg) {
  const products = payload && Array.isArray(payload.products) ? payload.products : [];
  const subtitle = payload && Number.isFinite(payload.total_units) ? `Total units: ${formatNumber(payload.total_units, 0)}` : "";
  return {
    subtitle,
    body: renderProductList(products, (item) => [
      `${formatNumberSafe(item.units, 0)} units`,
      formatCurrencySafe(item.revenue, cfg),
      formatPercentSafe(item.unit_share_pct),
    ]),
  };
}

function buildTopProductsRevenueContent(payload, cfg) {
  const products = payload && Array.isArray(payload.products) ? payload.products : [];
  const subtitle = payload && Number.isFinite(payload.total_revenue) ? `Total revenue: ${formatCurrencySafe(payload.total_revenue, cfg)}` : "";
  return {
    subtitle,
    body: renderProductList(products, (item) => [
      formatCurrencySafe(item.revenue, cfg),
      `${formatNumberSafe(item.units, 0)} units`,
      formatPercentSafe(item.revenue_share_pct),
    ]),
  };
}

function buildProductMomentumContent(payload, cfg) {
  const products = payload && Array.isArray(payload.products) ? payload.products : [];
  const metric = payload && payload.metric ? String(payload.metric) : "revenue";
  const body = !products.length
    ? '<p class="text-sm text-zinc-400">No momentum data yet.</p>'
    : `<ul class="space-y-2">${products
        .slice(0, ADVANCED_CELL_LIST_LIMIT)
        .map((item) => {
          const delta = Number.isFinite(item.delta) ? item.delta : null;
          const chipClass = delta > 0 ? "trend-chip-up" : delta < 0 ? "trend-chip-down" : "trend-chip-flat";
          const deltaText =
            metric === "units"
              ? `${delta > 0 ? "+" : ""}${formatNumberSafe(delta, 0)}`
              : `${delta > 0 ? "+" : ""}${formatCurrencySafe(delta, cfg)}`;
          return `
            <li class="rounded-md border border-white/10 p-2">
              <div class="flex items-center justify-between gap-2">
                <span class="font-medium">${escapeHtml(item.title || "Unknown")}</span>
                <span class="rounded-full px-2 py-0.5 text-xs ${chipClass}">${escapeHtml(deltaText)}</span>
              </div>
              <div class="mt-1 flex flex-wrap gap-2 text-xs text-zinc-400">
                <span>Current: ${metric === "units" ? formatNumberSafe(item.current_value, 0) : formatCurrencySafe(item.current_value, cfg)}</span>
                <span>Prev: ${metric === "units" ? formatNumberSafe(item.previous_value, 0) : formatCurrencySafe(item.previous_value, cfg)}</span>
              </div>
            </li>
          `;
        })
        .join("")}</ul>`;
  return { subtitle: `Metric: ${metric}`, body };
}

function buildDailyPaceContent(payload, cfg) {
  const mtdSales = parseNumber(payload && payload.mtd_sales);
  const prevMonthTotal = parseNumber(payload && payload.previous_month_total_sales);
  return {
    subtitle: "",
    body: renderKeyValueList([
      ["Month Goal", formatCurrencySafe(payload && payload.month_goal, cfg)],
      ["MTD Total Sales", formatCurrencySafe(Number.isFinite(mtdSales) ? mtdSales : payload && payload.mtd_gross_sales, cfg)],
      ["Last Month Total", formatCurrencySafe(prevMonthTotal, cfg)],
      ["Today Sales", formatCurrencySafe(payload && payload.today_sales, cfg)],
      ["Required Daily Pace", formatCurrencySafe(payload && payload.required_daily_pace, cfg)],
      ["Days Remaining", formatNumberSafe(payload && payload.days_remaining, 0)],
    ]),
  };
}

function buildProjectionContent(payload, cfg) {
  const mtdSales = parseNumber(payload && payload.mtd_sales);
  return {
    subtitle: "",
    body: renderKeyValueList([
      ["MTD Total Sales", formatCurrencySafe(Number.isFinite(mtdSales) ? mtdSales : payload && payload.mtd_gross_sales, cfg)],
      ["Month Goal", formatCurrencySafe(payload && payload.month_goal, cfg)],
      ["Progress", formatPercentSafe(payload && payload.progress_pct_of_target)],
      ["Projected Month-End", formatCurrencySafe(payload && payload.projected_month_end_sales, cfg)],
      ["Projected vs Target", formatPercentSafe(payload && payload.projected_vs_target_pct)],
    ]),
  };
}

function buildGrossNetReturnsContent(payload, cfg) {
  return {
    subtitle: "",
    body: renderKeyValueList([
      ["Gross Sales", formatCurrencySafe(payload && payload.gross_sales, cfg)],
      ["Net Sales", formatCurrencySafe(payload && payload.net_sales, cfg)],
      ["Total Sales", formatCurrencySafe(payload && payload.total_sales, cfg)],
      ["Returns", formatCurrencySafe(payload && payload.returns_amount, cfg)],
      ["Returns % of Net", formatPercentSafe(payload && payload.returns_rate_pct_of_net)],
    ]),
  };
}

function buildAovContent(payload, cfg, rootPayload) {
  const payloadMtdAov = parseNumber(payload && payload.mtd_aov);
  const payloadPreviousAov = parseNumber(payload && payload.previous_period_aov);
  const payloadChangePct = parseNumber(payload && payload.aov_change_pct);
  const payloadMtdOrders = parseNumber(payload && payload.mtd_orders);

  const summaryCurrentAov = parseNumber(rootPayload && rootPayload.kpis && rootPayload.kpis.current && rootPayload.kpis.current.aov);
  const summaryPreviousAov = parseNumber(rootPayload && rootPayload.kpis && rootPayload.kpis.previous && rootPayload.kpis.previous.aov);
  const summaryChangePct = parseNumber(rootPayload && rootPayload.kpis && rootPayload.kpis.change && rootPayload.kpis.change.aov_pct);
  const summaryCurrentOrders = parseNumber(rootPayload && rootPayload.kpis && rootPayload.kpis.current && rootPayload.kpis.current.orders);

  const mtdAov = Number.isFinite(payloadMtdAov) ? payloadMtdAov : summaryCurrentAov;
  const previousAov = Number.isFinite(payloadPreviousAov) ? payloadPreviousAov : summaryPreviousAov;
  const aovChangePct = Number.isFinite(payloadChangePct) ? payloadChangePct : summaryChangePct;
  const mtdOrders = Number.isFinite(payloadMtdOrders) ? payloadMtdOrders : summaryCurrentOrders;
  const unavailableReason = String(payload && payload.unavailable_reason ? payload.unavailable_reason : "").trim();
  const body = renderKeyValueList([
    ["MTD AOV", formatCurrencySafe(mtdAov, cfg)],
    ["Previous AOV", formatCurrencySafe(previousAov, cfg)],
    ["AOV Change", formatPercentSafe(aovChangePct)],
    ["MTD Orders", formatNumberSafe(mtdOrders, 0)],
  ]);

  const hasAnyAovData =
    Number.isFinite(mtdAov) ||
    Number.isFinite(previousAov) ||
    Number.isFinite(aovChangePct) ||
    Number.isFinite(mtdOrders);

  return {
    subtitle: "",
    body:
      unavailableReason && !hasAnyAovData
        ? `${body}<p class="mt-2 text-xs text-zinc-400">${escapeHtml(unavailableReason)}</p>`
        : body,
  };
}

function buildWebsiteSessionsMtdContent(payload, _rootPayload) {
  const hasPayloadObject = payload && typeof payload === "object";
  const payloadMtd = parseNumber(payload && payload.mtd_sessions);
  const payloadPrev = parseNumber(payload && payload.previous_mtd_sessions);
  const payloadDelta = parseNumber(payload && payload.sessions_change);
  const payloadDeltaPct = parseNumber(payload && payload.sessions_change_pct);
  const unavailableReason = String(payload && payload.unavailable_reason ? payload.unavailable_reason : "").trim();
  const hasCurrent = Number.isFinite(payloadMtd);
  const hasPrevious = Number.isFinite(payloadPrev);
  const derivedDelta = hasCurrent && hasPrevious ? payloadMtd - payloadPrev : NaN;
  const derivedDeltaPct = hasCurrent && hasPrevious ? calcPercentChange(payloadMtd, payloadPrev) : NaN;
  const sessionsDelta = Number.isFinite(payloadDelta) ? payloadDelta : derivedDelta;
  const sessionsDeltaPct = Number.isFinite(payloadDeltaPct) ? payloadDeltaPct : derivedDeltaPct;
  const unavailable = !hasCurrent;

  const mtdSessions = payloadMtd;
  const previousMtdSessions = payloadPrev;

  const trendBasis = Number.isFinite(sessionsDeltaPct) ? sessionsDeltaPct : sessionsDelta;
  const trend = resolveTrend(trendBasis, "");
  const trendPillClass = trend === "up" ? "trend-pill-up" : trend === "down" ? "trend-pill-down" : "trend-pill-flat";
  const trendSummary = unavailable
    ? "Sessions data unavailable"
    : trend === "up"
      ? "Trending up this month"
      : trend === "down"
        ? "Down this period"
        : "Stable this period";
  const trendSummaryClass = trendClass(trend);

  const pctSign = Number.isFinite(sessionsDeltaPct) && sessionsDeltaPct > 0 ? "+" : "";
  const deltaSign = Number.isFinite(sessionsDelta) && sessionsDelta > 0 ? "+" : "";
  const changeText = unavailable
    ? "--"
    : Number.isFinite(sessionsDeltaPct)
    ? `${trendArrow(trend)} ${pctSign}${formatNumber(sessionsDeltaPct, 0)}%`
    : Number.isFinite(sessionsDelta)
      ? `${trendArrow(trend)} ${deltaSign}${formatNumber(sessionsDelta, 0)}`
      : "--";
  const previousText = Number.isFinite(previousMtdSessions)
    ? `vs previous MTD: ${formatNumber(previousMtdSessions, 0)}`
    : unavailableReason || (hasPayloadObject ? "vs previous MTD: --" : "Sessions payload missing from backend.");

  return {
    mode: "metric_like",
    title: "Website Sessions",
    value: mtdSessions,
    valueDigits: 0,
    valueText: formatNumberSafe(mtdSessions, 0),
    trendPillClass,
    changeText,
    trendSummary,
    trendSummaryClass,
    previousText,
    subtitle: "",
    body: "",
  };
}

function getYtdComparisonSnapshot(rootPayload) {
  const payload = rootPayload && typeof rootPayload.ytd_comparison === "object" ? rootPayload.ytd_comparison : null;
  const currentSales = parseNumber(payload?.current?.sales_amount);
  const previousSales = parseNumber(payload?.previous?.sales_amount);
  const currentOrders = parseNumber(payload?.current?.orders);
  const previousOrders = parseNumber(payload?.previous?.orders);
  const salesPctRaw = parseNumber(payload?.change?.sales_amount_pct);
  const ordersPctRaw = parseNumber(payload?.change?.orders_pct);
  const growthPctRaw = parseNumber(payload?.change?.growth_rate_pct);

  return {
    currentSales,
    previousSales,
    currentOrders,
    previousOrders,
    salesPct: Number.isFinite(salesPctRaw) ? salesPctRaw : calcPercentChange(currentSales, previousSales),
    ordersPct: Number.isFinite(ordersPctRaw) ? ordersPctRaw : calcPercentChange(currentOrders, previousOrders),
    growthPct: Number.isFinite(growthPctRaw) ? growthPctRaw : calcPercentChange(currentSales, previousSales),
  };
}

function buildYtdMetricLikeContent({
  title,
  value,
  valueFormat,
  valueDigits = 0,
  signed = false,
  changePct,
  previousText,
  unavailableSummary,
  trendSummary,
}) {
  const unavailable = !Number.isFinite(value);
  const trend = resolveTrend(changePct, "");
  const trendPillClass = trend === "up" ? "trend-pill-up" : trend === "down" ? "trend-pill-down" : "trend-pill-flat";
  const changeSign = Number.isFinite(changePct) && changePct > 0 ? "+" : "";
  const changeText = Number.isFinite(changePct) ? `${trendArrow(trend)} ${changeSign}${formatNumber(changePct, 0)}%` : "--";
  const defaultSummary = unavailable
    ? unavailableSummary || "YTD data unavailable"
    : trend === "up"
      ? "Up vs last year YTD"
      : trend === "down"
        ? "Down vs last year YTD"
        : "Flat vs last year YTD";

  return {
    mode: "metric_like",
    title,
    value,
    valueFormat,
    valueDigits,
    signed,
    valueText: "--",
    trendPillClass,
    changeText,
    trendSummary: trendSummary || defaultSummary,
    trendSummaryClass: trendClass(trend),
    previousText: previousText || "LY YTD: --",
    subtitle: "",
    body: "",
  };
}

function buildYtdOrdersContent(rootPayload, _cfg) {
  const snapshot = getYtdComparisonSnapshot(rootPayload);
  return buildYtdMetricLikeContent({
    title: "Orders YTD",
    value: snapshot.currentOrders,
    valueFormat: "count",
    valueDigits: 0,
    changePct: snapshot.ordersPct,
    previousText: Number.isFinite(snapshot.previousOrders) ? `LY YTD: ${formatNumber(snapshot.previousOrders, 0)}` : "LY YTD: --",
    unavailableSummary: "YTD orders unavailable",
  });
}

function buildYtdTotalSalesContent(rootPayload, cfg) {
  const snapshot = getYtdComparisonSnapshot(rootPayload);
  return buildYtdMetricLikeContent({
    title: "Total Sales YTD",
    value: snapshot.currentSales,
    valueFormat: "currency",
    valueDigits: 0,
    changePct: snapshot.salesPct,
    previousText: Number.isFinite(snapshot.previousSales)
      ? `LY YTD: ${formatCurrency(snapshot.previousSales, cfg.currencySymbol)}`
      : "LY YTD: --",
    unavailableSummary: "YTD sales unavailable",
  });
}

function buildYtdGrowthRateContent(rootPayload, cfg) {
  const snapshot = getYtdComparisonSnapshot(rootPayload);
  return buildYtdMetricLikeContent({
    title: "Growth Rate YTD",
    value: snapshot.growthPct,
    valueFormat: "percent",
    valueDigits: 0,
    signed: true,
    changePct: snapshot.growthPct,
    previousText:
      Number.isFinite(snapshot.currentSales) || Number.isFinite(snapshot.previousSales)
        ? `Current / LY YTD: ${formatCurrencySafe(snapshot.currentSales, cfg)} / ${formatCurrencySafe(snapshot.previousSales, cfg)}`
        : "Current / LY YTD: --",
    unavailableSummary: "YTD growth unavailable",
    trendSummary: "Revenue growth vs last year YTD",
  });
}

function buildNewVsReturningContent(payload, cfg) {
  const revenue = payload && payload.revenue ? payload.revenue : {};
  const shares = payload && payload.shares_pct ? payload.shares_pct : {};
  return {
    subtitle: "",
    body: renderKeyValueList([
      ["New Revenue", `${formatCurrencySafe(revenue.new, cfg)} (${formatPercentSafe(shares.new)})`],
      ["Returning Revenue", `${formatCurrencySafe(revenue.returning, cfg)} (${formatPercentSafe(shares.returning)})`],
      ["Unknown Revenue", `${formatCurrencySafe(revenue.unknown, cfg)} (${formatPercentSafe(shares.unknown)})`],
    ]),
  };
}

function buildChannelSplitContent(payload, cfg) {
  const channels = payload && Array.isArray(payload.channels) ? payload.channels : [];
  const body = !channels.length
    ? '<p class="text-sm text-zinc-400">No channel data yet.</p>'
    : `<ul class="space-y-2">${channels
        .slice(0, ADVANCED_CELL_LIST_LIMIT)
        .map(
          (item) => `
            <li class="rounded-md border border-white/10 p-2">
              <div class="flex items-center justify-between gap-2">
                <span class="font-medium">${escapeHtml(item.channel || "unknown")}</span>
                <span>${formatCurrencySafe(item.revenue, cfg)}</span>
              </div>
              <div class="mt-1 flex flex-wrap gap-2 text-xs text-zinc-400">
                <span>Orders: ${formatNumberSafe(item.orders, 0)}</span>
                <span>${renderTrendAwareInlineText(`Share: ${formatPercentSafe(item.revenue_share_pct)}`)}</span>
              </div>
            </li>
          `
        )
        .join("")}</ul>`;

  return { subtitle: "", body };
}

function buildDiscountImpactContent(payload, cfg) {
  return {
    subtitle: "",
    body: renderKeyValueList([
      ["Discounted Orders %", formatPercentSafe(payload && payload.discounted_orders_pct)],
      ["Total Discounts", formatCurrencySafe(payload && payload.total_discounts, cfg)],
      ["Avg Discount / Order", formatCurrencySafe(payload && payload.avg_discount_per_order, cfg)],
      ["Discount % of Gross", formatPercentSafe(payload && payload.discount_rate_pct_of_gross)],
    ]),
  };
}

function buildHeatmapContent(payload) {
  const heatmap = payload && Array.isArray(payload.heatmap) ? payload.heatmap : [];
  const body = !heatmap.length
    ? '<p class="text-sm text-zinc-400">No hourly data yet.</p>'
    : `<div class="grid grid-cols-2 gap-2">${heatmap
        .map(
          (item) => `
            <div class="rounded-md border border-white/10 p-2 text-xs">
              <div class="font-medium">${escapeHtml(item.hour_utc)}:00</div>
              <div class="text-zinc-400">${formatNumberSafe(item.orders, 0)} ord</div>
            </div>
          `
        )
        .join("")}</div>`;

  return { subtitle: "", body };
}

function buildRefundWatchlistContent(payload, cfg) {
  const products = payload && Array.isArray(payload.products) ? payload.products : [];
  const body = !products.length
    ? '<p class="text-sm text-zinc-400">No significant refunds this month.</p>'
    : `<ul class="space-y-2">${products
        .slice(0, ADVANCED_CELL_LIST_LIMIT)
        .map(
          (item) => `
            <li class="rounded-md border border-white/10 p-2">
              <div class="flex items-center justify-between gap-2">
                <span class="font-medium">${escapeHtml(item.title || "Unknown")}</span>
                <span>${formatPercentSafe(item.return_rate_pct)}</span>
              </div>
              <div class="mt-1 flex flex-wrap gap-2 text-xs text-zinc-400">
                <span>Returned units: ${formatNumberSafe(item.returned_units, 0)}</span>
                <span>Returned rev: ${formatCurrencySafe(item.returned_revenue, cfg)}</span>
              </div>
            </li>
          `
        )
        .join("")}</ul>`;

  return { subtitle: "", body };
}

function renderProductList(products, metaBuilder) {
  if (!products.length) {
    return '<p class="text-sm text-zinc-400">No product data yet.</p>';
  }

  return `<ul class="space-y-2">${products
    .slice(0, ADVANCED_CELL_LIST_LIMIT)
    .map((item) => {
      const meta = metaBuilder(item);
      return `
        <li class="rounded-md border border-white/10 p-2">
          <div class="font-medium">#${escapeHtml(String(item.rank || ""))} ${escapeHtml(item.title || "Unknown")}</div>
          <div class="mt-1 flex flex-wrap gap-2 text-xs text-zinc-400">
            <span>${renderTrendAwareInlineText(meta[0])}</span>
            <span>${renderTrendAwareInlineText(meta[1])}</span>
            <span>${renderTrendAwareInlineText(meta[2])}</span>
          </div>
        </li>
      `;
    })
    .join("")}</ul>`;
}

function renderKeyValueList(entries) {
  return `<ul class="space-y-1">${entries
    .map(
      ([label, value]) => `
        <li class="flex items-start justify-between gap-3 rounded-md border border-white/10 px-2 py-1.5 text-xs">
          <span class="text-zinc-400">${escapeHtml(label)}</span>
          <span class="text-right font-medium text-zinc-200">${renderTrendAwareInlineText(value)}</span>
        </li>
      `
    )
    .join("")}</ul>`;
}

function renderTrendAwareInlineText(value) {
  const raw = String(value ?? "--");
  const pattern = /[+-](?:\d{1,3}(?:,\d{3})*|\d+)(?:\.\d+)?%/g;
  let hasMatch = false;
  let html = "";
  let cursor = 0;

  raw.replace(pattern, (match, offset) => {
    hasMatch = true;
    if (offset > cursor) {
      html += escapeHtml(raw.slice(cursor, offset));
    }
    const trendTokenClass = match.startsWith("+") ? "trend-token-up" : "trend-token-down";
    html += `<span class="${trendTokenClass}">${escapeHtml(match)}</span>`;
    cursor = offset + match.length;
    return match;
  });

  if (!hasMatch) {
    return escapeHtml(raw);
  }
  if (cursor < raw.length) {
    html += escapeHtml(raw.slice(cursor));
  }
  return html;
}

export default App;
