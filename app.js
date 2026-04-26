const K = {
  theme: "jsa.theme",
  notes: "jsa.notes",
  pf: "jsa.portfolio.v2",
  wl: "jsa.watchlist.v1",
  chartOl: "jsa.chartOverlays",
  /** Prefix for `sessionStorage` keys — last good `/api/history` per symbol when live fetch fails. */
  histSnap: "jsa.histSnap.v1",
  /** 3-digit PIN for clearing portfolio / watchlist (default `000` until changed). */
  dangerPin: "jsa.dangerPin",
  /** Hide yellow “LLM not configured” strip until next browser session. */
  llmBannerDismiss: "jsa.llmBannerDismiss",
  /** Last successful GET /api/t212/rows `fetched_at` (ISO) for display next to Sync. */
  pfT212LastFetch: "jsa.pf.t212LastFetch",
  /**
   * Family read-only: long read token in session (same tab). Lets nav links keep `#/portfolio?view=family&token=…` after
   * leaving the page, so a bare `#/portfolio` tap does not show an empty local ledger.
   */
  pfFamilyRead: "jsa.pf.familyReadTokenV1",
  /** Set to `"1"` when this tab has successfully entered family view at least once (pairs with `pfFamilyRead`). */
  pfFamilyMode: "jsa.pf.familyModeV1",
};
/** Last successful search query — restored when returning from an instrument (`#/search`). */
const SEARCH_LAST_Q = "jsa.search.lastQ";
/** Which broker ledger is active in the Portfolio UI (`sessionStorage`). */
const PF_ACTIVE_KEY = "jsa.pf.activeBroker";
/** `sessionStorage` — per-ledger table sort for holdings (name, value, ccy, …). */
const PF_HOLDING_SORT_KEY = "jsa.pf.holdingSortV1";
/** Preferred currency order when sorting the CCY column (rest sort after this list). */
const PF_CCY_SORT_ORDER = [
  "USD",
  "EUR",
  "GBP",
  "CHF",
  "JPY",
  "CNY",
  "INR",
  "AUD",
  "CAD",
  "SEK",
  "NOK",
  "DKK",
  "PLN",
  "CZK",
  "HUF",
  "ZAR",
  "AED",
  "SAR",
  "KRW",
  "BRL",
  "TWD",
  "HKD",
  "SGD",
  "MXN",
  "TRY",
  "ILS",
  "THB",
  "IDR",
  "PHP",
  "BHD",
  "JOD",
  "KWD",
  "QAR",
  "EGP",
  "MAD",
  "RON",
  "BAM",
  "BGN",
  "HRK",
  "ISK",
  "GEL",
  "PEN",
  "CLP",
  "COP",
  "VND",
  "BDT",
  "PKR",
  "LKR",
  "MUR",
  "NGN",
  "GHS",
  "KES",
  "LBP",
  "DZD",
  "TND",
  "BWP",
  "NAD",
  "RUB",
];
const PF_T212 = "t212";
/** Crypto held in Trading 212 only (synced via read-only API). */
const PF_CRYPTO = "t212_crypto";
const PF_ZERODHA = "zerodha";
const PF_ETORO = "etoro";
const PF_INSURANCE = "insurance";
const PF_FIXED_DEPOSIT = "fixed_deposit";
const PF_MF_COIN = "mf_coin";
const PF_MF_KUVERA = "mf_kuvera";
/** Display order for tabs, storage, and combined EUR rollup. */
const PF_BROKER_IDS = [PF_T212, PF_CRYPTO, PF_ZERODHA, PF_ETORO, PF_MF_COIN, PF_MF_KUVERA, PF_INSURANCE, PF_FIXED_DEPOSIT];
/** When set, Portfolio shows the server snapshot for the family read-only link — not this device’s `localStorage`. */
let _pfSharedBundle = null;

const PF_BROKER_LABEL = {
  [PF_T212]: "Trading 212",
  [PF_CRYPTO]: "Crypto (T212)",
  [PF_ZERODHA]: "Zerodha",
  [PF_ETORO]: "eToro",
  [PF_MF_COIN]: "Zerodha Coin MF",
  [PF_MF_KUVERA]: "Kuvera MF",
  [PF_INSURANCE]: "Insurances",
  [PF_FIXED_DEPOSIT]: "Fixed deposits",
};

/** Insurance sub-ledgers (one tab each under Insurances). */
const PF_INS_CO_IDS = ["sbi_life", "aditya_birla", "allianz_retirement", "vrk_retirement", "other"];
const PF_INS_CO_LABEL = {
  sbi_life: "SBI LIFE",
  aditya_birla: "Aditya Birla Capital Insurance",
  allianz_retirement: "Allianz Retirement",
  vrk_retirement: "VRK Retirement",
  other: "Other (5th slot)",
};
const PF_INS_CO_KEY = "jsa.pf.insCompany";
/** Last-selected MF sub-ledger when main “Mutual funds” tab is active (`mf_coin` / `mf_kuvera`). */
const PF_MF_SUB_KEY = "jsa.pf.mfSub";
/** UI grouping: MF ledgers map to main tab id `mf`. */
const PF_MAIN_TAB_MF = "mf";
/** Top-level tab id for the T212 crypto ledger. */
const PF_MAIN_TAB_CRYPTO = "crypto";

function brokerToMainTab(b) {
  if (b === PF_MF_COIN || b === PF_MF_KUVERA) return PF_MAIN_TAB_MF;
  if (b === PF_CRYPTO) return PF_MAIN_TAB_CRYPTO;
  return b;
}

function getMainPortfolioTab() {
  return brokerToMainTab(getActiveBroker());
}

function isPfMainTabId(s) {
  return (
    s === PF_T212 ||
    s === PF_MAIN_TAB_CRYPTO ||
    s === PF_ZERODHA ||
    s === PF_ETORO ||
    s === PF_MAIN_TAB_MF ||
    s === PF_INSURANCE ||
    s === PF_FIXED_DEPOSIT
  );
}

function getMfSubBroker() {
  try {
    const s = sessionStorage.getItem(PF_MF_SUB_KEY);
    if (s === PF_MF_COIN || s === PF_MF_KUVERA) return s;
  } catch {
    /* ignore */
  }
  return PF_MF_COIN;
}

function isPfInsCoId(s) {
  return typeof s === "string" && PF_INS_CO_IDS.includes(s);
}

function getPfInsuranceCompany() {
  try {
    const s = sessionStorage.getItem(PF_INS_CO_KEY);
    if (isPfInsCoId(s)) return s;
  } catch {
    /* ignore */
  }
  return PF_INS_CO_IDS[0];
}

function setPfInsuranceCompany(id) {
  if (!isPfInsCoId(id)) return;
  try {
    sessionStorage.setItem(PF_INS_CO_KEY, id);
  } catch {
    /* ignore */
  }
}

function pfRowId(r) {
  const id = String(r?.pfRowId || "").trim();
  if (id) return id;
  return "";
}

function ensurePfRowId(r) {
  if (!r || typeof r !== "object") return;
  if (!pfRowId(r)) {
    try {
      r.pfRowId = crypto.randomUUID();
    } catch {
      r.pfRowId = `r${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
    }
  }
}

/** @param {unknown} r */
function sumInsurancePayments(r) {
  const arr = Array.isArray(r?.payments) ? r.payments : [];
  let s = 0;
  for (const p of arr) s += num(p?.amount);
  return s;
}

/** Total invested (initial + logged premiums). */
function insuranceInvestedTotal(r) {
  return num(r?.valueAtPurchase) + sumInsurancePayments(r);
}

function isInsuranceAltRow(r) {
  return Boolean(r && (r.policyName != null || r.policyNo != null || r.valueAtPurchase != null));
}

function isFdAltRow(r) {
  return Boolean(r && (r.fdName != null || r.principal != null || r.fdBank != null));
}

function isStockLikePfBroker(b) {
  return b === PF_T212 || b === PF_CRYPTO || b === PF_ZERODHA || b === PF_ETORO || b === PF_MF_COIN || b === PF_MF_KUVERA;
}

function isPfBrokerId(s) {
  return typeof s === "string" && PF_BROKER_IDS.includes(s);
}

function pfBrokersEmpty(brokers) {
  return PF_BROKER_IDS.every((id) => !(brokers?.[id]?.rows?.length > 0));
}

/** Ensure every known broker key exists with a `rows` array (mutates). @returns {boolean} whether structure changed */
function ensurePfBrokerShape(brokers) {
  if (!brokers || typeof brokers !== "object") return false;
  let changed = false;
  for (const id of PF_BROKER_IDS) {
    if (!brokers[id]) {
      brokers[id] = { rows: [] };
      changed = true;
    } else if (!Array.isArray(brokers[id].rows)) {
      brokers[id].rows = [];
      changed = true;
    }
  }
  return changed;
}

/** Normalize insurance company keys, row ids, etc. @returns {boolean} whether data changed */
function migratePortfolioBundleShape(bundle) {
  if (!bundle || bundle.v !== 2 || !bundle.brokers) return false;
  let changed = false;
  const insRows = bundle.brokers[PF_INSURANCE]?.rows;
  if (Array.isArray(insRows)) {
    for (const r of insRows) {
      if (!r || typeof r !== "object") continue;
      const co = String(r.insCompany || "").trim();
      if (!isPfInsCoId(co)) {
        r.insCompany = "other";
        changed = true;
      }
      if (isInsuranceAltRow(r) && !pfRowId(r)) {
        ensurePfRowId(r);
        changed = true;
      }
    }
  }
  const fdRows = bundle.brokers[PF_FIXED_DEPOSIT]?.rows;
  if (Array.isArray(fdRows)) {
    for (const r of fdRows) {
      if (!r || typeof r !== "object") continue;
      if (isFdAltRow(r) && !pfRowId(r)) {
        ensurePfRowId(r);
        changed = true;
      }
    }
  }
  for (const brId of PF_BROKER_IDS) {
    const rws = bundle.brokers[brId]?.rows;
    if (!Array.isArray(rws)) continue;
    for (const r of rws) {
      if (!r || typeof r !== "object" || r.ccy == null) continue;
      const t = String(r.ccy).trim();
      if (t === "€" || t === "\u20ac" || t.toUpperCase() === "EURO") {
        r.ccy = "EUR";
        changed = true;
      }
    }
  }
  return changed;
}

function getDangerPin() {
  try {
    const p = localStorage.getItem(K.dangerPin);
    return p && /^\d{3}$/.test(p) ? p : "000";
  } catch {
    return "000";
  }
}

/** @returns {boolean} */
function requireDangerPin(actionLabel) {
  const pin = getDangerPin();
  const hint = pin === "000" ? " Default PIN is 000 — set your own under Portfolio → Safety PIN." : "";
  const entered = window.prompt(`Enter your 3-digit safety PIN to ${actionLabel}.${hint}`, "");
  if (entered === null) return false;
  if (String(entered).trim() !== pin) {
    window.alert("PIN does not match. Nothing was changed.");
    return false;
  }
  return true;
}

/** Must match `api_revision` in `GET /api/health` from this repo’s `server.py` (bump both when API surface changes). */
const MIN_API_REVISION = 13;

/** Short-lived cache so tabbing routes does not spam `/api/health`. */
let _healthBannerCache = /** @type {{ t: number; j: object | null; ok: boolean } | null} */ (null);

/** Last `/api/history` bars so chart type / theme changes can redraw without refetching. */
let _histChartBars = null;
let _histChartMeta = { source: "", range: "" };
/** True when bars come from `sessionStorage` snapshot (live API failed). */
let _histChartStale = false;
/** @type {"area" | "candle" | "heikin"} */
let _histChartType = "area";
/** Deduplicate concurrent `loadHistoryChart` calls (Technicals + expand chart). */
let _histLoadPromise = /** @type {Promise<void> | null} */ (null);

function histSnapStorageKey(sym, ex, rng) {
  return `${K.histSnap}:${String(sym || "").trim().toUpperCase()}|${String(ex || "").trim().toUpperCase()}|${String(rng || "1y").toLowerCase()}`;
}

function readHistSnap(sym, ex, rng) {
  try {
    const raw = sessionStorage.getItem(histSnapStorageKey(sym, ex, rng));
    if (!raw) return null;
    const o = JSON.parse(raw);
    if (!o || o.v !== 1 || !Array.isArray(o.bars) || o.bars.length < 2) return null;
    if (typeof o.source !== "string" || typeof o.range !== "string") return null;
    return o;
  } catch {
    return null;
  }
}

function writeHistSnap(sym, ex, rng, payload) {
  try {
    const out = {
      v: 1,
      bars: payload.bars,
      source: payload.source,
      range: payload.range,
      rsExtra: payload.rsExtra ?? null,
      t: Date.now(),
    };
    sessionStorage.setItem(histSnapStorageKey(sym, ex, rng), JSON.stringify(out));
  } catch {
    /* quota or private mode */
  }
}

function setChartStaleBanner(on, detail) {
  const el = $("chartStaleBanner");
  if (!el) return;
  if (!on) {
    el.hidden = true;
    el.textContent = "";
    return;
  }
  el.hidden = false;
  el.textContent =
    detail ||
    "Showing a cached snapshot from this browser session. Live history could not be loaded (often provider rate limits).";
}

/**
 * Restores chart + technicals from session snapshot when the history API fails.
 * @returns {boolean} true if a snapshot was applied
 */
function tryApplyHistSnap(sym, ex, rng, cv, mountEl) {
  const snap = readHistSnap(sym, ex, rng);
  if (!snap) return false;
  _histChartBars = snap.bars;
  _histChartMeta = { source: snap.source, range: snap.range };
  _histChartStale = true;
  drawHistoryChart(cv, snap.bars, _histChartMeta.range, _histChartType);
  syncHistChartTypeButtons();
  updateChartStatusLine();
  setChartStaleBanner(
    true,
    "Cached snapshot from this browser session — live history could not load (often provider rate limits). Retry later; server cache (API_HISTORY_CACHE_SECONDS) also helps after one good fetch.",
  );
  if (mountEl) mountEl.innerHTML = technicalsHtmlFromBars(snap.bars, snap.rsExtra || null);
  setInstrHistoryContext(snap.bars, snap.rsExtra || null);
  return true;
}

function syncHistChartTypeButtons() {
  const root = $("instrChartDock");
  if (!root) return;
  root.querySelectorAll("button[data-chart-type]").forEach((b) => {
    if (b instanceof HTMLButtonElement) b.setAttribute("aria-pressed", b.dataset.chartType === _histChartType ? "true" : "false");
  });
}

function updateChartStatusLine() {
  const st = $("chartSt");
  if (!st || !_histChartBars || _histChartBars.length < 2) return;
  const mode =
    _histChartType === "area" ? "close (area)" : _histChartType === "candle" ? "OHLC candles" : "Heikin Ashi";
  const stale = _histChartStale ? " · session cache" : "";
  st.textContent = `Chart: ${_histChartMeta.source} · ${_histChartMeta.range} · ${_histChartBars.length} trading days · ${mode}${stale}`;
}

function redrawHistCanvas() {
  const cv = $("histCanvas");
  if (!(cv instanceof HTMLCanvasElement) || !_histChartBars || _histChartBars.length < 2) return;
  drawHistoryChart(cv, _histChartBars, _histChartMeta.range, _histChartType);
  syncVolPanel();
}

function syncVolPanel() {
  const dock = $("volDock");
  const vc = $("histVolCanvas");
  if (!(vc instanceof HTMLCanvasElement)) return;
  const on = getChartOverlayState().vol && _histChartBars && _histChartBars.length >= 2;
  if (dock instanceof HTMLElement) dock.hidden = !on;
  if (on) drawVolumePanel(vc, _histChartBars);
  else {
    const ctx = vc.getContext("2d");
    if (ctx) {
      vc.width = vc.width;
    }
  }
}

function loadChartOverlayPrefs() {
  try {
    const x = JSON.parse(localStorage.getItem(K.chartOl) || "{}");
    return {
      ma20: x.ma20 !== false,
      ma50: x.ma50 !== false,
      bb: x.bb === true,
      vol: x.vol !== false,
      trend: x.trend !== false,
    };
  } catch {
    /* ignore */
  }
  return { ma20: true, ma50: true, bb: false, vol: true, trend: true };
}

function saveChartOverlayPrefs(o) {
  try {
    localStorage.setItem(K.chartOl, JSON.stringify(o));
  } catch {
    /* ignore */
  }
}

function getChartOverlayState() {
  const g = (id) => {
    const el = $(id);
    return el instanceof HTMLInputElement && el.type === "checkbox" ? el.checked : false;
  };
  if ($("olMa20")) {
    return {
      ma20: g("olMa20"),
      ma50: g("olMa50"),
      bb: g("olBb"),
      vol: g("olVol"),
      trend: g("olTrend"),
    };
  }
  return loadChartOverlayPrefs();
}

/** Benchmark for RS55-style relative strength (Bajaj: stock vs index, ~55 sessions). */
function benchSymbolForExchange(ex) {
  const u = String(ex || "").toUpperCase();
  if (u === "NSE" || u === "BSE") return "^NSEI";
  return "SPY";
}

function benchHumanLabel(benchSym) {
  if (benchSym === "^NSEI") return "Nifty 50";
  if (benchSym === "SPY") return "SPY (S&P 500 proxy)";
  return benchSym || "benchmark";
}

/** Aligned { ratio: stockClose/benchClose } oldest → newest, matching calendar `t`. */
function alignStockBenchByT(stockBars, benchBars) {
  const m = new Map();
  for (const b of benchBars) {
    const c = Number(b.c);
    if (Number.isFinite(c) && c !== 0) m.set(Number(b.t), c);
  }
  const out = [];
  for (const s of stockBars) {
    const t = Number(s.t);
    const cs = Number(s.c);
    const cb = m.get(t);
    if (!Number.isFinite(cs) || !Number.isFinite(cb) || cb === 0) continue;
    out.push({ t, ratio: cs / cb });
  }
  return out;
}

/** Last vs mean of last 55 ratios: >0 ⇒ current stock/bench ratio above its 55-day average (strength). */
function rs55BajajRatio(aligned) {
  if (aligned.length < 55) return null;
  const w = aligned.slice(-55);
  const last = w[w.length - 1].ratio;
  const mean = w.reduce((a, b) => a + b.ratio, 0) / 55;
  if (mean === 0 || !Number.isFinite(mean)) return null;
  return last / mean - 1;
}

function emptyPfBundle() {
  /** @type {Record<string, { rows: unknown[] }>} */
  const brokers = {};
  for (const id of PF_BROKER_IDS) brokers[id] = { rows: [] };
  return { v: 2, brokers };
}

function loadPfBundle() {
  if (_pfSharedBundle != null) {
    const x = JSON.parse(JSON.stringify(_pfSharedBundle));
    if (x?.v === 2 && x.brokers && typeof x.brokers === "object") {
      ensurePfBrokerShape(x.brokers);
      migratePortfolioBundleShape(x);
    }
    return x;
  }
  try {
    const raw = localStorage.getItem(K.pf);
    if (!raw) return emptyPfBundle();
    const x = JSON.parse(raw);
    // Accept any v2 object with a brokers map — missing broker keys are filled by
    // ensurePfBrokerShape. Requiring t212+zerodha up front caused total loss if
    // one key was absent (JSON looked "v2" but load fell through to emptyPfBundle).
    if (x?.v === 2 && x.brokers && typeof x.brokers === "object") {
      if (ensurePfBrokerShape(x.brokers)) savePfBundle(x);
      if (migratePortfolioBundleShape(x)) savePfBundle(x);
      return x;
    }
    if (Array.isArray(x?.rows)) {
      const b = emptyPfBundle();
      b.brokers[PF_T212].rows = x.rows;
      savePfBundle(b);
      return b;
    }
  } catch {
    /* ignore */
  }
  return emptyPfBundle();
}

function savePfBundle(bundle) {
  if (_pfSharedBundle != null) {
    return;
  }
  try {
    localStorage.setItem(K.pf, JSON.stringify(bundle));
  } catch {
    /* ignore */
  }
}

function clearFamilySessionPair() {
  try {
    sessionStorage.removeItem(K.pfFamilyRead);
    sessionStorage.removeItem(K.pfFamilyMode);
  } catch {
    /* ignore */
  }
}

function isFamilySessionPairActive() {
  try {
    if (sessionStorage.getItem(K.pfFamilyMode) !== "1") return false;
    return Boolean((sessionStorage.getItem(K.pfFamilyRead) || "").trim());
  } catch {
    return false;
  }
}

/** True when this browser has no v2 portfolio rows on disk (typical for a “family” phone; owners usually have at least one row). */
function isLocalPfStorageEmpty() {
  try {
    const raw = localStorage.getItem(K.pf);
    if (!raw) return true;
    const x = JSON.parse(raw);
    if (x?.v === 2 && x.brokers && typeof x.brokers === "object") {
      for (const id of PF_BROKER_IDS) {
        if (Array.isArray(x.brokers[id]?.rows) && x.brokers[id].rows.length > 0) return false;
      }
    } else if (Array.isArray(x?.rows) && x.rows.length > 0) {
      return false;
    }
  } catch {
    return true;
  }
  return true;
}

/**
 * Strips a query key from the location hash and replaces history, without adding a new session entry.
 * @param {string} key
 */
function stripHashQueryKey(key) {
  const k = String(key || "").trim();
  if (!k) return;
  const raw0 = (location.hash || "#").replace(/^#/, "");
  const qi = raw0.indexOf("?");
  if (qi < 0) return;
  const pathPart = raw0.slice(0, qi) || "/search";
  const q = new URLSearchParams(raw0.slice(qi + 1));
  if (!q.has(k)) return;
  q.delete(k);
  const nq = q.toString();
  const next = nq ? `#${pathPart}?${nq}` : `#${pathPart}`;
  if (location.hash === next) return;
  try {
    history.replaceState(null, "", `${location.pathname}${location.search}${next}`);
  } catch {
    /* ignore */
  }
}

function familyReadOnlyPortfolioHref(/** @type {string} */ t) {
  return `#/portfolio?view=family&token=${encodeURIComponent(t)}`;
}

/**
 * @returns {string} `#/portfolio` or the family read-only link when this tab is using a family snapshot
 *   (in-memory, or a resume on a device with no local ledger rows so Search → Portfolio keeps the token).
 */
function hrefPortfolio() {
  try {
    const { sp } = parseLocationHash();
    if ((sp.get("view") || "").trim().toLowerCase() === "family") {
      const u = familyReadTokenFromUrl(sp);
      if (u) return familyReadOnlyPortfolioHref(u);
    }
  } catch {
    /* ignore */
  }
  if (!isFamilySessionPairActive()) return "#/portfolio";
  const t = (() => {
    try {
      return (sessionStorage.getItem(K.pfFamilyRead) || "").trim();
    } catch {
      return "";
    }
  })();
  if (!t) return "#/portfolio";
  if (_pfSharedBundle != null) return familyReadOnlyPortfolioHref(t);
  if (isLocalPfStorageEmpty()) return familyReadOnlyPortfolioHref(t);
  return "#/portfolio";
}

function updateFamilyNavHrefs() {
  const a = document.querySelector("a[data-pf-nav]");
  if (!(a instanceof HTMLAnchorElement)) return;
  a.href = hrefPortfolio();
}

/** @returns {Promise<{ ok: boolean, updated_at?: string, detail?: string, err?: string, status?: number }>} */
async function fetchSharedFamilyPortfolio(readToken) {
  _pfSharedBundle = null;
  const t = String(readToken || "").trim();
  if (!t) return { ok: false };
  try {
    const r = await fetch(`/api/shared/portfolio?token=${encodeURIComponent(t)}`, { cache: "no-store" });
    const j = await r.json().catch(() => ({}));
    if (!r.ok || !j.ok || !j.bundle || typeof j.bundle !== "object") {
      return {
        ok: false,
        detail: typeof j.detail === "string" ? j.detail : "",
        err: String(j.error || ""),
        status: r.status,
      };
    }
    const b = j.bundle;
    if (b.v !== 2 || !b.brokers) return { ok: false, err: "invalid bundle shape" };
    _pfSharedBundle = JSON.parse(JSON.stringify(b));
    return { ok: true, updated_at: typeof j.updated_at === "string" ? j.updated_at : undefined };
  } catch {
    _pfSharedBundle = null;
    return { ok: false, err: "network" };
  }
}

/**
 * Family read-only: load server snapshot for `?view=family&token=…` in the hash, or recover from
 * a prior successful load in this tab (so Search → Portfolio does not show an empty ledger on phones).
 */
async function applyPortfolioSharedFromHash(sp) {
  const ban = $("pfSharedBanner");
  const manage = $("pfManage");
  const view = (sp.get("view") || "").trim().toLowerCase();
  const exitFam = (sp.get("exitFamily") || "").trim() === "1";
  if (exitFam) {
    clearFamilySessionPair();
    _pfSharedBundle = null;
    stripHashQueryKey("exitFamily");
    if (ban instanceof HTMLElement) {
      ban.hidden = true;
      ban.textContent = "";
      ban.className = "card2 mt";
    }
    if (manage instanceof HTMLElement) manage.hidden = false;
    const fte = $("pfFamilyTools");
    if (fte instanceof HTMLElement) fte.hidden = true;
    updateFamilyNavHrefs();
    return;
  }
  const tokInUrl = familyReadTokenFromUrl(sp);
  let tok = tokInUrl;
  if (!tok) {
    if (isFamilySessionPairActive() && isLocalPfStorageEmpty()) {
      try {
        tok = (sessionStorage.getItem(K.pfFamilyRead) || "").trim();
      } catch {
        tok = "";
      }
    }
  }
  const canResumeFromSession = !tokInUrl && isFamilySessionPairActive() && isLocalPfStorageEmpty();
  const wantsFamily = Boolean(tok) && (view === "family" || canResumeFromSession);
  if (wantsFamily) {
    const { ok, updated_at: ua, detail, err, status: httpSt } = await fetchSharedFamilyPortfolio(tok);
    if (ok) {
      try {
        sessionStorage.setItem(K.pfFamilyRead, tok);
        sessionStorage.setItem(K.pfFamilyMode, "1");
      } catch {
        /* ignore */
      }
      if (ban instanceof HTMLElement) {
        ban.hidden = false;
        const when = ua ? ` · snapshot ${esc(String(ua))}` : "";
        ban.className = "card2 mt migrateBanner";
        ban.innerHTML = `<p class="sml"><strong>Family view (read-only)</strong> — data from the server${when}. Scrolling: totals load after FX; use the buttons for latest prices. <a class="backLink" href="#/portfolio?exitFamily=1">Open normal portfolio</a> (this device only) — clears the family link in this tab.</p>`;
      }
      if (manage instanceof HTMLElement) manage.hidden = true;
      const ft = $("pfFamilyTools");
      if (ft instanceof HTMLElement) ft.hidden = false;
      updateFamilyNavHrefs();
      return;
    }
    if (ban instanceof HTMLElement) {
      ban.hidden = false;
      ban.className = "card2 mt globalApiBannerErr";
      const hint =
        detail || err
          ? `<br/><span class="sml">${esc(detail || err || "")}</span>${
              httpSt === 404
                ? ` <span class="sml">On <strong>Render free</strong>, a snapshot saved only to disk can disappear when the app sleeps. Ask the owner to set <strong>GitHub Gist</strong> storage and publish again (see <code>/.env.example</code>).</span>`
                : ""
            }`
          : "";
      ban.innerHTML = `<p class="sml"><strong>Could not load family snapshot.</strong> Check the read token, or the owner may need to publish again. <a href="/api/health" target="_blank" rel="noopener">/api/health</a> → <code>shared_family_portfolio</code>.${hint}</p>`;
    }
    if (manage instanceof HTMLElement) manage.hidden = false;
    const ftb = $("pfFamilyTools");
    if (ftb instanceof HTMLElement) ftb.hidden = true;
    if (!tokInUrl) clearFamilySessionPair();
    updateFamilyNavHrefs();
    return;
  }
  _pfSharedBundle = null;
  if (ban instanceof HTMLElement) {
    ban.hidden = true;
    ban.textContent = "";
    ban.className = "card2 mt";
  }
  if (manage instanceof HTMLElement) manage.hidden = false;
  const ftx = $("pfFamilyTools");
  if (ftx instanceof HTMLElement) ftx.hidden = true;
  updateFamilyNavHrefs();
}

async function publishPortfolioToServer() {
  if (_pfSharedBundle != null) {
    status("Open the normal portfolio (not the family link) to publish your data.");
    return;
  }
  const k = window.prompt(
    "Paste the publish key (must match SHARED_PORTFOLIO_WRITE_TOKEN in Render). It is not stored in this app.",
  );
  if (!k || !String(k).trim()) {
    status("Publish cancelled");
    return;
  }
  const bundle = loadPfBundle();
  try {
    const r = await fetch("/api/shared/portfolio", {
      method: "PUT",
      headers: { "Content-Type": "application/json", "X-Portfolio-Write-Key": String(k).trim() },
      body: JSON.stringify(bundle),
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) {
      const err = String(j.error || "");
      if (r.status === 403 && (err === "forbidden" || !j.detail)) {
        status(
          "Publish failed: write key rejected. Paste SHARED_PORTFOLIO_WRITE_TOKEN from Render exactly (this is not the read token; no extra spaces).",
        );
        return;
      }
      if (r.status === 503 && (err === "publish_not_configured" || j.detail)) {
        status(
          String(
            j.detail || "Set SHARED_PORTFOLIO_WRITE_TOKEN in Render, redeploy, then try again.",
          ),
        );
        return;
      }
      status(String(j.detail || j.error || `HTTP ${r.status}`));
      return;
    }
    status(`Published for family view · ${j.updated_at || "ok"}`);
  } catch (e) {
    status(e instanceof Error ? e.message : String(e));
  }
}

/** Read-only family link: re-fetch /api/quote for every row (no owner needed). */
async function familyRefreshAllMarketPrices() {
  if (!_pfSharedBundle) {
    status("Not in family view");
    return;
  }
  const bundle = JSON.parse(JSON.stringify(_pfSharedBundle));
  let n = 0;
  for (const id of PF_BROKER_IDS) {
    const rows = bundle.brokers[id]?.rows;
    if (!Array.isArray(rows) || !rows.length) continue;
    status(`Prices: ${PF_BROKER_LABEL[id] || id}…`);
    n += await applyLiveQuotesToRowsForBroker(bundle, id);
  }
  _pfSharedBundle = JSON.parse(JSON.stringify(bundle));
  renderPf();
  status(n > 0 ? `Market prices updated (${n} quote run(s))` : "Market prices — no symbols updated (check API keys if empty)");
}

/** Re-download the last snapshot the owner published (same token as in the URL). */
async function familyReloadOwnerSnapshot() {
  if (!_pfSharedBundle) {
    status("Not in family view");
    return;
  }
  const { sp } = parseLocationHash();
  let tok = familyReadTokenFromUrl(sp);
  if (!tok) {
    try {
      tok = (sessionStorage.getItem(K.pfFamilyRead) || "").trim();
    } catch {
      tok = "";
    }
  }
  if (!tok) {
    status("Missing token in link");
    return;
  }
  status("Reloading owner snapshot…");
  const { ok, updated_at: ua } = await fetchSharedFamilyPortfolio(tok);
  if (!ok) {
    status("Could not reload — check link or ask the owner to publish again");
    return;
  }
  renderPf();
  const when = ua ? ` · ${ua}` : "";
  status(`Owner snapshot reloaded${when}`);
}

function getActiveBroker() {
  try {
    const s = sessionStorage.getItem(PF_ACTIVE_KEY);
    if (isPfBrokerId(s)) return s;
  } catch {
    /* ignore */
  }
  return PF_T212;
}

function syncPortfolioTabAria() {
  const main = getMainPortfolioTab();
  const act = getActiveBroker();
  document.querySelectorAll("[data-main-tab]").forEach((btn) => {
    if (!(btn instanceof HTMLButtonElement)) return;
    btn.setAttribute("aria-selected", btn.dataset.mainTab === main ? "true" : "false");
  });
  document.querySelectorAll("[data-mf-sub]").forEach((btn) => {
    if (!(btn instanceof HTMLButtonElement)) return;
    const sel = main === PF_MAIN_TAB_MF && btn.dataset.mfSub === act;
    btn.setAttribute("aria-selected", sel ? "true" : "false");
  });
}

function setActiveBroker(b) {
  if (!isPfBrokerId(b)) return;
  try {
    sessionStorage.setItem(PF_ACTIVE_KEY, b);
    if (b === PF_MF_COIN || b === PF_MF_KUVERA) sessionStorage.setItem(PF_MF_SUB_KEY, b);
  } catch {
    /* ignore */
  }
  syncPortfolioTabAria();
  updatePfBrokerCaption();
}

function insuranceRowsForCompany(bundle, co) {
  const rows = bundle?.brokers?.[PF_INSURANCE]?.rows;
  if (!Array.isArray(rows)) return [];
  return rows.filter((r) => String(r?.insCompany || "other") === co);
}

function updatePfBrokerCaption() {
  const cap = $("pfBrokerCap");
  if (!cap) return;
  const b = getActiveBroker();
  const bundle = loadPfBundle();
  let n = 0;
  if (b === PF_INSURANCE) {
    n = insuranceRowsForCompany(bundle, getPfInsuranceCompany()).length;
    const lab = PF_INS_CO_LABEL[getPfInsuranceCompany()] || "";
    cap.textContent = `${PF_BROKER_LABEL[b]} · ${lab} · ${n} row(s)`;
    return;
  }
  if (b === PF_MF_COIN || b === PF_MF_KUVERA) {
    const rows = bundle.brokers[b]?.rows;
    n = Array.isArray(rows) ? rows.length : 0;
    cap.textContent = `Mutual funds · ${PF_BROKER_LABEL[b]} · ${n} row(s)`;
    return;
  }
  const rows = bundle.brokers[b]?.rows;
  n = Array.isArray(rows) ? rows.length : 0;
  cap.textContent = `${PF_BROKER_LABEL[b] ?? b} · ${n} row(s)`;
}

function emptyWl() {
  return { v: 1, rows: [] };
}

function loadWl() {
  try {
    const raw = localStorage.getItem(K.wl);
    if (!raw) return emptyWl();
    const x = JSON.parse(raw);
    if (x?.v === 1 && Array.isArray(x.rows)) {
      return x;
    }
  } catch {
    /* ignore */
  }
  return emptyWl();
}

function saveWl(w) {
  try {
    localStorage.setItem(K.wl, JSON.stringify(w));
  } catch {
    /* ignore */
  }
}

function wlKey(r) {
  return `${String(r.sym || "").trim().toUpperCase()}|${String(r.ex || "").trim().toUpperCase()}`;
}

function addWlRow(row) {
  const w = loadWl();
  const k = wlKey(row);
  if (w.rows.some((r) => wlKey(r) === k)) return false;
  w.rows.push({
    sym: String(row.sym || "").trim(),
    ex: String(row.ex || "").trim(),
    nm: String(row.nm || "").trim(),
    ccy: String(row.ccy || "").trim().toUpperCase(),
  });
  saveWl(w);
  return true;
}

function removeWlRow(sym, ex) {
  const w = loadWl();
  const k = `${String(sym || "").trim().toUpperCase()}|${String(ex || "").trim().toUpperCase()}`;
  w.rows = w.rows.filter((r) => wlKey(r) !== k);
  saveWl(w);
}

/** Old PWA portfolio key (pre–v2 rebuild) — auto-migrated once into `K.pf`. */
const LEGACY_PORTFOLIOS = "johnsstockapp.portfolios.v1";
const LEGACY_NOTICE = "jsa.legacyMigrateNotice";

let _legacyMigrated = false;

/** Pull holdings from legacy localStorage into v2 `{ rows }` if v2 is empty. */
function migrateLegacyPortfolioFromV1() {
  if (_legacyMigrated) return 0;
  _legacyMigrated = true;
  try {
    const cur = localStorage.getItem(K.pf);
    if (cur) {
      try {
        const p = JSON.parse(cur);
        if (p?.v === 2 && p?.brokers) {
          const n = PF_BROKER_IDS.reduce((s, id) => s + (p.brokers[id]?.rows?.length || 0), 0);
          if (n > 0) return 0;
        } else if (Array.isArray(p?.rows) && p.rows.length > 0) return 0;
      } catch {
        /* ignore */
      }
    }
    const raw = localStorage.getItem(LEGACY_PORTFOLIOS);
    if (!raw) return 0;
    const old = JSON.parse(raw);
    const items = Array.isArray(old?.items) ? old.items : [];
    const rows = [];
    for (const it of items) {
      for (const h of Array.isArray(it?.holdings) ? it.holdings : []) {
        const sym = String(h.symbol || "").trim();
        const ccy = String(h.currency || "").trim().toUpperCase();
        const qty = num(h.qty);
        if (!sym || !ccy || qty <= 0) continue;
        rows.push({
          sym,
          ex: String(h.exchange || "").trim(),
          ccy,
          qty,
          avg: num(h.avgPrice),
          last: num(h.lastPrice),
          nm: String(h.name || "").trim(),
        });
      }
    }
    if (rows.length) {
      const bundle = loadPfBundle();
      bundle.brokers[PF_T212].rows = rows;
      savePfBundle(bundle);
      try {
        sessionStorage.setItem(LEGACY_NOTICE, String(rows.length));
      } catch {
        /* ignore */
      }
      return rows.length;
    }
  } catch {
    /* ignore */
  }
  return 0;
}

const $ = (id) => document.getElementById(id);

function isAppDevOnLocalhost() {
  const h = (location.hostname || "").toLowerCase();
  return h === "localhost" || h === "127.0.0.1" || h === "[::1]";
}

function isFileProtocol() {
  return location.protocol === "file:";
}

/** True when the app is opened on the public internet (Render, custom domain) — not file://, not localhost. */
function isHostedWebApp() {
  if (isFileProtocol()) return false;
  return !isAppDevOnLocalhost();
}

/** Cancels in-flight search when the user types again (debounce does not cancel fetch alone). */
let _searchAbort = null;

function esc(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function debounce(fn, ms) {
  let t;
  return (...a) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...a), ms);
  };
}

function status(a, b = "") {
  const L = $("stL");
  const R = $("stR");
  if (L) L.textContent = a;
  if (R) R.textContent = b;
}

/** Parse `#/path` and optional `?query` (e.g. `#/symbol?sym=AAPL&ex=NASDAQ`). */
function parseLocationHash() {
  const raw = (location.hash || "#/search").replace(/^#/, "");
  const qi = raw.indexOf("?");
  let pathPart = (qi >= 0 ? raw.slice(0, qi) : raw).trim();
  if (!pathPart || pathPart === "/") pathPart = "/search";
  const pathname = pathPart.toLowerCase();
  const q = qi >= 0 ? raw.slice(qi + 1) : "";
  const sp = new URLSearchParams(q);
  return { pathname, sp };
}

/** Strip accidental quotes from CSV / spreadsheet cells. */
function cleanHashParam(s) {
  let x = String(s ?? "").trim();
  if ((x.startsWith('"') && x.endsWith('"')) || (x.startsWith("'") && x.endsWith("'"))) {
    x = x.slice(1, -1).trim();
  }
  return x;
}

/**
 * `token=…` in the location hash. Decode safely (e.g. `%40` for @) and strip quotes from spreadsheets.
 * @param {URLSearchParams} sp
 * @returns {string}
 */
function familyReadTokenFromUrl(sp) {
  const raw = sp.get("token");
  if (raw == null) return "";
  let x = String(raw).trim();
  try {
    x = decodeURIComponent(x);
  } catch {
    /* keep */
  }
  return cleanHashParam(x).trim();
}

/** Hash link to the instrument workspace (do not wrap in `esc()` — that breaks `&` / `"` in the URL). */
/** @param {{ fromPf?: boolean }} [opts] — `fromPf` adds `pf=1` so the instrument page can show saved holding rows. */
function instrumentHref(sym, ex, nm, ccy, opts) {
  const p = new URLSearchParams();
  p.set("sym", cleanHashParam(sym));
  const exC = cleanHashParam(ex);
  const nmC = cleanHashParam(nm);
  const ccyC = cleanHashParam(ccy);
  if (exC) p.set("ex", exC);
  if (nmC) p.set("nm", nmC);
  if (ccyC) p.set("ccy", ccyC);
  if (opts?.fromPf) p.set("pf", "1");
  return `#/symbol?${p.toString()}`;
}

/** All saved rows matching symbol (and exchange when both sides have a code). */
function findPfHoldingMatches(sym, ex) {
  const su = String(sym || "").trim().toUpperCase();
  const eu = String(ex || "").trim().toUpperCase();
  if (!su) return [];
  const bundle = loadPfBundle();
  /** @type {{ brokerId: string, r: Record<string, unknown> }[]} */
  const out = [];
  for (const id of PF_BROKER_IDS) {
    if (id === PF_INSURANCE || id === PF_FIXED_DEPOSIT) continue;
    const rows = bundle.brokers[id]?.rows || [];
    for (const r of rows) {
      if (String(r.sym || "").trim().toUpperCase() !== su) continue;
      const re = String(r.ex || "").trim().toUpperCase();
      if (eu && re && eu !== re) continue;
      out.push({ brokerId: id, r });
    }
  }
  return out;
}

/** When opened from portfolio (`pf=1`), show saved qty / avg / last / P+L above the quote card. */
function paintInstrPfHold(sym, ex) {
  const el = $("instrPfHold");
  if (!(el instanceof HTMLElement)) return;
  const hits = findPfHoldingMatches(sym, ex);
  if (!hits.length) {
    el.hidden = true;
    el.innerHTML = "";
    return;
  }
  el.hidden = false;
  const blocks = hits
    .map(({ brokerId, r }, i) => {
      const qty = num(r.qty);
      const avg = num(r.avg);
      const last = num(r.last);
      const val = qty * last;
      const cost = qty * avg;
      const pl = val - cost;
      const ccy = String(r.ccy || "").toUpperCase() || "—";
      const plCls = pl >= 0 ? "plp" : "pln";
      const nm = String(r.nm || "").trim();
      const exR = String(r.ex || "").trim();
      const plPct = cost !== 0 && Number.isFinite(cost) ? (pl / cost) * 100 : null;
      const sub =
        nm || exR
          ? `<p class="sml muted">${nm ? esc(nm) : ""}${nm && exR ? " · " : ""}${exR ? esc(exR) : ""}</p>`
          : "";
      const pctTxt =
        Number.isFinite(plPct) && plPct !== 0
          ? ` <span class="muted">(${pl >= 0 ? "+" : ""}${esc(fmtN(plPct, 1))}% on cost)</span>`
          : "";
      const topPad = i > 0 ? " instrPfHoldBlockGap" : "";
      return `<div class="instrPfHoldBlock${topPad}">
      <div class="h3">Your holding · ${esc(PF_BROKER_LABEL[brokerId])}</div>
      ${sub}
      <div class="grid3 sml mt">
        <div>Qty <strong>${esc(fmtN(qty, 4))}</strong></div>
        <div>Avg <strong>${esc(fmtMoney(ccy, avg))}</strong></div>
        <div>Last <strong>${esc(fmtMoney(ccy, last))}</strong></div>
      </div>
      <div class="grid2 sml mt">
        <div>Value <strong>${esc(fmtMoney(ccy, val))}</strong></div>
        <div class="${plCls}">Unrealized P/L <strong>${esc(fmtMoney(ccy, pl))}</strong>${pctTxt}</div>
      </div>
      <div class="instrPfEurStrip" data-instr-pf-eur="${i}" aria-live="polite"><p class="sml muted">Loading EUR estimate (ECB reference)…</p></div>
    </div>`;
    })
    .join("");
  el.innerHTML = `<div class="instrPfHoldInner">${blocks}<p class="sml muted mt">Live quote below; expand <strong>Price history</strong> for the chart. Update <strong>Last</strong> from <a class="backLink" href="${hrefPortfolio()}">Portfolio</a> → <strong>Refresh prices</strong> on the ledger tab that holds this row.</p></div>`;
  void enrichInstrPfHoldEur(sym, ex);
}

/** Same FX as Portfolio combined EUR — fills <code>[data-instr-pf-eur]</code> strips. */
async function enrichInstrPfHoldEur(sym, ex) {
  const root = $("instrPfHold");
  if (!(root instanceof HTMLElement) || root.hidden) return;
  const hits = findPfHoldingMatches(sym, ex);
  if (!hits.length) return;
  const markErr = (msg) => {
    hits.forEach((_, i) => {
      const strip = root.querySelector(`[data-instr-pf-eur="${i}"]`);
      if (strip) strip.innerHTML = `<p class="sml">${msg}</p>`;
    });
  };
  let j;
  try {
    j = await fetchEurFxTable();
  } catch {
    markErr(
      `<span class="err">Could not load EUR reference.</span> <span class="muted">Try <a href="/api/fx-eur" target="_blank" rel="noopener"><code>/api/fx-eur</code></a> with the server running.</span>`,
    );
    return;
  }
  const eurPer = j.eur_per_unit;
  const src = esc(j.source || "ECB reference");
  const dt = esc(j.date || "");
  const discFoot = j.disclaimer ? `<p class="sml muted instrPfEurDisc">${esc(j.disclaimer)}</p>` : "";
  hits.forEach((hit, i) => {
    const strip = root.querySelector(`[data-instr-pf-eur="${i}"]`);
    if (!strip) return;
    const { totE, costE, plE, miss } = portfolioEurFromRows([hit.r], eurPer);
    if (miss.size) {
      strip.innerHTML = `<p class="sml"><strong>EUR</strong> <span class="muted">(${src}, ${dt})</span> — <span class="err">No ECB rate for ${esc([...miss].join(", "))}</span>. Set row <strong>currency</strong> (e.g. INR, USD) on Portfolio.</p>${i === 0 ? discFoot : ""}`;
      return;
    }
    if (!Number.isFinite(totE) && !Number.isFinite(costE)) {
      strip.innerHTML = `<p class="sml muted">EUR (${src}, ${dt}): no convertible amounts.</p>${i === 0 ? discFoot : ""}`;
      return;
    }
    const plPctE = costE !== 0 && Number.isFinite(costE) ? (plE / costE) * 100 : null;
    const plCls = plE >= 0 ? "plp" : "pln";
    const pctStr =
      Number.isFinite(plPctE) && plPctE !== 0
        ? ` <span class="muted">(${plE >= 0 ? "+" : ""}${esc(fmtN(plPctE, 1))}% on cost in EUR)</span>`
        : "";
    strip.innerHTML = `<p class="sml"><strong>EUR</strong> <span class="muted">(${src}, ${dt}; mid reference, not broker cash)</span> · Value <strong>${esc(fmtMoney("EUR", totE))}</strong> · Cost <strong>${esc(fmtMoney("EUR", costE))}</strong> · P/L <strong class="${plCls}">${esc(fmtMoney("EUR", plE))}</strong>${pctStr}</p>${i === 0 ? discFoot : ""}`;
  });
}

/** Last loaded quote on symbol page — used for “Add to portfolio”. */
let _instrQuoteCtx = {
  sym: "",
  ex: /** @type {string|null} */ (null),
  nm: "",
  ccyHint: "",
  /** @type {Record<string, unknown>|null} */
  q: null,
};

/** Latest daily bars + RS55 context on the instrument page (for POST /api/ai-commentary). */
let _instrHistoryContext = /** @type {{ bars: unknown[], rsExtra: object | null } | null} */ (null);

function setInstrHistoryContext(bars, rsExtra) {
  if (Array.isArray(bars) && bars.length >= 2) {
    _instrHistoryContext = { bars, rsExtra: rsExtra && typeof rsExtra === "object" ? rsExtra : null };
  } else {
    _instrHistoryContext = null;
  }
}

function fmtN(n, d = 2) {
  if (!Number.isFinite(n)) return "—";
  return new Intl.NumberFormat(undefined, {
    minimumFractionDigits: d,
    maximumFractionDigits: d,
  }).format(n);
}

/**
 * Map non-ISO codes from imports/APIs to ISO (ECB/Frankfurter `eur_per_unit` keys are ISO only).
 * Example: T212 / manual rows sometimes use "EURO" or "€" — treat as "EUR" for rates and display.
 */
function normalizeCcyForFx(ccy) {
  const t = String(ccy ?? "")
    .trim();
  if (t === "€" || t === "\u20ac") return "EUR";
  const u = t.toUpperCase();
  if (u === "EURO") return "EUR";
  return u;
}

/**
 * `eur_per_unit[CCY]` = EUR value of 1 unit of `CCY` (Frankfurter, ECB; server usually injects `EUR: 1`).
 * @returns a positive rate, or `undefined` if the table has no valid rate.
 */
function eurPerUnitToEur(eurPer, ccyRaw) {
  const c = normalizeCcyForFx(String(ccyRaw || ""));
  if (!c) return undefined;
  if (c === "EUR") {
    const e = eurPer?.EUR;
    if (typeof e === "number" && Number.isFinite(e) && e > 0) return e;
    return 1;
  }
  const v = eurPer?.[c];
  if (typeof v === "number" && Number.isFinite(v) && v > 0) return v;
  return undefined;
}

/** Convert a quote price in `qCcy` to an amount in `rowCcy` using the same ECB/Frankfurter `eur_per_unit` table. */
function quotePriceToRowCcy(px, qCcyRaw, rowCcyRaw, eurPer) {
  const a = normalizeCcyForFx(qCcyRaw);
  const b = normalizeCcyForFx(rowCcyRaw);
  if (!a || !b || a === b) return px;
  const e1 = eurPer ? eurPerUnitToEur(eurPer, a) : undefined;
  const e2 = eurPer ? eurPerUnitToEur(eurPer, b) : undefined;
  if (e1 == null || e2 == null || e2 === 0) return px;
  return (px * e1) / e2;
}

/** Ccy column label: euro zone → symbol, else ISO. */
function formatCcyLabel(ccy) {
  const c = normalizeCcyForFx(ccy);
  if (c === "EUR") return "€";
  return c || "—";
}

/** Format a number with a simple currency prefix/suffix (no FX conversion). */
function fmtMoney(ccy, n) {
  const c = normalizeCcyForFx(ccy);
  const sym = { USD: "$", EUR: "€", GBP: "£", INR: "₹", CHF: "CHF " }[c];
  const pfx = sym ?? (c ? `${c} ` : "");
  return pfx + fmtN(n);
}

/* theme — stored preference may be Auto; DOM uses resolved light/dark for CSS */
let _themeUserPref = "auto";

function getResolvedFromPref(pref) {
  if (pref === "light" || pref === "dark") return pref;
  return window.matchMedia?.("(prefers-color-scheme: light)").matches ? "light" : "dark";
}

function applyThemePref(pref) {
  if (pref !== "light" && pref !== "dark" && pref !== "auto") pref = "auto";
  _themeUserPref = pref;
  document.documentElement.dataset.theme = getResolvedFromPref(pref);
  try {
    localStorage.setItem(K.theme, pref);
  } catch {
    /* ignore */
  }
  const tb = $("btnTheme");
  if (tb instanceof HTMLButtonElement) {
    tb.textContent =
      pref === "auto" ? "Theme · Auto" : pref === "light" ? "Theme · Light" : "Theme · Dark";
  }
  redrawHistCanvas();
}

function loadThemePref() {
  try {
    const x = localStorage.getItem(K.theme);
    if (x === "light" || x === "dark" || x === "auto") return x;
  } catch {
    /* ignore */
  }
  return "auto";
}

/* install */
function setupInstall() {
  const b = $("btnInstall");
  if (!(b instanceof HTMLButtonElement)) return;
  let ev = null;
  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    ev = e;
    b.hidden = false;
  });
  b.addEventListener("click", async () => {
    if (!ev) return;
    b.disabled = true;
    try {
      ev.prompt();
      await ev.userChoice;
    } finally {
      ev = null;
      b.hidden = true;
      b.disabled = false;
    }
  });
}

/* SW — on Render (or any `*.onrender.com`) we skip registration and clear old caches. A poisoned
   precache (from earlier broken deploys) made `app.js` “half-load”: header OK, but `#/search`
   `wire()` never bound — nav hash routing + search input did nothing. Localhost still uses SW. */
async function setupSw() {
  if (!("serviceWorker" in navigator)) return;
  const isRenderHost = /\.onrender\.com$/i.test(location.hostname);
  if (isRenderHost) {
    try {
      const regs = await navigator.serviceWorker.getRegistrations();
      await Promise.all(regs.map((r) => r.unregister()));
    } catch {
      /* ignore */
    }
    if ("caches" in window) {
      try {
        const keys = await caches.keys();
        await Promise.all(keys.map((k) => caches.delete(k)));
      } catch {
        /* ignore */
      }
    }
    return;
  }
  const b = document.documentElement.getAttribute("data-build") || "1";
  try {
    const reg = await navigator.serviceWorker.register(`/service-worker.js?v=${encodeURIComponent(b)}`, {
      scope: "/",
      /** @type {RegistrationOptions} */ updateViaCache: "none",
    });
    void reg.update();
  } catch {
    /* ignore */
  }
}

/* routing */
function formatApiErrPayload(j, r) {
  const err = String(j?.error || j?.detail || r.status || "Request failed");
  let hint = j?.hint != null ? String(j.hint) : "";
  if (hint.length > 280) hint = `${hint.slice(0, 277)}…`;
  return `<p class="err">${esc(err)}</p>${hint ? `<p class="muted sml">${esc(hint)}</p>` : ""}<p class="sml muted"><a href="/api/health" target="_blank" rel="noopener">Open <code>/api/health</code></a> · <a href="/README.html#restart-server">Restart server</a></p>`;
}

async function refreshGlobalApiBanner() {
  const el = $("globalApiBanner");
  if (!el) return;
  const now = Date.now();
  if (_healthBannerCache && now - _healthBannerCache.t < 8000) {
    paintGlobalApiBanner(el, _healthBannerCache.j, _healthBannerCache.ok);
    return;
  }
  const maxAttempts = isHostedWebApp() ? 3 : 1;
  const delays = [0, 400, 1000];
  let j = null;
  let ok = false;
  for (let attempt = 0; attempt < maxAttempts; attempt++) {
    if (delays[attempt] > 0) {
      await new Promise((r) => setTimeout(r, delays[attempt]));
    }
    try {
      const r = await fetch("/api/health", { cache: "no-store" });
      const text = await r.text();
      let parsed = null;
      if (text) {
        try {
          parsed = JSON.parse(text);
        } catch {
          parsed = null;
        }
      }
      ok = r.ok && parsed != null && typeof parsed === "object";
      if (ok) {
        j = parsed;
        break;
      }
      j = parsed;
    } catch {
      j = null;
      ok = false;
    }
  }
  _healthBannerCache = { t: Date.now(), j, ok };
  paintGlobalApiBanner(el, j, ok);
}

function paintGlobalApiBanner(el, j, fetchOk) {
  if (!fetchOk || !j || typeof j !== "object") {
    el.hidden = false;
    el.className = "globalApiBanner globalApiBannerErr";
    const origin = `${location.protocol}//${location.host}`;
    if (isFileProtocol()) {
      el.innerHTML = `<strong>Opened as a file.</strong> The API lives in <code>server.py</code> — use the app from the server URL instead. In Finder, double‑click <code>JohnsStockApp.command</code>, then open <strong><code>http://localhost:8844</code></strong> (or the port printed in Terminal). <a href="/README.html#checklist-daily">Checklist</a>`;
    } else if (isAppDevOnLocalhost()) {
      el.innerHTML = `<strong>API unreachable at <code>${esc(origin)}</code>.</strong> Start <code>server.py</code> from this project folder (double‑click <code>JohnsStockApp.command</code>) and keep that Terminal window open. Port must match the URL (see <code>PORT</code> in <code>.env</code>). <a href="/README.html#checklist-daily">Checklist</a> · <a href="/api/health" target="_blank" rel="noopener">Try <code>/api/health</code></a>`;
    } else {
      el.innerHTML = `<div class="globalApiBannerRow">
      <div class="globalApiBannerBody">
        <strong>Could not load status from the server (yet).</strong>
        <span class="sml"> On <strong>free</strong> cloud hosting the app may need <strong>~30–60 seconds</strong> to wake, or the network hiccuped. This is <em>not</em> something you fix in Terminal. Wait, then <strong>reload the page</strong> or use <strong>Retry</strong>. <a href="/api/health" target="_blank" rel="noopener">Open <code>/api/health</code> in a new tab</a> — you should see <code>ok: true</code> in the JSON when the server is ready.</span>
      </div>
      <button type="button" class="btn ghost smlBtn" id="hostedApiBannerRetry">Retry</button>
    </div>`;
      const btn = el.querySelector("#hostedApiBannerRetry");
      if (btn instanceof HTMLButtonElement) {
        btn.addEventListener("click", () => {
          _healthBannerCache = null;
          void refreshGlobalApiBanner();
        });
      }
    }
    return;
  }
  const rev = j.api_revision;
  const hasLlm = j.llm_commentary && typeof j.llm_commentary === "object";
  const revNum = typeof rev === "number" ? rev : NaN;
  const stale = !hasLlm || !Number.isFinite(revNum) || revNum < MIN_API_REVISION;
  if (stale) {
    el.hidden = false;
    el.className = "globalApiBanner globalApiBannerErr";
    const cur = Number.isFinite(revNum) ? String(revNum) : "missing";
    const hp = location.port ? `${location.protocol}//${location.hostname}:${location.port}` : `${location.protocol}//${location.host}`;
    if (isHostedWebApp()) {
      el.innerHTML = `<strong>Server API looks older than this page needs.</strong> <span class="sml">Redeploy the latest <code>server.py</code> on your host (e.g. <strong>Render → Manual deploy</strong> from the current <code>main</code> branch), then hard-reload. Last seen <code>api_revision</code>: <code>${esc(cur)}</code> (need ≥ <strong>${MIN_API_REVISION}</strong> and <code>llm_commentary</code> in <a href="${esc(hp)}/api/health" target="_blank" rel="noopener">/api/health</a>).</span>`;
    } else {
      el.innerHTML = `<strong>The server on this port is still an old copy.</strong> Close any Terminal that was running the app, then in <strong>Finder</strong> go to your project folder and <strong>double‑click <code>JohnsStockApp.command</code></strong> — it now <strong>stops the old listener for you</strong> and starts the new server. Then open <a href="${esc(hp)}/api/health" target="_blank" rel="noopener"><code>/api/health</code></a> and check for <code>api_revision</code> ≥ <strong>${MIN_API_REVISION}</strong> and <code>llm_commentary</code> (yours was: <code>${esc(cur)}</code>). Same folder: <code>OPEN_APP_URL.txt</code>.`;
    }
    return;
  }
  const cfg = j.llm_commentary?.configured;
  const anyKey =
    cfg &&
    typeof cfg === "object" &&
    Boolean(cfg.openai || cfg.anthropic || cfg.google || cfg.xai);
  if (!anyKey) {
    el.hidden = true;
    el.textContent = "";
    el.className = "globalApiBanner";
    return;
  }
  el.hidden = true;
  el.textContent = "";
  el.className = "globalApiBanner";
}

function route() {
  migrateLegacyPortfolioFromV1();
  const { pathname, sp } = parseLocationHash();
  if (!pathname.startsWith("/portfolio") && _pfSharedBundle != null) {
    _pfSharedBundle = null;
  }
  document.querySelectorAll("[data-route]").forEach((a) => {
    if (!(a instanceof HTMLAnchorElement)) return;
    const href = a.getAttribute("href") || "";
    if (pathname.startsWith("/symbol")) a.removeAttribute("aria-current");
    else if (href === `#${pathname}`) a.setAttribute("aria-current", "page");
    else a.removeAttribute("aria-current");
  });
  const v = $("view");
  if (!v) return;
  if (pathname.startsWith("/portfolio")) v.innerHTML = portfolioHtml();
  else if (pathname.startsWith("/watchlist")) v.innerHTML = watchlistHtml();
  else if (pathname.startsWith("/notes")) v.innerHTML = notesHtml();
  else if (pathname.startsWith("/symbol")) {
    const sym = (sp.get("sym") || "").trim();
    if (!sym) {
      v.innerHTML = `<p class="muted">Missing symbol. <a class="backLink" href="#/search">Back to search</a></p>`;
      wire();
      return;
    }
    const exRaw = (sp.get("ex") || "").trim();
    const nm = (sp.get("nm") || "").trim();
    const ccyHint = (sp.get("ccy") || "").trim();
    v.innerHTML = instrumentHtml(sym, exRaw || null, nm, ccyHint);
  } else v.innerHTML = searchHtml();
  wire();
  void refreshGlobalApiBanner();
  updateFamilyNavHrefs();
}

function searchHtml() {
  return `
    <h1 class="h1">Search</h1>
    <p class="lead">Type a ticker or company. Results show <strong>country · exchange · currency</strong> when known; use the filters to narrow listings.
      <span class="sml muted">Indian names: pick the row that says <strong>NSE</strong> or <strong>BSE</strong> as you want — NSE rows are listed first when both exist.</span>
      <a class="inlineHealth" href="/api/health" target="_blank" rel="noopener" title="Opens in a new tab">Open API health (new tab)</a>
      · <a href="/README.html#checklist-daily">Start server checklist</a>
      · <a href="/README.html#python3-explained">What is python3?</a></p>
    <input class="in" id="q" placeholder="e.g. NVDA, Reliance, VUSA" autocomplete="off" />
    <div id="searchFilt" class="searchFilt" hidden></div>
    <div id="results" class="mt"></div>
    <div class="card2 mt" id="aiAskCard">
      <div class="h3">Ask AI <span class="sml muted">(general)</span></div>
      <p class="sml muted">Uses <code>GET /api/ai-ask?q=…</code> — same LLM keys as Fundamentals commentary. Good for definitions (e.g. RS55, RSI) or how to read a metric; it does <strong>not</strong> pull live prices unless you paste them.</p>
      <textarea id="aiAskQ" class="ta" rows="3" placeholder="e.g. What is RS55? How should I read RSI with moving averages?"></textarea>
      <div class="aiComRow mt">
        <button type="button" class="btn" id="aiAskBtn">Ask</button>
        <a class="btn ghost" href="https://copilot.microsoft.com/" target="_blank" rel="noopener noreferrer" title="Opens Microsoft’s site in a new tab — cannot be embedded here">Copilot (new tab)</a>
      </div>
      <p class="sml muted mt">Copilot’s site cannot run inside this app (browsers block that for security). Use the button above, then paste your question there if you like.</p>
      <div id="aiAskOut" class="aiCommentaryOut mt" aria-live="polite"></div>
    </div>
    <p class="sml muted mt">Pick a row to open the <strong>instrument</strong> page (quote, tabs, chart, add to portfolio).</p>
    <div class="st" role="status"><span id="stL"></span><span id="stR"></span></div>`;
}

function instrumentHtml(sym, ex, nm, ccyHint) {
  const title = nm ? `${esc(sym)} · ${esc(nm)}` : esc(sym);
  const sub = [ex || "—", ccyHint || ""].filter(Boolean).join(" · ");
  return `
    <div class="instrumentView">
    <div class="instrStickyTop">
    <div class="instrNav">
      <a class="backLink" href="#/search">← Search</a>
      <a class="backLink muted sml" href="#/watchlist">Watchlist</a>
      <a class="backLink muted sml" href="${hrefPortfolio()}">Portfolio</a>
    </div>
    <header class="instrHead">
      <h1 class="h1" id="instrTitle">${title}</h1>
      <p class="sml muted" id="instrSub">${esc(sub || "—")}</p>
    </header>
    <div class="tabBar" role="tablist" aria-label="Instrument">
      <button type="button" class="tabBtn" role="tab" aria-selected="true" aria-controls="tabInstrOv" id="tabBtnOv" data-tab="ov">Overview</button>
      <button type="button" class="tabBtn" role="tab" aria-selected="false" aria-controls="tabInstrTech" id="tabBtnTech" data-tab="tech">Technicals &amp; chart</button>
      <button type="button" class="tabBtn" role="tab" aria-selected="false" aria-controls="tabInstrFun" id="tabBtnFun" data-tab="fun">Fundamentals</button>
    </div>
    </div>
    <details class="instrChartDetails card2 mt" id="instrChartDetails">
    <summary class="instrChartDetailsSum" id="instrChartDetailsSum">Price history <span class="muted sml">— click to show chart, overlays, and volume</span></summary>
    <div class="instrChartDock" id="instrChartDock">
      <p class="muted sml" id="chartSt">Expand this section to load the chart, or open the <strong>Technicals</strong> tab for indicators.</p>
      <p class="chartStaleBanner" id="chartStaleBanner" hidden role="status" aria-live="polite"></p>
      <div class="chartTypeBar" role="toolbar" aria-label="Chart type">
        <span class="muted sml">Type</span>
        <button type="button" class="btn ghost smlBtn" data-chart-type="area" aria-pressed="true">Area</button>
        <button type="button" class="btn ghost smlBtn" data-chart-type="candle" aria-pressed="false">Candles</button>
        <button type="button" class="btn ghost smlBtn" data-chart-type="heikin" aria-pressed="false">Heikin Ashi</button>
      </div>
      <div class="chartOlBar" role="group" aria-label="Chart overlays">
        <span class="muted sml">Overlays</span>
        <label class="chkLab"><input type="checkbox" id="olMa20" checked /> SMA20</label>
        <label class="chkLab"><input type="checkbox" id="olMa50" checked /> SMA50</label>
        <label class="chkLab"><input type="checkbox" id="olBb" /> Bollinger (20, 2σ)</label>
        <label class="chkLab"><input type="checkbox" id="olVol" checked /> Volume</label>
        <label class="chkLab"><input type="checkbox" id="olTrend" checked /> Trendline</label>
      </div>
      <div class="histCanvasWrap"><canvas id="histCanvas" class="histCanvas" aria-label="Price history chart"></canvas></div>
      <div id="volDock" class="volDock" hidden>
        <div class="volDockHd sml muted">Volume (same dates as price)</div>
        <div class="volCanvasWrap"><canvas id="histVolCanvas" class="histVolCanvas" aria-label="Trading volume"></canvas></div>
      </div>
      <p class="sml muted">Data from <code>/api/history</code> (India: Yahoo → EODHD → Alpha Vantage → Twelve; others: Yahoo → EODHD → Twelve → Alpha Vantage). If live data fails, a <strong>session snapshot</strong> may be shown. For information only.</p>
    </div>
    </details>
    <section id="tabInstrOv" class="tabPanel" role="tabpanel" aria-labelledby="tabBtnOv">
      <div id="instrPfHold" class="card2 instrPfHold mt" hidden role="region" aria-label="Your portfolio holding"></div>
      <div id="instrQuoteMount" class="mt"><p class="muted">Loading quote…</p></div>
      <div class="card2 mt">
        <div class="h3">News</div>
        <div id="instrNews" class="instrFeed"><p class="muted sml">Loading headlines…</p></div>
      </div>
      <div class="card2 mt">
        <div class="h3">Corporate actions</div>
        <div id="instrCorp" class="instrFeed"><p class="muted sml">Loading calendar…</p></div>
      </div>
      <div class="card2 mt" id="addPfCard">
        <div class="h3">Add to portfolio</div>
        <p class="sml muted">Qty is required. Currency defaults from the search row or quote when available.</p>
        <label class="lbl mt">Broker
          <select class="in" id="addPfBroker">
            ${PF_BROKER_IDS.map((id) => `<option value="${esc(id)}">${esc(PF_BROKER_LABEL[id])}</option>`).join("")}
          </select>
        </label>
        <div class="addPfGrid">
          <label class="lbl">Qty <input class="in" id="addPfQty" inputmode="decimal" value="1" /></label>
          <label class="lbl">Avg buy (optional) <input class="in" id="addPfAvg" inputmode="decimal" placeholder="0" /></label>
          <label class="lbl">Currency <input class="in" id="addPfCcy" placeholder="e.g. EUR, INR" /></label>
        </div>
        <button type="button" class="btn mt" id="addPfBtn">Add row</button>
        <p class="sml muted mt" id="addPfMsg" role="status" aria-live="polite"></p>
      </div>
      <div class="card2 mt" id="addWlCard">
        <div class="h3">Watchlist</div>
        <p class="sml muted">Save this listing for quick access (stored in this browser only).</p>
        <button type="button" class="btn mt" id="addWlBtn">Add to watchlist</button>
        <p class="sml muted mt" id="addWlMsg" role="status" aria-live="polite"></p>
      </div>
    </section>
    <section id="tabInstrTech" class="tabPanel" role="tabpanel" aria-labelledby="tabBtnTech" hidden>
      <p class="sml muted">Indicators use the same daily series as the chart. Expand <strong>Price history</strong> (above) when you want the chart on screen; the series loads when you open that section or this tab.</p>
      <div id="techMount" class="techMount mt" aria-live="polite"></div>
    </section>
    <section id="tabInstrFun" class="tabPanel" role="tabpanel" aria-labelledby="tabBtnFun" hidden>
      <div id="funMount" class="funMount"><p class="muted sml">Open this tab after the quote loads, or switch away and back.</p></div>
    </section>
    <div class="st" role="status"><span id="stL"></span><span id="stR"></span></div>
    </div>`;
}

function notesHtml() {
  return `
    <h1 class="h1">Notes</h1>
    <p class="lead">Stored only in this browser.</p>
    <textarea class="ta" id="notes" rows="10" placeholder="…"></textarea>
    <div class="row"><button type="button" class="btn" id="notesSave">Save</button></div>
    <div class="st" role="status"><span id="stL"></span><span id="stR"></span></div>`;
}

function portfolioHtml() {
  let banner = "";
  try {
    const n = sessionStorage.getItem(LEGACY_NOTICE);
    if (n) {
      sessionStorage.removeItem(LEGACY_NOTICE);
      banner = `<div class="migrateBanner" role="status">Restored <strong>${esc(n)}</strong> holding(s) from the previous app save on this device.</div>`;
    }
  } catch {
    /* ignore */
  }
  return `
    ${banner}
    <h1 class="h1">Portfolio</h1>
    <div id="pfSharedBanner" class="card2 mt" hidden aria-live="polite"></div>
    <div id="pfFamilyTools" class="card2 mt pfFamilyTools" hidden role="region" aria-label="Family read-only actions">
      <p class="sml muted" style="margin:0 0 8px 0">Tap <strong>Refresh market prices</strong> to pull live quotes (waits for the server). <strong>Reload owner snapshot</strong> uses the last data the owner published. No edits — read-only.</p>
      <div class="rowgap pfFamilyToolBtns" style="display:flex;flex-wrap:wrap;gap:10px;">
        <button type="button" class="btn" id="btnFamilyRefPx">Refresh market prices</button>
        <button type="button" class="btn ghost" id="btnFamilySnap">Reload owner snapshot</button>
      </div>
    </div>
    <div class="pfTopSummary" aria-label="Total and breakdown in euro">
      <div id="pfGrandTotalMount" class="pfGrandTotalMount" aria-live="polite" hidden></div>
      <div id="pfCombinedEur" class="mt" hidden></div>
    </div>
    <div class="brokerBar brokerBarWide" role="tablist" aria-label="Portfolio ledgers">
      <button type="button" class="brokerTab brokerTabMain" role="tab" data-main-tab="t212" aria-selected="true">Trading 212</button>
      <button type="button" class="brokerTab brokerTabMain" role="tab" data-main-tab="crypto" aria-selected="false">Crypto</button>
      <button type="button" class="brokerTab brokerTabMain" role="tab" data-main-tab="zerodha" aria-selected="false">Zerodha</button>
      <button type="button" class="brokerTab brokerTabMain" role="tab" data-main-tab="etoro" aria-selected="false">eToro</button>
      <button type="button" class="brokerTab brokerTabMain" role="tab" data-main-tab="mf" aria-selected="false">Mutual funds</button>
      <button type="button" class="brokerTab brokerTabMain" role="tab" data-main-tab="insurance" aria-selected="false">Insurances</button>
      <button type="button" class="brokerTab brokerTabMain" role="tab" data-main-tab="fixed_deposit" aria-selected="false">Fixed deposits</button>
      <span class="sml muted pfBrokerCap" id="pfBrokerCap" aria-live="polite"></span>
    </div>
    <div id="pfSubLedgerBar" class="pfSubLedgerBar brokerBarWide" role="tablist" aria-label="Sub-ledger" hidden></div>
    <div id="pfLedgerTotals" class="mt" hidden aria-live="polite"></div>
    <div id="pfT212EurMount"></div>
    <div id="pfCharts" class="pfCharts" hidden></div>
    <div id="tbl" class="mt"></div>
    <section id="pfManage" class="pfManageSection card2 mt" aria-labelledby="pfManageHd">
      <h2 class="h2" id="pfManageHd">Import &amp; manual rows</h2>
      <div class="grid2 mt">
        <div>
          <label class="lbl">CSV import</label>
          <input type="file" id="csv" accept=".csv,text/csv,.txt" class="in" />
          <div class="rowgap mt" style="display:flex;flex-wrap:wrap;gap:8px;">
            <button type="button" class="btn" id="btnImp">Import</button>
            <button type="button" class="btn ghost" id="btnTpl">Download CSV template</button>
          </div>
        </div>
        <div>
          <div class="lbl" id="lblPfAddRow">Add row</div>
          <div id="pfAddFieldsMount" class="pfAddRowGrid mt" role="group" aria-labelledby="lblPfAddRow"></div>
          <div class="pfAddFieldBtnRow mt">
            <button type="button" class="btn pfAddBtn" id="btnAdd">Add</button>
          </div>
        </div>
      </div>
      <div class="rowgap mt pfToolRow" style="display:flex;flex-wrap:wrap;gap:10px;">
        <button type="button" class="btn" id="btnT212Sync" title="Replaces T212 + Crypto ledgers with open positions (read-only API)">Sync from Trading 212</button>
        <span class="sml muted pfT212SyncStamp" id="pfT212SyncTime" role="status" hidden></span>
        <button type="button" class="btn" id="btnPfPublish" title="Uploads this device’s portfolio to the server so family can open a read-only link (needs write key in Render)">Publish for family (server)</button>
        <button type="button" class="btn ghost" id="btnRef">Refresh prices</button>
        <button type="button" class="btn ghost" id="btnPfExport">Export CSV</button>
        <button type="button" class="btn ghost" id="btnPfJsonExport" title="All ledgers in one file — use to move data to phone or another browser">Backup all ledgers (JSON)</button>
        <button type="button" class="btn ghost" id="btnPfJsonRestore" title="Replaces entire portfolio from a JSON backup">Restore backup (JSON)</button>
        <input type="file" id="pfJsonImp" accept="application/json,.json" hidden />
        <button type="button" class="btn ghost" id="btnClrPf">Clear all</button>
      </div>
      <div class="card2 mt pfSafetyCard" style="margin-top:14px;">
        <div class="h3">Safety PIN</div>
        <div class="rowgap" style="display:flex;flex-wrap:wrap;gap:8px;align-items:flex-end;">
          <label class="lbl" style="margin:0">New PIN (3 digits)
            <input class="in" id="pfPinNew" maxlength="3" inputmode="numeric" placeholder="e.g. 842" autocomplete="off" />
          </label>
          <label class="lbl" style="margin:0">Confirm
            <input class="in" id="pfPinConf" maxlength="3" inputmode="numeric" placeholder="repeat" autocomplete="off" />
          </label>
          <button type="button" class="btn ghost smlBtn" id="btnSavePin">Save PIN</button>
        </div>
      </div>
    </section>
    <div class="st" role="status"><span id="stL"></span><span id="stR"></span></div>`;
}

let _pfPortfolioViewBound = false;

function bindPfPortfolioViewOnce() {
  if (_pfPortfolioViewBound) return;
  _pfPortfolioViewBound = true;
  $("view")?.addEventListener("click", onPortfolioViewClick);
}

function onPortfolioViewClick(ev) {
  if (!String(location.hash || "").toLowerCase().startsWith("#/portfolio")) return;
  const t = ev.target;
  if (!(t instanceof Node)) return;
  const mainBtn = t instanceof Element ? t.closest("[data-main-tab]") : null;
  if (mainBtn instanceof HTMLButtonElement) {
    const m = mainBtn.dataset.mainTab;
    if (!isPfMainTabId(m)) return;
    if (m === PF_MAIN_TAB_MF) setActiveBroker(getMfSubBroker());
    else if (m === PF_MAIN_TAB_CRYPTO) setActiveBroker(PF_CRYPTO);
    else if (isPfBrokerId(m)) setActiveBroker(m);
    renderPf();
    const lab =
      m === PF_MAIN_TAB_MF
        ? "Mutual funds"
        : m === PF_INSURANCE
          ? "Insurances"
          : m === PF_MAIN_TAB_CRYPTO
            ? PF_BROKER_LABEL[PF_CRYPTO]
            : PF_BROKER_LABEL[m] || m;
    status(lab, document.documentElement.dataset.theme || "");
    return;
  }
  const mfSub = t instanceof Element ? t.closest("[data-mf-sub]") : null;
  if (mfSub instanceof HTMLButtonElement) {
    const sub = mfSub.dataset.mfSub;
    if (sub === PF_MF_COIN || sub === PF_MF_KUVERA) {
      setActiveBroker(sub);
      renderPf();
      status(PF_BROKER_LABEL[sub], document.documentElement.dataset.theme || "");
    }
    return;
  }
  const insBtn = t instanceof Element ? t.closest("[data-insco]") : null;
  if (insBtn instanceof HTMLButtonElement) {
    const id = insBtn.dataset.insco || "";
    if (!isPfInsCoId(id)) return;
    setPfInsuranceCompany(id);
    renderPf();
    status(PF_INS_CO_LABEL[id] || id, document.documentElement.dataset.theme || "");
  }
}

function watchlistHtml() {
  return `
    <h1 class="h1">Watchlist</h1>
    <p class="lead">Tickers you save from an instrument’s <strong>Add to watchlist</strong> button. Data stays in this browser. Each row opens the same instrument workspace as Search (quote, chart, news, portfolio add). <strong>Removing a row</strong> or <strong>clearing the list</strong> asks for your safety PIN (set on the <strong>Portfolio</strong> page; default <strong>000</strong>).</p>
    <div class="rowgap mt" style="display:flex;flex-wrap:wrap;gap:10px;">
      <button type="button" class="btn ghost" id="btnWlClr">Clear watchlist</button>
    </div>
    <div id="wlTbl" class="mt"></div>
    <div class="st" role="status"><span id="stL"></span><span id="stR"></span></div>`;
}

function wire() {
  const { pathname, sp } = parseLocationHash();

  if (pathname.startsWith("/symbol")) {
    const sym = (sp.get("sym") || "").trim();
    if (!sym) return;
    const exRaw = (sp.get("ex") || "").trim();
    wireInstrument(
      sym,
      exRaw || null,
      (sp.get("nm") || "").trim(),
      (sp.get("ccy") || "").trim(),
      sp.get("pf") === "1",
    );
    return;
  }

  if (pathname.startsWith("/notes")) {
    const ta = $("notes");
    try {
      if (ta instanceof HTMLTextAreaElement) ta.value = localStorage.getItem(K.notes) || "";
    } catch {
      /* ignore */
    }
    $("notesSave")?.addEventListener("click", () => {
      try {
        if (ta instanceof HTMLTextAreaElement) localStorage.setItem(K.notes, ta.value);
        status("Saved");
      } catch {
        status("Save failed");
      }
    });
    const saveDeb = debounce(() => {
      try {
        if (ta instanceof HTMLTextAreaElement) localStorage.setItem(K.notes, ta.value);
        status("Autosaved");
      } catch {
        status("Save failed");
      }
    }, 400);
    ta?.addEventListener("input", saveDeb);
    status("Notes", document.documentElement.dataset.theme || "");
    return;
  }

  if (pathname.startsWith("/watchlist")) {
    renderWl();
    $("btnWlClr")?.addEventListener("click", () => {
      if (!requireDangerPin("clear the entire watchlist")) return;
      if (!confirm("Remove every symbol from your watchlist? This cannot be undone.")) return;
      saveWl(emptyWl());
      renderWl();
      status("Watchlist cleared");
    });
    status("Watchlist", document.documentElement.dataset.theme || "");
    return;
  }

  if (pathname.startsWith("/portfolio")) {
    $("btnImp")?.addEventListener("click", async () => {
      const f = $("csv")?.files?.[0];
      if (!f) {
        status("Pick a file");
        return;
      }
      const t = await f.text();
      const bundle = loadPfBundle();
      const b = getActiveBroker();
      let rows = [];
      if (b === PF_INSURANCE) {
        rows = parseInsuranceCsv(t);
        if (!rows.length) rows = parseCsvSmart(t);
      } else if (b === PF_FIXED_DEPOSIT) {
        rows = parseFdCsv(t);
        if (!rows.length) rows = parseCsvSmart(t);
      } else {
        rows = parseCsvSmart(t);
      }
      if (!rows.length) {
        status(
          b === PF_ZERODHA
            ? "0 rows — Zerodha gives XLSX: open in Excel/LibreOffice, Save As CSV (UTF-8), then import. Holdings need Symbol + Quantity (+ Avg/LTP). Or use the app CSV template with a currency column."
            : b === PF_INSURANCE || b === PF_FIXED_DEPOSIT
              ? "0 rows — use the app’s CSV template for this tab, or check headers (insurance: policyName + currency; FD: fdBank + fdName + principal + currency)."
              : b === PF_CRYPTO
                ? "0 rows — use Sync from Trading 212 or the app CSV template (symbol, currency, qty, …)."
                : "0 rows — use Trading 212 history CSV, the app template (symbol,currency,qty,…), a Zerodha holdings CSV, or Sync from Trading 212.",
        );
        return;
      }
      bundle.brokers[b].rows.push(...rows);
      savePfBundle(bundle);
      renderPf();
      if (b === PF_T212) {
        status(
          `Imported ${rows.length} → ${PF_BROKER_LABEL[b]} — CSV keeps broker tickers and currencies as in the file. For display symbols, exchange, and per-listing currency, click "Sync from Trading 212" (replaces T212 + Crypto ledgers with live API data).`,
        );
      } else {
        status(`Imported ${rows.length} → ${PF_BROKER_LABEL[b]}`);
      }
    });
    $("btnAdd")?.addEventListener("click", () => {
      const bundle = loadPfBundle();
      const b = getActiveBroker();
      if (b === PF_INSURANCE) {
        const pn = val("fPolName");
        const ccy = normalizeCcyForFx(val("fPolCcy").toUpperCase());
        if (!pn || !ccy) {
          status("Need policy name and currency");
          return;
        }
        const row = {
          policyName: pn,
          policyNo: val("fPolNo"),
          purchaseDate: val("fPolPur"),
          valueAtPurchase: num(val("fPolV0")),
          growthPct: num(val("fPolGr")),
          currentValue: num(val("fPolCur")),
          ccy,
          insCompany: getPfInsuranceCompany(),
          payments: [],
        };
        ensurePfRowId(row);
        bundle.brokers[PF_INSURANCE].rows.push(row);
        savePfBundle(bundle);
        renderPf();
        status(`Added policy → ${PF_INS_CO_LABEL[getPfInsuranceCompany()]}`);
        return;
      }
      if (b === PF_FIXED_DEPOSIT) {
        const bank = val("fFdBank");
        const name = val("fFdName");
        const ccy = normalizeCcyForFx(val("fFdCcy").toUpperCase());
        const pr = num(val("fFdPrin"));
        if (!bank || !name || !ccy || pr <= 0) {
          status("Need bank, deposit name, currency, and principal > 0");
          return;
        }
        const row = {
          fdBank: bank,
          fdName: name,
          fdRef: val("fFdRef"),
          openDate: val("fFdOpen"),
          principal: pr,
          ratePct: num(val("fFdRate")),
          currentValue: num(val("fFdCur")) || pr,
          maturityDate: val("fFdMat"),
          ccy,
          fdCountry: val("fFdCtry"),
        };
        ensurePfRowId(row);
        bundle.brokers[PF_FIXED_DEPOSIT].rows.push(row);
        savePfBundle(bundle);
        renderPf();
        status("Added fixed deposit");
        return;
      }
      const sym = val("fSym");
      const ex = val("fEx");
      const ccy = normalizeCcyForFx(val("fCcy").toUpperCase() || (b === PF_CRYPTO ? "USD" : ""));
      const qty = num(val("fQty"));
      if (!sym || !ccy || qty <= 0) {
        status("Need symbol, currency, qty");
        return;
      }
      bundle.brokers[b].rows.push({
        sym: b === PF_CRYPTO ? sym.toUpperCase() : sym,
        ex,
        ccy,
        qty,
        avg: num(val("fAvg")),
        last: num(val("fLast")),
        nm: val("fNm"),
      });
      savePfBundle(bundle);
      if (b === PF_CRYPTO) {
        void (async () => {
          const b2 = loadPfBundle();
          try {
            await applyLiveQuotesToRowsForBroker(b2, PF_CRYPTO);
            savePfBundle(b2);
          } finally {
            renderPf();
            status("Added to Crypto (T212) — last column updates when a live quote is returned. Use Refresh for a full re-fetch.");
          }
        })();
        return;
      }
      renderPf();
      status(`Added → ${PF_BROKER_LABEL[b]}`);
    });
    $("btnTpl")?.addEventListener("click", () => {
      const b = getActiveBroker();
      let csv = "symbol,name,exchange,currency,qty,avgPrice,lastPrice\nTSLA,Tesla Inc,NASDAQ,USD,2,180,195\n";
      let name = "johnsstockapp-template.csv";
      if (b === PF_INSURANCE) {
        csv =
          "insCompany,policyName,policyNo,purchaseDate,valueAtPurchase,growthPct,currentValue,currency\n" +
          "sbi_life,Smart Wealth,SW123,2020-01-15,50000,7,62000,INR\n";
        name = "johnsstockapp-insurance-template.csv";
      } else if (b === PF_FIXED_DEPOSIT) {
        csv =
          "fdBank,fdCountry,fdName,fdRef,openDate,principal,ratePct,currentValue,maturityDate,currency\n" +
          "SBI,India,12-month FD,R1,2024-04-01,100000,7.1,105000,2025-04-01,INR\n";
        name = "johnsstockapp-fd-template.csv";
      }
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
      a.download = name;
      a.click();
      setTimeout(() => URL.revokeObjectURL(a.href), 2000);
      status("Template downloaded");
    });
    $("btnT212Sync")?.addEventListener("click", async () => {
      status("Syncing from Trading 212…");
      try {
        const j = await applyTrading212SyncToBundle();
        renderPf();
        const ns = Number(j.n_t212 ?? 0) || 0;
        const nc = Number(j.n_t212_crypto ?? 0) || 0;
        const when = typeof j.fetched_at === "string" && j.fetched_at.trim() ? ` · server ${formatServerIsoLocal(j.fetched_at)}` : "";
        status(`T212 sync · ${ns} stock/ETF, ${nc} crypto${when}`);
      } catch (e) {
        status(e instanceof Error ? e.message : String(e));
      }
    });
    $("btnRef")?.addEventListener("click", refreshPf);
    $("btnFamilyRefPx")?.addEventListener("click", () => void familyRefreshAllMarketPrices());
    $("btnFamilySnap")?.addEventListener("click", () => void familyReloadOwnerSnapshot());
    $("btnPfExport")?.addEventListener("click", () => exportPfCsv());
    $("btnPfJsonExport")?.addEventListener("click", () => exportPfBackupJson());
    $("btnPfJsonRestore")?.addEventListener("click", () => {
      const inp = $("pfJsonImp");
      if (inp instanceof HTMLInputElement) inp.click();
    });
    $("pfJsonImp")?.addEventListener("change", () => void importPfBackupFromFile($("pfJsonImp")));
    $("btnSavePin")?.addEventListener("click", () => {
      const a = $("pfPinNew");
      const c = $("pfPinConf");
      const p1 = a instanceof HTMLInputElement ? a.value.trim() : "";
      const p2 = c instanceof HTMLInputElement ? c.value.trim() : "";
      if (!/^\d{3}$/.test(p1)) {
        status("PIN must be exactly 3 digits");
        return;
      }
      if (p1 !== p2) {
        status("PIN fields do not match");
        return;
      }
      try {
        localStorage.setItem(K.dangerPin, p1);
        if (a instanceof HTMLInputElement) a.value = "";
        if (c instanceof HTMLInputElement) c.value = "";
        status("Safety PIN saved");
      } catch {
        status("Could not save PIN");
      }
    });
    $("btnClrPf")?.addEventListener("click", () => {
      const b = getActiveBroker();
      const lab = PF_BROKER_LABEL[b];
      if (!requireDangerPin(`clear every ${lab} row`)) return;
      const clrMsg =
        b === PF_INSURANCE
          ? `Delete ALL insurance rows on this device (every provider tab: SBI Life, Birla, Allianz, VRK, Other)? This cannot be undone.`
          : `Delete all ${lab} rows on this device? This cannot be undone.`;
      if (!confirm(clrMsg)) return;
      const bundle = loadPfBundle();
      bundle.brokers[b].rows = [];
      savePfBundle(bundle);
      renderPf();
      status(`Cleared ${lab}`);
    });
    $("btnPfPublish")?.addEventListener("click", () => {
      void publishPortfolioToServer();
    });
    bindPfPortfolioViewOnce();
    setActiveBroker(getActiveBroker());
    void (async () => {
      await applyPortfolioSharedFromHash(sp);
      renderPf();
      updateFamilyNavHrefs();
      status("Portfolio", document.documentElement.dataset.theme || "");
    })();
    return;
  }

  /* search */
  const inp = $("q");
  const res = $("results");
  if (!(inp instanceof HTMLInputElement) || !res) return;

  async function runSearchQuery() {
    const q = inp.value.trim();
    if (!q) {
      _searchAbort?.abort();
      _searchAbort = null;
      res.innerHTML = "";
      try {
        sessionStorage.removeItem(SEARCH_LAST_Q);
      } catch {
        /* ignore */
      }
      status("Ready");
      return;
    }
    status("Searching…");
    _searchAbort?.abort();
    _searchAbort = new AbortController();
    const { signal } = _searchAbort;
    try {
      const [r, t212r] = await Promise.all([
        fetch(`/api/search?q=${encodeURIComponent(q)}&limit=25`, { signal }),
        fetch(`/api/t212/instruments?q=${encodeURIComponent(q)}&limit=18&region=all`, { signal }).catch(() => null),
      ]);
      let t212j = null;
      if (t212r && t212r.ok) {
        try {
          t212j = await t212r.json();
        } catch {
          t212j = null;
        }
      }
      const raw = await r.text();
      let j;
      try {
        j = JSON.parse(raw);
      } catch {
        res.innerHTML = `<p class="err">Server returned non-JSON — the app server may not be running.</p>
          <p class="muted sml"><a href="/README.html#checklist-daily">Open the start checklist</a> · <a href="/README.html#python3-explained">What is python3?</a></p>
          <pre class="sml">${esc(raw.slice(0, 400))}</pre>`;
        status("Search failed");
        return;
      }
      if (!r.ok) {
        const detail = String(j.detail || j.error || r.status);
        const hint = j.hint ? `<p class="muted sml">${esc(j.hint)}</p>` : "";
        const r429 =
          /429|too many requests/i.test(detail) || /429|too many requests/i.test(String(j.hint || ""));
        const tip429 = r429
          ? `<div class="box tip429" role="note"><strong>Right now:</strong> wait a few minutes, clear the box, type <code>NVDA</code> (ticker), pause, then search. After Yahoo the server tries <strong>Twelve Data</strong>, <strong>EODHD</strong>, then <strong>Alpha Vantage</strong> — set keys in <code>.env</code> (Twelve must be its <em>own</em> key) and <strong>restart</strong> <code>server.py</code>.</div>`
          : "";
        res.innerHTML = `<p class="err">${esc(detail)}</p>${tip429}${hint}<p class="muted sml">
          <a class="inlineHealth" href="/api/health" target="_blank" rel="noopener">Open API health (new tab)</a> ·
          <a href="/README.html#troubleshooting">Troubleshooting</a> ·
          <a href="/README.html#checklist-daily">Start checklist</a></p>`;
        status("Search failed");
        return;
      }
      const list = Array.isArray(j) ? j : [];
      res.innerHTML = "";
      const filtEl = $("searchFilt");
      if (filtEl) {
        filtEl.innerHTML = "";
        filtEl.hidden = true;
      }
      const t212Has = Boolean(t212j && t212j.ok && Array.isArray(t212j.matches) && t212j.matches.length);
      if (!list.length && !t212Has) {
        res.innerHTML = `<p class="muted">No results (Yahoo / FMP and Trading 212 cache).</p>`;
        status("0 results");
        return;
      }
      if (!list.length && t212Has) {
        res.innerHTML = `<p class="muted">No Yahoo / FMP hits — showing Trading 212 symbols below.</p>`;
      }
      const ul = document.createElement("div");
      ul.className = "list";
      for (const row of list.slice(0, 25)) {
        const sym = String(row.symbol || "").trim();
        if (!sym) continue;
        const ex = String(row.exchangeShortName || row.exchange || "").trim();
        const nm = String(row.name || "").trim();
        const country = String(row.country || "").trim();
        const ccy = String(row.currency || "").trim();
        const qt = String(row.quoteType || "").trim();
        const chip =
          qt && /ETF|MUTUALFUND|INDEX|FUND|CRYPTO/i.test(qt)
            ? ` <span class="qtChip" title="Quote type">${esc(qt)}</span>`
            : "";
        const meta = [country, ex, ccy].filter(Boolean).join(" · ") || "—";
        const b = document.createElement("button");
        b.type = "button";
        b.className = "hit";
        b.dataset.country = country;
        b.dataset.exch = ex;
        b.dataset.ccy = ccy;
        b.innerHTML = `<div class="hitTop"><strong>${esc(sym)}</strong>${chip}</div>
          <div class="hitName sml">${esc(nm)}</div>
          <div class="hitMeta sml muted">${esc(meta)}</div>`;
        b.addEventListener("click", () => {
          location.hash = instrumentHref(sym, ex, nm, ccy);
        });
        ul.appendChild(b);
      }
      if (list.length) res.appendChild(ul);
      if (t212Has) {
        const wrap = document.createElement("div");
        wrap.className = "t212SearchBlock mt";
        const h = document.createElement("div");
        h.className = "h3";
        h.textContent = "Trading 212 (broker symbol list)";
        wrap.appendChild(h);
        if (t212j.hint) {
          const ph = document.createElement("p");
          ph.className = "sml muted";
          ph.textContent = String(t212j.hint);
          wrap.appendChild(ph);
        }
        const meta = document.createElement("p");
        meta.className = "sml muted";
        meta.textContent = `Cached ${Number(t212j.n_cached || 0)} instruments · complete: ${t212j.cache_complete ? "yes" : "no"}${t212j.cache_loading ? " · still loading…" : ""}`;
        wrap.appendChild(meta);
        const t212ul = document.createElement("div");
        t212ul.className = "list";
        for (const row of t212j.matches.slice(0, 18)) {
          const sym = String(row.symbol || "").trim();
          const tkt = String(row.t212Ticker || "").trim();
          if (!sym) continue;
          const nm = String(row.name || "").trim();
          const ccy = String(row.currency || "").trim();
          const qt = String(row.quoteType || "").trim();
          const chip = qt ? ` <span class="qtChip" title="Type">${esc(qt)}</span>` : "";
          const b2 = document.createElement("button");
          b2.type = "button";
          b2.className = "hit t212Hit";
          b2.innerHTML = `<div class="hitTop"><strong>${esc(sym)}</strong>${chip} <span class="muted sml">T212</span></div>
            <div class="hitName sml">${esc(tkt || sym)}</div>
            <div class="hitMeta sml muted">${esc(nm || "—")} · ${esc(ccy || "—")}</div>`;
          b2.addEventListener("click", () => {
            location.hash = instrumentHref(sym, "T212", nm, ccy);
          });
          t212ul.appendChild(b2);
        }
        wrap.appendChild(t212ul);
        res.appendChild(wrap);
      }
      const nHit = list.length ? ul.querySelectorAll("button.hit").length : 0;
      if (!nHit && list.length > 0) {
        res.innerHTML = `<p class="muted">No results with a usable symbol (upstream returned ${list.length} row(s) in an unexpected shape).</p>`;
        status("0 results");
        return;
      }
      if (filtEl && list.length) {
        const uniq = (fn) =>
          [...new Set(list.map(fn).map((x) => String(x || "").trim()).filter(Boolean))].sort((a, b) =>
            a.localeCompare(b, undefined, { sensitivity: "base" })
          );
        const countries = uniq((r) => r.country);
        const exchs = uniq((r) => r.exchangeShortName || r.exchange);
        const ccys = uniq((r) => r.currency);
        const mkSel = (id, label, values) => {
          if (!values || values.length < 1) return null;
          const lab = document.createElement("label");
          lab.className = "filtLab";
          const sp = document.createElement("span");
          sp.textContent = label;
          const sel = document.createElement("select");
          sel.className = "in filtSel";
          sel.id = id;
          const o0 = document.createElement("option");
          o0.value = "";
          o0.textContent = "All";
          sel.appendChild(o0);
          for (const v of values) {
            const o = document.createElement("option");
            o.value = v;
            o.textContent = v;
            sel.appendChild(o);
          }
          lab.appendChild(sp);
          lab.appendChild(sel);
          return lab;
        };
        const row = document.createElement("div");
        row.className = "searchFiltRow";
        const cLab = mkSel("fCountry", "Country", countries);
        const eLab = mkSel("fExch", "Exchange", exchs);
        const yLab = mkSel("fCcy", "Currency", ccys);
        [cLab, eLab, yLab].forEach((x) => {
          if (x) row.appendChild(x);
        });
        if (row.children.length) {
          filtEl.appendChild(row);
          filtEl.hidden = false;
          const applyFilt = () => {
            const gv = (id) => {
              const e = $(id);
              return e instanceof HTMLSelectElement ? e.value : "";
            };
            const c = gv("fCountry");
            const x = gv("fExch");
            const y = gv("fCcy");
            ul.querySelectorAll("button.hit").forEach((btn) => {
              const ok =
                (!c || btn.dataset.country === c) && (!x || btn.dataset.exch === x) && (!y || btn.dataset.ccy === y);
              btn.hidden = !ok;
            });
            const vis = ul.querySelectorAll("button.hit:not([hidden])").length;
            status(vis < nHit ? `${vis} shown (of ${nHit})` : `${nHit} results`);
          };
          row.querySelectorAll("select").forEach((s) => s.addEventListener("change", applyFilt));
        }
      }
      const nt = t212Has ? t212j.matches.length : 0;
      status(
        nt && nHit
          ? `${nHit} results + ${nt} T212`
          : nt
            ? `${nt} T212 (broker list)`
            : `${nHit} results`,
      );
      try {
        sessionStorage.setItem(SEARCH_LAST_Q, q);
      } catch {
        /* ignore */
      }
    } catch (e) {
      if (e?.name === "AbortError") return;
      res.innerHTML = `<p class="err">Could not reach the app server (network or server not running).</p>
        <p class="muted sml"><a href="/README.html#checklist-daily">Start checklist</a> ·
          <a href="/README.html#restart-server">Start / stop server</a> ·
          <a href="/README.html#python3-explained">What is python3?</a></p>`;
      status("Error");
    }
  }

  const debouncedSearch = debounce(() => void runSearchQuery(), 520);
  inp.addEventListener("input", debouncedSearch);
  try {
    const last = sessionStorage.getItem(SEARCH_LAST_Q);
    if (last && !inp.value.trim()) {
      inp.value = last;
      void runSearchQuery();
    }
  } catch {
    /* ignore */
  }
  const askBtn = $("aiAskBtn");
  if (askBtn instanceof HTMLButtonElement) askBtn.onclick = () => void runAiAsk();
  status("Search", document.documentElement.dataset.theme || "");
}

function setInstrumentTab(which) {
  const map = [
    ["ov", $("tabInstrOv"), $("tabBtnOv")],
    ["tech", $("tabInstrTech"), $("tabBtnTech")],
    ["fun", $("tabInstrFun"), $("tabBtnFun")],
  ];
  for (const [key, panel, btn] of map) {
    if (!(panel instanceof HTMLElement) || !(btn instanceof HTMLButtonElement)) continue;
    const on = key === which;
    panel.hidden = !on;
    btn.setAttribute("aria-selected", on ? "true" : "false");
  }
  if (which === "tech" && _instrQuoteCtx?.sym) {
    void ensureHistoryChartLoaded(_instrQuoteCtx.sym, _instrQuoteCtx.ex);
  }
}

function wireInstrument(sym, ex, nm, ccyHint, fromPf) {
  _instrQuoteCtx = { sym, ex, nm: nm || "", ccyHint: ccyHint || "", q: null };
  setInstrHistoryContext([], null);
  _histChartBars = null;
  _histLoadPromise = null;
  const st0 = $("chartSt");
  if (st0) st0.textContent = "Expand Price history (above) or open the Technicals tab to load the chart and indicators.";
  const chartD0 = $("instrChartDetails");
  if (chartD0 instanceof HTMLDetailsElement) chartD0.open = false;
  const holdMount = $("instrPfHold");
  if (holdMount instanceof HTMLElement) {
    if (fromPf) paintInstrPfHold(sym, ex);
    else {
      holdMount.hidden = true;
      holdMount.innerHTML = "";
    }
  }
  const ccyInp = $("addPfCcy");
  if (ccyInp instanceof HTMLInputElement && ccyHint && !ccyInp.value.trim()) {
    ccyInp.value = String(ccyHint).toUpperCase();
  }
  const brSel = $("addPfBroker");
  if (brSel instanceof HTMLSelectElement) {
    const b = getActiveBroker();
    brSel.value = isPfBrokerId(b) ? b : PF_T212;
  }

  document.querySelectorAll(".tabBar .tabBtn").forEach((btn) => {
    if (!(btn instanceof HTMLButtonElement)) return;
    btn.addEventListener("click", () => {
      const t = btn.dataset.tab;
      if (t === "ov" || t === "tech" || t === "fun") setInstrumentTab(t);
    });
  });

  $("addPfBtn")?.addEventListener("click", () => {
    const msg = $("addPfMsg");
    const qtyE = $("addPfQty");
    const avgE = $("addPfAvg");
    const ccyE = $("addPfCcy");
    const qty = qtyE instanceof HTMLInputElement ? num(qtyE.value) : 0;
    const avg = avgE instanceof HTMLInputElement ? num(avgE.value) : 0;
    let ccy = ccyE instanceof HTMLInputElement ? ccyE.value.trim().toUpperCase() : "";
    if (!ccy) ccy = String(_instrQuoteCtx.q?.currency || _instrQuoteCtx.ccyHint || "").toUpperCase();
    ccy = normalizeCcyForFx(ccy);
    if (!sym || !ccy || qty <= 0) {
      if (msg) msg.textContent = "Enter qty (> 0) and currency.";
      return;
    }
    const last = num(_instrQuoteCtx.q?.price);
    const brPick = $("addPfBroker");
    const br = brPick instanceof HTMLSelectElement && isPfBrokerId(brPick.value) ? brPick.value : PF_T212;
    const bundle = loadPfBundle();
    bundle.brokers[br].rows.push({
      sym,
      ex: ex || "",
      ccy,
      qty,
      avg,
      last,
      nm: nm || String(_instrQuoteCtx.q?.name || ""),
    });
    savePfBundle(bundle);
    if (msg) msg.textContent = `Row added to ${PF_BROKER_LABEL[br]}. Open Portfolio to refresh prices.`;
    status(`Added → ${PF_BROKER_LABEL[br]}`);
  });

  $("addWlBtn")?.addEventListener("click", () => {
    const msg = $("addWlMsg");
    let ccy = String(_instrQuoteCtx.ccyHint || "").trim().toUpperCase();
    if (!ccy) ccy = String(_instrQuoteCtx.q?.currency || "").trim().toUpperCase();
    const row = { sym, ex: ex || "", nm: nm || String(_instrQuoteCtx.q?.name || ""), ccy };
    if (!addWlRow(row)) {
      if (msg) msg.textContent = "Already on your watchlist.";
      status("Watchlist — duplicate");
      return;
    }
    if (msg) msg.textContent = "Saved. Open Watchlist in the top nav to manage symbols.";
    status("Added → watchlist");
  });

  const mount = $("instrQuoteMount");
  if (mount) fillInstrQuoteMount(mount, sym, ex);
  syncHistChartTypeButtons();
  const chartDetails = $("instrChartDetails");
  if (chartDetails instanceof HTMLDetailsElement) {
    chartDetails.addEventListener("toggle", () => {
      if (!chartDetails.open) return;
      if (!_instrQuoteCtx?.sym) return;
      void ensureHistoryChartLoaded(_instrQuoteCtx.sym, _instrQuoteCtx.ex).then(() => {
        requestAnimationFrame(() => {
          redrawHistCanvas();
        });
      });
    });
  }
  $("instrChartDock")?.querySelectorAll("button[data-chart-type]").forEach((btn) => {
    if (!(btn instanceof HTMLButtonElement)) return;
    btn.addEventListener("click", () => {
      const t = btn.dataset.chartType;
      if (t !== "area" && t !== "candle" && t !== "heikin") return;
      _histChartType = t;
      syncHistChartTypeButtons();
      redrawHistCanvas();
      updateChartStatusLine();
    });
  });
  const ol = loadChartOverlayPrefs();
  for (const [id, key] of [
    ["olMa20", "ma20"],
    ["olMa50", "ma50"],
    ["olBb", "bb"],
    ["olVol", "vol"],
    ["olTrend", "trend"],
  ]) {
    const el = $(id);
    if (el instanceof HTMLInputElement) {
      el.checked = !!ol[key];
      el.addEventListener("change", () => {
        saveChartOverlayPrefs(getChartOverlayState());
        redrawHistCanvas();
      });
    }
  }
  void loadInstrumentNewsCorp(sym, ex);
  setInstrumentTab("ov");
  status(sym, document.documentElement.dataset.theme || "");
}

/** True when server returned 404 unknown route for instrument extras (stale `server.py`). */
function isStaleInstrumentExtrasError(detail, errKey) {
  const d = String(detail || "");
  const e = String(errKey || "");
  if (e === "unknown route") return true;
  if (d === "/api/news" || d === "/api/corporate") return true;
  return /unknown route/i.test(d + e);
}

function staleServerExtrasHintHtml() {
  return `<p class="muted sml">This usually means an <strong>old</strong> <code>server.py</code> is still running. In Terminal: <code>cd</code> to this app’s folder, press <strong>Ctrl+C</strong>, then run <code>python3 server.py</code> again. Open <a class="inlineHealth" href="/api/health" target="_blank" rel="noopener">/api/health</a> — <code>api_revision</code> should be <strong>10</strong> or higher for news, corporate, and EUR FX routes.</p>`;
}

/** Render `/api/news` JSON into overview HTML (headlines + disclaimer). */
function renderInstrNewsPayload(j) {
  if (!j || (j.error && !Array.isArray(j.items))) {
    const raw = String(j?.detail || j?.error || "Request failed");
    const stale = isStaleInstrumentExtrasError(raw, j?.error);
    const det = esc(raw);
    const h = j?.hint ? `<p class="muted sml">${esc(j.hint)}</p>` : "";
    const sh = stale ? staleServerExtrasHintHtml() : "";
    return `<p class="err">${det}</p>${sh}${h}`;
  }
  const items = Array.isArray(j.items) ? j.items : [];
  if (!items.length) {
    const h = j.hint ? `<p class="muted sml">${esc(j.hint)}</p>` : `<p class="muted sml">No articles.</p>`;
    return h;
  }
  const src =
    Array.isArray(j.sources) && j.sources.length
      ? `<p class="newsSrc sml muted">Sources: ${esc(j.sources.join(" · "))}</p>`
      : "";
  const disc = j.disclaimer ? `<p class="sml muted mt">${esc(j.disclaimer)}</p>` : "";
  const lis = items
    .map((it) => {
      const u = String(it.url || "").trim();
      const t = esc(it.title || "");
      const pub = it.published
        ? `<span class="newsMeta">${esc(it.published)}${it.source ? ` · ${esc(it.source)}` : ""}</span>`
        : it.source
          ? `<span class="newsMeta">${esc(it.source)}</span>`
          : "";
      const inner = u ? `<a href="${esc(u)}" target="_blank" rel="noopener noreferrer">${t}</a>` : t;
      return `<li>${inner}${pub ? `<br />${pub}` : ""}</li>`;
    })
    .join("");
  return `${src}<ul class="newsList" role="list">${lis}</ul>${disc}`;
}

/** Render `/api/corporate` JSON (earnings, dividends, splits). */
function renderInstrCorpPayload(j) {
  if (!j || (j.error && !Array.isArray(j.dividends) && !Array.isArray(j.splits) && !Array.isArray(j.earnings))) {
    const raw = String(j?.detail || j?.error || "Request failed");
    const stale = isStaleInstrumentExtrasError(raw, j?.error);
    const det = esc(raw);
    const h = j?.hint ? `<p class="muted sml">${esc(j.hint)}</p>` : "";
    const sh = stale ? staleServerExtrasHintHtml() : "";
    return `<p class="err">${det}</p>${sh}${h}`;
  }
  const earn = Array.isArray(j.earnings) ? j.earnings : [];
  const divs = Array.isArray(j.dividends) ? j.dividends : [];
  const spls = Array.isArray(j.splits) ? j.splits : [];
  const src =
    Array.isArray(j.sources) && j.sources.length
      ? `<p class="corpSrc sml muted">Sources: ${esc(j.sources.join(" · "))}</p>`
      : "";
  const disc = j.disclaimer ? `<p class="sml muted mt">${esc(j.disclaimer)}</p>` : "";
  const blocks = [];
  if (earn.length) {
    const rows = earn
      .map(
        (e) => `<tr>
        <td>${esc(e.date || "—")}</td>
        <td class="sml">${e.epsEstimated != null && e.epsEstimated !== "" ? esc(fmtN(Number(e.epsEstimated), 2)) : "—"}</td>
        <td class="sml">${e.epsActual != null && e.epsActual !== "" ? esc(fmtN(Number(e.epsActual), 2)) : "—"}</td>
        <td class="sml">${esc(e.time || "")}</td>
      </tr>`,
      )
      .join("");
    blocks.push(`<div class="corpBlock"><div class="h4 corpH">Earnings history</div>
      <table class="corpTbl"><thead><tr><th>Date</th><th>EPS est.</th><th>EPS actual</th><th>Time</th></tr></thead><tbody>${rows}</tbody></table></div>`);
  }
  if (divs.length) {
    const rows = divs
      .map(
        (d) => `<tr>
        <td>${esc(d.date || "—")}</td>
        <td class="sml">${d.amount != null && d.amount !== "" ? esc(fmtN(Number(d.amount), 4)) : "—"}</td>
        <td class="sml">${esc(d.currency || "")}</td>
        <td class="sml">${esc([d.recordDate, d.paymentDate].filter(Boolean).join(" / "))}</td>
      </tr>`,
      )
      .join("");
    blocks.push(`<div class="corpBlock"><div class="h4 corpH">Dividends</div>
      <table class="corpTbl"><thead><tr><th>Date</th><th>Amount</th><th>Ccy</th><th>Record / pay</th></tr></thead><tbody>${rows}</tbody></table></div>`);
  }
  if (spls.length) {
    const rows = spls.map((s) => `<tr><td>${esc(s.date || "—")}</td><td>${esc(s.ratio || "—")}</td></tr>`).join("");
    blocks.push(`<div class="corpBlock"><div class="h4 corpH">Stock splits</div>
      <table class="corpTbl"><thead><tr><th>Date</th><th>Ratio</th></tr></thead><tbody>${rows}</tbody></table></div>`);
  }
  if (!blocks.length) {
    const h = j.hint ? `<p class="muted sml">${esc(j.hint)}</p>` : `<p class="muted sml">No rows returned.</p>`;
    return `${src}${h}${disc}`;
  }
  return `${src}<div class="corpGrid">${blocks.join("")}</div>${disc}`;
}

async function loadInstrumentNewsCorp(sym, ex) {
  const newsEl = $("instrNews");
  const corpEl = $("instrCorp");
  if (!newsEl || !corpEl) return;
  newsEl.innerHTML = `<p class="muted sml">Loading headlines…</p>`;
  corpEl.innerHTML = `<p class="muted sml">Loading calendar…</p>`;
  const u = new URLSearchParams({ symbol: sym });
  if (ex) u.set("exchange", ex);
  try {
    const [rn, rc] = await Promise.all([fetch(`/api/news?${u}`), fetch(`/api/corporate?${u}`)]);
    const txtN = await rn.text();
    const txtC = await rc.text();
    let jn = null;
    let jc = null;
    try {
      jn = JSON.parse(txtN);
    } catch {
      jn = { error: "bad response", detail: txtN.slice(0, 240) };
    }
    try {
      jc = JSON.parse(txtC);
    } catch {
      jc = { error: "bad response", detail: txtC.slice(0, 240) };
    }
    newsEl.innerHTML = rn.ok
      ? renderInstrNewsPayload(jn)
      : renderInstrNewsPayload({ error: true, detail: jn?.detail || jn?.error || String(rn.status), hint: jn?.hint });
    corpEl.innerHTML = rc.ok
      ? renderInstrCorpPayload(jc)
      : renderInstrCorpPayload({ error: true, detail: jc?.detail || jc?.error || String(rc.status), hint: jc?.hint });
  } catch {
    newsEl.innerHTML = `<p class="err">Network error loading news.</p>`;
    corpEl.innerHTML = `<p class="err">Network error loading corporate data.</p>`;
  }
}

function fmtVolShort(n) {
  if (!Number.isFinite(n) || n <= 0) return "—";
  if (n >= 1e9) return `${fmtN(n / 1e9, 2)}B`;
  if (n >= 1e6) return `${fmtN(n / 1e6, 2)}M`;
  if (n >= 1e3) return `${fmtN(n / 1e3, 0)}k`;
  return String(Math.round(n));
}

/** Large integers (e.g. market cap) in compact form. */
function fmtBigNumber(n) {
  if (!Number.isFinite(n) || n <= 0) return "—";
  if (n >= 1e12) return `${fmtN(n / 1e12, 2)}T`;
  if (n >= 1e9) return `${fmtN(n / 1e9, 2)}B`;
  if (n >= 1e6) return `${fmtN(n / 1e6, 2)}M`;
  if (n >= 1e3) return `${fmtN(n / 1e3, 1)}k`;
  return fmtN(n, 0);
}

function renderFundamentalsMount(sym, ex, q) {
  const el = $("funMount");
  if (!el) return;
  if (!q) {
    el.innerHTML = `<p class="muted sml">Fundamentals appear after a successful quote on <strong>Overview</strong>.</p>`;
    return;
  }
  el.innerHTML = fundamentalsHtml(sym, ex, q);
  wireFundamentalsAi(sym, ex, q);
}

/** Optional block when this symbol exists in the saved portfolio (same device / URL). */
function fundamentalsPositionBlock(sym, ex) {
  const hits = findPfHoldingMatches(sym, ex);
  if (!hits.length) return "";
  const rows = hits
    .map(({ brokerId, r }) => {
      const qty = num(r.qty);
      const last = num(r.last);
      const avg = num(r.avg);
      const ccyR = String(r.ccy || "").toUpperCase() || "—";
      const val = qty * last;
      const pl = val - qty * avg;
      const plCls = pl >= 0 ? "plp" : "pln";
      return `<div class="funRow"><span class="muted">${esc(PF_BROKER_LABEL[brokerId])}</span><strong>${esc(fmtN(qty, 4))} × last → ${esc(fmtMoney(ccyR, val))}</strong> <span class="${plCls}">P/L ${esc(fmtMoney(ccyR, pl))}</span></div>`;
    })
    .join("");
  return `<div class="card2 mt">
    <div class="h3">Your position (saved here)</div>
    <p class="sml muted">From Portfolio on <strong>this</strong> browser address only.</p>
    <div class="funGrid mt">${rows}</div>
  </div>`;
}

function fundamentalsHtml(sym, ex, q) {
  const ccy = (q.currency || "").toUpperCase();
  const vol = Number(q.volume);
  const volS = fmtVolShort(vol);
  const exs = esc(String(ex || q.exchange || "—"));
  const priceN = Number(q.price);
  const prevN = Number(q.previousClose);
  const openN = Number(q.open);
  const dhN = Number(q.dayHigh);
  const dlN = Number(q.dayLow);
  const yhN = Number(q.yearHigh);
  const ylN = Number(q.yearLow);
  const mcapN = Number(q.marketCap ?? q.market_cap);
  const snap = [];
  const pushMoney = (lab, n) => {
    if (Number.isFinite(n)) snap.push(`<div class="funRow"><span class="muted">${esc(lab)}</span><strong>${esc(fmtMoney(ccy, n))}</strong></div>`);
  };
  if (Number.isFinite(mcapN) && mcapN > 0) {
    snap.push(
      `<div class="funRow"><span class="muted">Market cap (≈ company size)</span><strong>${esc(fmtBigNumber(mcapN))}</strong> <span class="sml muted">(≈ ${esc(fmtMoney(ccy, mcapN))} notional)</span></div>`,
    );
  }
  pushMoney("Previous close", prevN);
  pushMoney("Session open", openN);
  pushMoney("Day high", dhN);
  pushMoney("Day low", dlN);
  pushMoney("52-week high", yhN);
  pushMoney("52-week low", ylN);
  const extras = [];
  for (const [k, lab, kind] of [
    ["trailingPE", "Trailing P/E", "num"],
    ["forwardPE", "Forward P/E", "num"],
    ["peRatio", "P/E", "num"],
    ["priceToBook", "Price / book", "num"],
    ["beta", "Beta (β)", "num"],
    ["dividendYield", "Dividend yield", "pct"],
    ["dividendRate", "Dividend rate (per share)", "num"],
    ["epsTrailing", "EPS (trailing)", "num"],
    ["epsForward", "EPS (forward)", "num"],
    ["bookValue", "Book value / share", "num"],
    ["avgPrice50d", "50-day avg price", "money"],
    ["avgPrice200d", "200-day avg price", "money"],
    ["changePercent", "Session change %", "pct"],
    ["changeAmount", "Session change", "money"],
    ["averageVolume", "Avg volume (quote)", "vol"],
    ["avgVolume3Mo", "Avg volume (3 mo)", "vol"],
    ["avgVolume10d", "Avg volume (10 d)", "vol"],
    ["marketOpen", "Market open (flag)", "raw"],
  ]) {
    const v = q[k];
    if (v == null || v === "") continue;
    if (kind === "raw") {
      extras.push(`<div class="funRow"><span class="muted">${esc(lab)}</span><strong>${esc(String(v))}</strong></div>`);
      continue;
    }
    const n = Number(v);
    if (kind === "big" && Number.isFinite(n)) {
      extras.push(
        `<div class="funRow"><span class="muted">${esc(lab)}</span><strong>${esc(fmtBigNumber(n))}</strong> <span class="sml muted">(raw ${esc(fmtN(n, 0))})</span></div>`,
      );
    } else if (kind === "pct" && Number.isFinite(n)) {
      const disp = Math.abs(n) <= 1 && n !== 0 ? n * 100 : n;
      extras.push(`<div class="funRow"><span class="muted">${esc(lab)}</span><strong>${esc(fmtN(disp, 2))}%</strong></div>`);
    } else if (kind === "vol" && Number.isFinite(n)) {
      extras.push(`<div class="funRow"><span class="muted">${esc(lab)}</span><strong>${esc(fmtVolShort(n))}</strong></div>`);
    } else if (kind === "money" && Number.isFinite(n)) {
      extras.push(`<div class="funRow"><span class="muted">${esc(lab)}</span><strong>${esc(fmtMoney(ccy, n))}</strong></div>`);
    } else {
      extras.push(
        `<div class="funRow"><span class="muted">${esc(lab)}</span><strong>${Number.isFinite(n) ? esc(fmtN(n, 4)) : esc(String(v))}</strong></div>`,
      );
    }
  }
  const xSym = esc(String(sym || q.symbol || ""));
  const xNm = esc(String(q.name || ""));
  const snapBlock =
    snap.length > 0
      ? `<div class="funExtras mt funSnap">${snap.join("")}</div>`
      : `<p class="sml muted mt">No day range / 52w / market cap in this quote — Yahoo and other sources differ; try again later or another listing.</p>`;
  return `<div class="card2">
    <div class="h3">Fundamentals snapshot</div>
    <p class="sml muted">Same live <code>/api/quote</code> data as Overview. Extra lines depend on the quote provider.</p>
    <div class="funGrid mt">
      <div class="funRow"><span class="muted">Symbol</span><strong>${xSym}</strong></div>
      <div class="funRow"><span class="muted">Name</span><strong>${xNm || "—"}</strong></div>
      <div class="funRow"><span class="muted">Exchange</span><strong>${exs}</strong></div>
      <div class="funRow"><span class="muted">Currency</span><strong>${esc(ccy || "—")}</strong></div>
      <div class="funRow"><span class="muted">Last price</span><strong>${Number.isFinite(priceN) ? esc(fmtMoney(ccy, priceN)) : "—"}</strong></div>
      <div class="funRow"><span class="muted">Volume (session)</span><strong>${esc(volS)}</strong></div>
      <div class="funRow"><span class="muted">Quote type</span><strong>${esc(String(q.quoteType || "—"))}</strong></div>
    </div>
    ${snapBlock}
    ${extras.length ? `<div class="funExtras mt">${extras.join("")}</div>` : ""}
    ${extras.length === 0 && snap.length > 0 ? `<p class="sml muted mt">No P/E, dividend, or beta fields in this quote payload.</p>` : ""}
  </div>
  ${fundamentalsPositionBlock(sym, ex)}
  <div class="card2 mt" id="aiCommentaryCard">
    <div class="h3">AI commentary <span class="sml muted">(server API)</span></div>
    <p class="sml muted"><strong>To turn this on (you do this on the Mac, once):</strong> open the project’s <code>.env</code> in a text editor and add a line like <code>OPENAI_API_KEY=sk-…</code> (or <code>ANTHROPIC_API_KEY</code>, <code>GOOGLE_AI_API_KEY</code> / <code>GEMINI_API_KEY</code>, <code>XAI_API_KEY</code>). Save, then restart <code>JohnsStockApp.command</code>. Open <a href="/api/health" target="_blank" rel="noopener"><code>/api/health</code></a> — under <code>llm_commentary.configured</code> one entry should become <code>true</code>. Keys stay on the server; you never paste them here in chat.</p>
    <p class="sml muted">Uses <code>GET /api/ai-commentary</code>. Optional: <code>AI_COMMENTARY_PROVIDER=openai|anthropic|google|xai|auto</code>.</p>
    <div class="aiComRow mt">
      <button type="button" class="btn" id="aiCommentaryRun">Generate commentary</button>
    </div>
    <div id="aiCommentaryOut" class="aiCommentaryOut mt" aria-live="polite"></div>
    <p class="sml muted mt">Open the <strong>Technicals</strong> tab first if you want RS55 / RSI / mood numbers included; otherwise the model only sees the quote snapshot.</p>
  </div>`;
}

function wireFundamentalsAi(sym, ex, q) {
  const run = $("aiCommentaryRun");
  const out = $("aiCommentaryOut");
  if (!(run instanceof HTMLButtonElement) || !(out instanceof HTMLElement)) return;
  run.onclick = () => void runAiCommentary(sym, ex, q, out, run);
}

async function runAiCommentary(sym, ex, _quote, outEl, btn) {
  btn.disabled = true;
  outEl.innerHTML = `<p class="muted sml">Calling model… (often 10–45s)</p>`;
  const technical_summary = technicalSummaryPlainFromContext().slice(0, 4500);
  const params = new URLSearchParams();
  params.set("symbol", sym);
  params.set("exchange", ex || "");
  if (technical_summary) params.set("technical_summary", technical_summary);
  try {
    const r = await fetch(`/api/ai-commentary?${params.toString()}`, { cache: "no-store" });
    const raw = await r.text();
    let j;
    try {
      j = JSON.parse(raw);
    } catch {
      outEl.innerHTML = `<p class="err">Server returned non-JSON.</p><pre class="sml">${esc(raw.slice(0, 800))}</pre>`;
      return;
    }
    if (!r.ok) {
      outEl.innerHTML = formatApiErrPayload(j, r);
      return;
    }
    const text = j.text != null ? String(j.text) : "";
    const metaBits = [j.provider, j.model].filter(Boolean).map((x) => String(x));
    const meta = metaBits.length ? `<p class="sml muted">${esc(metaBits.join(" · "))}</p>` : "";
    outEl.innerHTML = `${meta}<pre class="aiCommentaryPre">${esc(text)}</pre>`;
  } catch {
    outEl.innerHTML = `<p class="err">Network error calling <code>/api/ai-commentary</code>.</p>`;
  } finally {
    btn.disabled = false;
  }
}

async function runAiAsk() {
  const ta = $("aiAskQ");
  const out = $("aiAskOut");
  const btn = $("aiAskBtn");
  if (!(ta instanceof HTMLTextAreaElement) || !(out instanceof HTMLElement) || !(btn instanceof HTMLButtonElement)) {
    return;
  }
  const q = ta.value.trim();
  if (!q) {
    status("Type a question first");
    return;
  }
  btn.disabled = true;
  out.innerHTML = `<p class="muted sml">Thinking… (often 10–45s)</p>`;
  try {
    const params = new URLSearchParams({ q: q.slice(0, 3000) });
    const r = await fetch(`/api/ai-ask?${params.toString()}`, { cache: "no-store" });
    const raw = await r.text();
    let j;
    try {
      j = JSON.parse(raw);
    } catch {
      out.innerHTML = `<p class="err">Server returned non-JSON.</p><pre class="sml">${esc(raw.slice(0, 800))}</pre>`;
      return;
    }
    if (!r.ok) {
      out.innerHTML = formatApiErrPayload(j, r);
      return;
    }
    const text = j.text != null ? String(j.text) : "";
    const metaBits = [j.provider, j.model].filter(Boolean).map((x) => String(x));
    const meta = metaBits.length ? `<p class="sml muted">${esc(metaBits.join(" · "))}</p>` : "";
    out.innerHTML = `${meta}<pre class="aiCommentaryPre">${esc(text)}</pre>`;
    status("AI answer ready");
  } catch {
    out.innerHTML = `<p class="err">Network error calling <code>/api/ai-ask</code>.</p>`;
    status("AI ask failed");
  } finally {
    btn.disabled = false;
  }
}

async function fillInstrQuoteMount(el, sym, ex) {
  el.innerHTML = `<p class="muted">Loading…</p>`;
  status(`Quote ${sym}…`);
  const u = new URLSearchParams({ symbol: sym });
  if (ex) u.set("exchange", ex);
  try {
    const r = await fetch(`/api/quote?${u}`);
    const raw = await r.text();
    let j;
    try {
      j = JSON.parse(raw);
    } catch {
      el.innerHTML = `<p class="err">Bad response from server (not JSON).</p>
        <p class="muted sml"><a href="/README.html#checklist-daily">Start checklist</a> · <a href="/README.html#python3-explained">What is python3?</a></p>
        <pre class="sml">${esc(raw.slice(0, 400))}</pre>`;
      status("Quote failed");
      _instrQuoteCtx.q = null;
      renderFundamentalsMount(sym, ex, null);
      return;
    }
    if (!r.ok) {
      const hintHtml = j.hint ? `<p class="muted sml">${esc(j.hint)}</p>` : "";
      el.innerHTML = `<p class="err">${esc(j.detail || j.error)}</p>${hintHtml}<p class="muted sml">
        <a class="inlineHealth" href="/api/health" target="_blank" rel="noopener">Open API health (new tab)</a> ·
        <a href="/README.html#troubleshooting">Troubleshooting</a></p>`;
      status("Quote failed");
      _instrQuoteCtx.q = null;
      renderFundamentalsMount(sym, ex, null);
      return;
    }
    const q = Array.isArray(j) ? j[0] : null;
    if (!q) {
      el.innerHTML = `<p class="err">Quote response had no row.</p>`;
      status("Quote failed");
      _instrQuoteCtx.q = null;
      renderFundamentalsMount(sym, ex, null);
      return;
    }
    _instrQuoteCtx.q = q;
    const ccyInp = $("addPfCcy");
    if (ccyInp instanceof HTMLInputElement && q?.currency && !ccyInp.value.trim()) {
      ccyInp.value = String(q.currency).toUpperCase();
    }
    el.innerHTML = quoteCard(q);
    renderFundamentalsMount(sym, ex, q);
    status("OK");
  } catch {
    el.innerHTML = `<p class="err">Network error.</p><p class="muted sml"><a href="/README.html#troubleshooting">Troubleshooting</a></p>`;
    status("Quote failed");
    _instrQuoteCtx.q = null;
    renderFundamentalsMount(sym, ex, null);
  }
}

/**
 * Load daily history once per symbol (shared by the Price history panel and Technicals tab).
 * If a fetch is already in flight, await the same promise.
 */
function ensureHistoryChartLoaded(sym, ex) {
  if (!sym) return Promise.resolve();
  const want = `${String(sym).trim()}\t${String(ex || "").trim()}`;
  if (_histChartBars && _histChartBars.length >= 2) {
    const have = _instrQuoteCtx
      ? `${String(_instrQuoteCtx.sym).trim()}\t${String(_instrQuoteCtx.ex || "").trim()}`
      : "";
    if (have && have === want) return Promise.resolve();
  }
  if (_histLoadPromise) return _histLoadPromise;
  const st = $("chartSt");
  if (st && (!_histChartBars || _histChartBars.length < 2)) st.textContent = "Loading chart data…";
  _histLoadPromise = (async () => {
    try {
      await loadHistoryChart(sym, ex);
    } finally {
      _histLoadPromise = null;
    }
  })();
  return _histLoadPromise;
}

/** Fetches OHLC history (+ benchmark for RS55) and draws the Technicals chart. */
async function loadHistoryChart(sym, ex) {
  const st = $("chartSt");
  const cv = $("histCanvas");
  const mount = $("techMount");
  if (!st || !(cv instanceof HTMLCanvasElement)) return;
  const requestedRange = "1y";
  _histChartStale = false;
  setChartStaleBanner(false);
  const u = new URLSearchParams({ symbol: sym, range: requestedRange });
  if (ex) u.set("exchange", ex);
  const benchSym = benchSymbolForExchange(ex);
  const benchPrms = benchSym ? new URLSearchParams({ symbol: benchSym, range: requestedRange }) : null;
  try {
    const stockRes = await fetch(`/api/history?${u}`);
    const raw = await stockRes.text();
    let j;
    try {
      j = JSON.parse(raw);
    } catch {
      if (tryApplyHistSnap(sym, ex, requestedRange, cv, mount)) return;
      _histChartBars = null;
      st.textContent = isHostedWebApp()
        ? "Chart: server returned a non-JSON response (host waking up or a short network error). Wait ~30s and refresh, or retry the chart tab."
        : "Chart: server returned non-JSON.";
      return;
    }
    if (!stockRes.ok) {
      const det = String(j.detail || j.error || stockRes.status);
      const unk = stockRes.status === 404 && /unknown route/i.test(String(j.error || ""));
      const extra = unk
        ? isHostedWebApp()
          ? " Redeploy the latest app on your host, or check /api/health in a new tab."
          : ` Stop the old server on this port (Terminal: Ctrl+C, or run: lsof -iTCP:${location.port} -sTCP:LISTEN then kill that PID), then start python3 server.py again. /api/health should show api_revision: 10.`
        : "";
      const h = j.hint ? ` ${j.hint}` : "";
      if (tryApplyHistSnap(sym, ex, requestedRange, cv, mount)) return;
      _histChartBars = null;
      st.textContent = `Chart unavailable: ${det}.${extra}${h}`;
      return;
    }
    const bars = Array.isArray(j.bars) ? j.bars : [];
    if (bars.length < 2) {
      if (tryApplyHistSnap(sym, ex, requestedRange, cv, mount)) return;
      _histChartBars = null;
      st.textContent = "Chart: not enough daily bars.";
      return;
    }
    let rsExtra = null;
    if (benchPrms) {
      try {
        await new Promise((r) => setTimeout(r, 450));
        const benchRes = await fetch(`/api/history?${benchPrms}`);
        const rawB = await benchRes.text();
        let jb = null;
        if (benchRes.ok) {
          try {
            jb = JSON.parse(rawB);
          } catch {
            jb = null;
          }
        }
        const bBars = Array.isArray(jb?.bars) ? jb.bars : [];
        const aligned = alignStockBenchByT(bars, bBars);
        const r55 = rs55BajajRatio(aligned);
        if (r55 != null) {
          rsExtra = {
            rs55: r55,
            benchLabel: benchHumanLabel(benchSym),
            benchSym,
          };
        }
      } catch {
        /* RS55 optional */
      }
    }
    const snapRng = j.range || requestedRange;
    _histChartBars = bars;
    _histChartMeta = { source: j.source || "", range: snapRng };
    _histChartStale = false;
    setChartStaleBanner(false);
    writeHistSnap(sym, ex, snapRng, {
      bars,
      source: _histChartMeta.source,
      range: snapRng,
      rsExtra,
    });
    drawHistoryChart(cv, bars, _histChartMeta.range, _histChartType);
    syncHistChartTypeButtons();
    updateChartStatusLine();
    if (mount) mount.innerHTML = technicalsHtmlFromBars(bars, rsExtra);
    setInstrHistoryContext(bars, rsExtra);
  } catch {
    if (tryApplyHistSnap(sym, ex, requestedRange, cv, mount)) return;
    _histChartBars = null;
    st.textContent = "Chart: network error.";
  }
}

/** Simple moving average of the last `period` closes (series oldest → newest). */
function smaLast(closes, period) {
  if (!Array.isArray(closes) || closes.length < period) return null;
  let s = 0;
  const start = closes.length - period;
  for (let i = start; i < closes.length; i++) s += closes[i];
  return s / period;
}

/** Full SMA series; entries before `period-1` are null. */
function smaSeries(closes, period) {
  const n = closes.length;
  const out = new Array(n).fill(null);
  if (n < period) return out;
  let s = 0;
  for (let i = 0; i < period; i++) s += closes[i];
  out[period - 1] = s / period;
  for (let i = period; i < n; i++) {
    s += closes[i] - closes[i - period];
    out[i] = s / period;
  }
  return out;
}

function bollingerSeries(closes, period, k) {
  const n = closes.length;
  const mid = smaSeries(closes, period);
  const upper = new Array(n).fill(null);
  const lower = new Array(n).fill(null);
  for (let i = period - 1; i < n; i++) {
    const m = mid[i];
    if (m == null) continue;
    let v = 0;
    for (let j = i - period + 1; j <= i; j++) {
      const d = closes[j] - m;
      v += d * d;
    }
    const sd = Math.sqrt(v / period);
    upper[i] = m + k * sd;
    lower[i] = m - k * sd;
  }
  return { mid, upper, lower };
}

function extendLoHi(lo, hi, seriesList) {
  let a = lo;
  let b = hi;
  for (const arr of seriesList) {
    if (!arr) continue;
    for (const v of arr) {
      if (v != null && Number.isFinite(v)) {
        a = Math.min(a, v);
        b = Math.max(b, v);
      }
    }
  }
  return { lo: a, hi: b };
}

function strokeLineSeries(ctx, series, xAt, yAt, n) {
  let pen = false;
  for (let i = 0; i <= n; i++) {
    const v = i < n ? series[i] : null;
    if (v == null || !Number.isFinite(v)) {
      if (pen) {
        ctx.stroke();
        pen = false;
      }
      continue;
    }
    const x = xAt(i);
    const y = yAt(v);
    if (!pen) {
      ctx.beginPath();
      ctx.moveTo(x, y);
      pen = true;
    } else {
      ctx.lineTo(x, y);
    }
  }
}

function fillBollingerBand(ctx, upper, lower, xAt, yAt, n, isLight) {
  let i0 = -1;
  for (let i = 0; i < n; i++) {
    if (upper[i] != null && lower[i] != null && Number.isFinite(upper[i]) && Number.isFinite(lower[i])) {
      i0 = i;
      break;
    }
  }
  if (i0 < 0) return;
  ctx.beginPath();
  ctx.moveTo(xAt(i0), yAt(upper[i0]));
  let lastU = i0;
  for (let i = i0 + 1; i < n; i++) {
    if (upper[i] == null || !Number.isFinite(upper[i])) break;
    ctx.lineTo(xAt(i), yAt(upper[i]));
    lastU = i;
  }
  for (let i = lastU; i >= i0; i--) {
    if (lower[i] == null || !Number.isFinite(lower[i])) break;
    ctx.lineTo(xAt(i), yAt(lower[i]));
  }
  ctx.closePath();
  ctx.fillStyle = isLight ? "rgba(99,102,241,0.14)" : "rgba(165,180,252,0.16)";
  ctx.fill();
  ctx.strokeStyle = isLight ? "rgba(79,70,229,0.4)" : "rgba(165,180,252,0.5)";
  ctx.lineWidth = 1;
  strokeLineSeries(ctx, upper, xAt, yAt, n);
  strokeLineSeries(ctx, lower, xAt, yAt, n);
}

function drawStudyOverlays(ctx, layout, closes, overlays) {
  const { padL, innerW, padT, innerH, lo, hi, n, isLight } = layout;
  const xAt = (i) => padL + (innerW * i) / (n - 1);
  const yAt = (v) => padT + innerH * (1 - (v - lo) / (hi - lo));
  if (overlays.bb && n >= 20) {
    const { upper, lower } = bollingerSeries(closes, 20, 2);
    fillBollingerBand(ctx, upper, lower, xAt, yAt, n, isLight);
  }
  if (overlays.ma20 && n >= 20) {
    const s20 = smaSeries(closes, 20);
    ctx.strokeStyle = isLight ? "#b45309" : "#fbbf24";
    ctx.lineWidth = 1.25;
    strokeLineSeries(ctx, s20, xAt, yAt, n);
  }
  if (overlays.ma50 && n >= 50) {
    const s50 = smaSeries(closes, 50);
    ctx.strokeStyle = isLight ? "#6d28d9" : "#c4b5fd";
    ctx.lineWidth = 1.25;
    strokeLineSeries(ctx, s50, xAt, yAt, n);
  }
}

/** OLS line through close index i = 0..n-1 (visual trend aid, not a trade system). */
function computeLinearRegression(closes) {
  const n = closes.length;
  if (n < 5) return null;
  let sx = 0;
  let sy = 0;
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < n; i++) {
    sx += i;
    sy += closes[i];
    sxy += i * closes[i];
    sxx += i * i;
  }
  const denom = n * sxx - sx * sx;
  if (Math.abs(denom) < 1e-14) return null;
  const slope = (n * sxy - sx * sy) / denom;
  const intercept = (sy - slope * sx) / n;
  return { slope, intercept, n };
}

function drawTrendlineOverlay(ctx, layout, closes, overlays) {
  if (!overlays.trend || !Array.isArray(closes) || closes.length < 5) return;
  const { padL, innerW, padT, innerH, lo, hi, n, isLight } = layout;
  if (closes.length !== n) return;
  const reg = computeLinearRegression(closes);
  if (!reg) return;
  const xAt = (i) => padL + (innerW * i) / (n - 1);
  const yAt = (v) => padT + innerH * (1 - (v - lo) / (hi - lo));
  const y0 = reg.intercept;
  const y1 = reg.intercept + reg.slope * (n - 1);
  ctx.save();
  ctx.strokeStyle = isLight ? "rgba(37,99,235,0.92)" : "rgba(147,197,253,0.95)";
  ctx.lineWidth = 1.75;
  ctx.setLineDash([7, 5]);
  ctx.beginPath();
  ctx.moveTo(xAt(0), yAt(y0));
  ctx.lineTo(xAt(n - 1), yAt(y1));
  ctx.stroke();
  ctx.restore();
}

function volPanelLayout(canvas) {
  const wrap = canvas.parentElement;
  const w = Math.max(280, Math.floor(wrap?.getBoundingClientRect().width || 600));
  const h = 56;
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  canvas.width = Math.floor(w * dpr);
  canvas.height = Math.floor(h * dpr);
  canvas.style.width = `${w}px`;
  canvas.style.height = `${h}px`;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, w, h);
  const padL = 44;
  const padR = 12;
  const padT = 6;
  const padB = 14;
  const innerW = w - padL - padR;
  const innerH = h - padT - padB;
  const isLight = document.documentElement.dataset.theme === "light";
  return { ctx, w, h, padL, padR, padT, padB, innerW, innerH, isLight };
}

function drawVolumePanel(canvas, bars) {
  const L = volPanelLayout(canvas);
  if (!L) return;
  const { ctx, w, h, padL, padR, padT, padB, innerW, innerH, isLight } = L;
  const grid = isLight ? "rgba(0,0,0,0.07)" : "rgba(255,255,255,0.08)";
  const axis = isLight ? "rgba(0,0,0,0.42)" : "rgba(255,255,255,0.42)";
  const vols = bars.map((b) => {
    const v = Number(b.v);
    return Number.isFinite(v) && v >= 0 ? v : 0;
  });
  const n = vols.length;
  if (n < 2) return;
  const mx = Math.max(...vols, 1);
  const xAt = (i) => padL + (innerW * i) / (n - 1);
  const yBase = padT + innerH;
  ctx.strokeStyle = grid;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(padL, yBase);
  ctx.lineTo(padL + innerW, yBase);
  ctx.stroke();
  const slot = n > 1 ? innerW / (n - 1) : innerW;
  const halfW = Math.max(1, Math.min(slot * 0.4, 12));
  const col = isLight ? "rgba(94,234,212,0.55)" : "rgba(94,234,212,0.42)";
  for (let i = 0; i < n; i++) {
    const v = vols[i];
    if (v <= 0) continue;
    const bh = innerH * (v / mx);
    const x = xAt(i);
    ctx.fillStyle = col;
    ctx.fillRect(x - halfW, yBase - bh, halfW * 2, Math.max(1, bh));
  }
  ctx.fillStyle = axis;
  ctx.font = "10px system-ui,sans-serif";
  ctx.textAlign = "right";
  const lab =
    mx >= 1e9 ? `${(mx / 1e9).toFixed(2)}B` : mx >= 1e6 ? `${(mx / 1e6).toFixed(2)}M` : mx >= 1e3 ? `${(mx / 1e3).toFixed(0)}k` : `${Math.round(mx)}`;
  ctx.fillText(lab, padL - 4, padT + 10);
}

/** Educational mood + B/H/S from RSI and moving averages (not a model). */
function computeMoodSignal(last, rsi, sma50, sma200) {
  const moodCls = (m) => (m === "bullish" ? "moodBull" : m === "bearish" ? "moodBear" : "moodNeu");
  const sigCls = (s) => (s === "buy" ? "sigBuy" : s === "sell" ? "sigSell" : "sigHold");
  if (rsi == null || !Number.isFinite(last)) {
    return {
      mood: "neutral",
      signal: "hold",
      moodLab: "Neutral",
      sigLab: "Hold",
      moodClass: moodCls("neutral"),
      sigClass: sigCls("hold"),
    };
  }
  const above50 = sma50 != null && last >= sma50;
  const below50 = sma50 != null && last < sma50;
  const above200 = sma200 != null && last >= sma200;
  let mood = "neutral";
  if (rsi >= 58 && (above50 || (sma50 == null && rsi >= 62))) mood = "bullish";
  else if (rsi <= 42 && (below50 || (sma50 == null && rsi <= 38))) mood = "bearish";
  else if (rsi >= 52 && above50) mood = "bullish";
  else if (rsi <= 48 && below50) mood = "bearish";
  else if (rsi >= 55 && above200 && sma50 != null && above50) mood = "bullish";
  else if (rsi <= 45 && sma200 != null && !above200) mood = "bearish";

  let signal = "hold";
  if (mood === "bullish" && rsi < 72) signal = "buy";
  else if (mood === "bullish" && rsi >= 72) signal = "hold";
  else if (mood === "bearish" || rsi <= 34) signal = "sell";
  else if (rsi >= 74) signal = "sell";

  const moodLab = mood === "bullish" ? "Bullish" : mood === "bearish" ? "Bearish" : "Neutral";
  const sigLab = signal === "buy" ? "Buy" : signal === "sell" ? "Sell" : "Hold";
  return {
    mood,
    signal,
    moodLab,
    sigLab,
    moodClass: moodCls(mood),
    sigClass: sigCls(signal),
  };
}

/** Wilder RSI at the last bar; `closes` oldest → newest. */
function rsiWilderLast(closes, period = 14) {
  if (!Array.isArray(closes) || closes.length < period + 1) return null;
  const gains = [];
  const losses = [];
  for (let i = 1; i < closes.length; i++) {
    const d = closes[i] - closes[i - 1];
    gains.push(d > 0 ? d : 0);
    losses.push(d < 0 ? -d : 0);
  }
  let avgG = 0;
  let avgL = 0;
  for (let i = 0; i < period; i++) {
    avgG += gains[i];
    avgL += losses[i];
  }
  avgG /= period;
  avgL /= period;
  for (let i = period; i < gains.length; i++) {
    avgG = (avgG * (period - 1) + gains[i]) / period;
    avgL = (avgL * (period - 1) + losses[i]) / period;
  }
  if (avgL === 0) return avgG === 0 ? 50 : 100;
  const rs = avgG / avgL;
  return 100 - 100 / (1 + rs);
}

/**
 * Build HTML for SMA / RSI / Bollinger / RS55 from daily `bars` (`c` = close).
 * @param {null | { rs55: number, benchLabel: string, benchSym?: string }} rsExtra
 */
function technicalsHtmlFromBars(bars, rsExtra) {
  const closes = bars.map((b) => Number(b.c)).filter((x) => Number.isFinite(x));
  if (closes.length < 15) {
    return `<p class="sml muted">Technicals: need at least ~15 daily closes for RSI (try 1y range).</p>`;
  }
  const last = closes[closes.length - 1];
  const rsi = rsiWilderLast(closes, 14);
  const sma20 = smaLast(closes, 20);
  const sma50 = smaLast(closes, 50);
  const sma200 = smaLast(closes, 200);
  const ms = computeMoodSignal(last, rsi, sma50, sma200);
  const rsiHint =
    rsi == null
      ? ""
      : rsi >= 70
        ? ` <span class="sml muted">(often “overbought” in classic reads)</span>`
        : rsi <= 30
          ? ` <span class="sml muted">(often “oversold”)</span>`
          : "";
  const row = (lab, body) =>
    `<div class="techRow"><span class="techLab">${esc(lab)}</span><span class="techVal">${body}</span></div>`;
  let html = `<div class="h3">Technicals (daily closes)</div>
    <div class="moodSigRow" role="group" aria-label="Mood and suggestion">
      <div class="moodSigCell"><span class="sml muted">Mood</span><div class="moLabel ${ms.moodClass}" aria-label="Mood">${esc(ms.moodLab)}</div></div>
      <div class="moodSigCell"><span class="sml muted">Suggestion</span><div class="moLabel ${ms.sigClass}" aria-label="Suggestion">${esc(ms.sigLab)}</div></div>
    </div>
    <p class="sml muted moodSigNote">Mood and Buy/Hold/Sell are <strong>rule-based labels</strong> from RSI and moving averages — not a trading system. Same logic runs on <strong>session-cached</strong> bars when live data is unavailable.</p>
    <div class="techGrid">`;
  html += row("Last close", `<strong>${esc(fmtN(last, 2))}</strong>`);
  if (rsi != null) html += row("RSI (14)", `<strong>${esc(fmtN(rsi, 1))}</strong>${rsiHint}`);
  if (sma20 != null) {
    const vs = last >= sma20 ? "above" : "below";
    html += row("SMA (20)", `<strong>${esc(fmtN(sma20, 2))}</strong> <span class="sml muted">· price ${vs} SMA20</span>`);
  } else {
    html += row("SMA (20)", `<span class="muted">— (need 20+ bars)</span>`);
  }
  if (sma50 != null) {
    const vs = last >= sma50 ? "above" : "below";
    html += row("SMA (50)", `<strong>${esc(fmtN(sma50, 2))}</strong> <span class="sml muted">· price ${vs} SMA50</span>`);
  } else {
    html += row("SMA (50)", `<span class="muted">— (need 50+ bars)</span>`);
  }
  if (sma200 != null) {
    const vs = last >= sma200 ? "above" : "below";
    html += row("SMA (200)", `<strong>${esc(fmtN(sma200, 2))}</strong> <span class="sml muted">· price ${vs} SMA200</span>`);
  } else {
    html += row("SMA (200)", `<span class="muted">— (need 200+ bars — try range=max)</span>`);
  }
  if (closes.length >= 20) {
    const { mid, upper, lower } = bollingerSeries(closes, 20, 2);
    const i = closes.length - 1;
    const mu = upper[i];
    const ml = lower[i];
    const mm = mid[i];
    if (mu != null && ml != null && mm != null) {
      const pctB = mu !== ml ? ((last - ml) / (mu - ml)) * 100 : null;
      const bHint =
        pctB == null
          ? ""
          : pctB > 100
            ? ` <span class="sml muted">(%B above 100 — over upper band)</span>`
            : pctB < 0
              ? ` <span class="sml muted">(%B under 0 — below lower band)</span>`
              : ` <span class="sml muted">(%B ${esc(fmtN(pctB, 0))})</span>`;
      html += row(
        "Bollinger (20, 2σ)",
        `U <strong>${esc(fmtN(mu, 2))}</strong> · Mid <strong>${esc(fmtN(mm, 2))}</strong> · L <strong>${esc(fmtN(ml, 2))}</strong>${bHint}`,
      );
    }
  } else {
    html += row("Bollinger (20, 2σ)", `<span class="muted">— (need 20+ bars)</span>`);
  }
  if (rsExtra && rsExtra.rs55 != null && Number.isFinite(rsExtra.rs55)) {
    const pct = rsExtra.rs55 * 100;
    const sign = pct >= 0 ? "+" : "";
    const bias = pct > 0 ? "above" : pct < 0 ? "below" : "at";
    html += row(
      `RS55 vs ${esc(rsExtra.benchLabel)}`,
      `<strong>${sign}${esc(fmtN(pct, 2))}%</strong> <span class="sml muted">· stock/bench close ratio ${bias} its 55-day mean (same calendar days)</span>`,
    );
  } else {
    html += row(
      "RS55 (vs benchmark)",
      `<span class="muted">— (needs 55 overlapping sessions vs the index proxy, or benchmark history unavailable)</span>`,
    );
  }
  const benchExplain =
    rsExtra && rsExtra.benchLabel
      ? esc(rsExtra.benchLabel)
      : "Nifty 50 for NSE/BSE, otherwise SPY as an S&amp;P 500 proxy";
  html += `</div>
  <details class="techDetails mt">
    <summary class="sml">What each line means (incl. RS55 in detail)</summary>
    <div class="techDetailsBody sml muted">
      <p><strong>Last close</strong> — Latest daily settlement price in the history series (same bar set as the chart).</p>
      <p><strong>RSI (14)</strong> — Wilder RSI on the last 14 daily closes. Roughly: values toward 100 suggest strong recent up-moves vs down-moves; toward 0 the opposite. Classic commentary uses 70/30 as “stretched” zones — not automatic buy/sell rules.</p>
      <p><strong>SMA (20 / 50 / 200)</strong> — Simple moving average of the last N closes. “Price above/below” compares the last close to that average to describe short-, medium-, or long-term level.</p>
      <p><strong>Bollinger (20, 2σ)</strong> — Middle band = 20-day SMA of close; upper/lower = middle ± 2× the standard deviation of closes in that window. <strong>%B</strong> (when shown) places the last close between the bands (0 = on lower, 100 = on upper).</p>
      <p><strong>RS55 vs benchmark (${benchExplain})</strong> — For each calendar session we use <strong>R = stock close ÷ benchmark close</strong> (same timestamp). We take the <strong>last 55</strong> sessions where both exist, compute the mean <strong>M</strong> of R, then show <strong>(R_last ÷ M − 1) × 100%</strong>. So: <strong>+10%</strong> means today’s relative price vs the index is about <strong>10% above</strong> its own 55-day average ratio; <strong>negative</strong> means below that average. It is a <strong>simplified relative-strength lens</strong> inspired by public RS55 discussions — not the full proprietary rule set from any vendor or educator.</p>
      <p><strong>Mood (Bullish / Neutral / Bearish)</strong> — Coloured label from a small rule table using RSI plus price vs SMA50/SMA200. It summarises tone, not a forecast.</p>
      <p><strong>Suggestion (Buy / Hold / Sell)</strong> — Same rule table: e.g. very high RSI tends toward Hold/Sell for “stretched” caution; weak RSI + weak price position tends toward Sell in this toy logic. <strong>Educational only</strong> — not investment advice.</p>
    </div>
  </details>
  <p class="sml muted techDisclaimer mt">From the same history as the chart (including session cache when live data fails). Educational only — not investment advice.</p>`;
  return html;
}

/** Plain-text technical summary for the server-side LLM (same math as Technicals card). */
function technicalSummaryPlain(bars, rsExtra) {
  const closes = bars.map((b) => Number(b.c)).filter((x) => Number.isFinite(x));
  if (closes.length < 15) return "Technicals: fewer than ~15 daily closes; RSI not computed.";
  const last = closes[closes.length - 1];
  const rsi = rsiWilderLast(closes, 14);
  const sma20 = smaLast(closes, 20);
  const sma50 = smaLast(closes, 50);
  const sma200 = smaLast(closes, 200);
  const ms = computeMoodSignal(last, rsi, sma50, sma200);
  const parts = [
    `last_close=${fmtN(last, 4)}`,
    rsi != null ? `rsi14=${fmtN(rsi, 2)}` : "",
    sma20 != null ? `sma20=${fmtN(sma20, 4)} (price ${last >= sma20 ? "above" : "below"})` : "",
    sma50 != null ? `sma50=${fmtN(sma50, 4)} (price ${last >= sma50 ? "above" : "below"})` : "",
    sma200 != null ? `sma200=${fmtN(sma200, 4)} (price ${last >= sma200 ? "above" : "below"})` : "",
    `app_mood=${ms.moodLab} app_suggestion=${ms.sigLab} (rule-based, not a forecast)`,
  ];
  if (closes.length >= 20) {
    const { mid, upper, lower } = bollingerSeries(closes, 20, 2);
    const i = closes.length - 1;
    const mu = upper[i];
    const ml = lower[i];
    const mm = mid[i];
    if (mu != null && ml != null && mm != null) {
      const pctB = mu !== ml ? ((last - ml) / (mu - ml)) * 100 : null;
      parts.push(
        `bollinger20_2sigma: upper=${fmtN(mu, 4)} mid=${fmtN(mm, 4)} lower=${fmtN(ml, 4)}` +
          (pctB != null ? ` pctB=${fmtN(pctB, 1)}` : ""),
      );
    }
  }
  if (rsExtra && rsExtra.rs55 != null && Number.isFinite(rsExtra.rs55)) {
    const pct = rsExtra.rs55 * 100;
    const sign = pct >= 0 ? "+" : "";
    parts.push(`rs55_vs_${String(rsExtra.benchLabel || "benchmark").replace(/\s+/g, "_")}=${sign}${fmtN(pct, 2)}% vs 55d mean stock/bench close ratio`);
  }
  return parts.filter(Boolean).join("; ");
}

function technicalSummaryPlainFromContext() {
  const h = _instrHistoryContext;
  if (!h?.bars?.length) return "";
  return technicalSummaryPlain(h.bars, h.rsExtra);
}

/** Shared canvas sizing + DPR for history charts. */
function histChartLayout(canvas) {
  const wrap = canvas.parentElement;
  const w = Math.max(280, Math.floor(wrap?.getBoundingClientRect().width || 600));
  const h = 220;
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  canvas.width = Math.floor(w * dpr);
  canvas.height = Math.floor(h * dpr);
  canvas.style.width = `${w}px`;
  canvas.style.height = `${h}px`;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, w, h);
  const padL = 44;
  const padR = 12;
  const padT = 14;
  const padB = 28;
  const innerW = w - padL - padR;
  const innerH = h - padT - padB;
  const isLight = document.documentElement.dataset.theme === "light";
  return { ctx, w, h, padL, padR, padT, padB, innerW, innerH, isLight };
}

function barsToOhlc(bars) {
  const out = [];
  for (const b of bars) {
    const c = Number(b.c);
    if (!Number.isFinite(c)) continue;
    let o = Number(b.o);
    let h = Number(b.h);
    let l = Number(b.l);
    if (!Number.isFinite(o)) o = c;
    if (!Number.isFinite(h)) h = Math.max(o, c);
    if (!Number.isFinite(l)) l = Math.min(o, c);
    out.push({ t: b.t, o, h, l, c });
  }
  return out;
}

function computeHeikinAshi(ohlc) {
  const out = [];
  for (let i = 0; i < ohlc.length; i++) {
    const b = ohlc[i];
    const hc = (b.o + b.h + b.l + b.c) / 4;
    const ho = i === 0 ? (b.o + b.c) / 2 : (out[i - 1].o + out[i - 1].c) / 2;
    const hh = Math.max(b.h, ho, hc);
    const hl = Math.min(b.l, ho, hc);
    out.push({ t: b.t, o: ho, h: hh, l: hl, c: hc });
  }
  return out;
}

function drawHistoryChart(canvas, bars, rangeLabel, chartType) {
  const ohlc = barsToOhlc(bars);
  const studyCloses = ohlc.map((b) => b.c);
  const overlays = getChartOverlayState();
  if (chartType === "candle" && ohlc.length >= 2) {
    drawOhlcCandles(canvas, ohlc, rangeLabel, "OHLC", overlays, studyCloses);
  } else if (chartType === "heikin" && ohlc.length >= 2) {
    const ha = computeHeikinAshi(ohlc);
    drawOhlcCandles(canvas, ha, rangeLabel, "Heikin Ashi", overlays, studyCloses);
  } else {
    drawAreaCloseHistory(canvas, bars, rangeLabel, overlays);
  }
  syncVolPanel();
}

function drawAreaCloseHistory(canvas, bars, rangeLabel, overlays) {
  const L = histChartLayout(canvas);
  if (!L) return;
  const { ctx, w, h, padL, padR, padT, padB, innerW, innerH, isLight } = L;
  const grid = isLight ? "rgba(0,0,0,0.08)" : "rgba(255,255,255,0.08)";
  const axis = isLight ? "rgba(0,0,0,0.45)" : "rgba(255,255,255,0.45)";
  const line = getComputedStyle(document.documentElement).getPropertyValue("--a").trim() || "#5eead4";
  const fill0 = isLight ? "rgba(94,234,212,0.2)" : "rgba(94,234,212,0.12)";
  const closes = bars.map((b) => Number(b.c)).filter((x) => Number.isFinite(x));
  if (closes.length < 2) return;
  let lo = Math.min(...closes);
  let hi = Math.max(...closes);
  const n = closes.length;
  const seriesForRange = [];
  if (overlays.ma20 && n >= 20) seriesForRange.push(smaSeries(closes, 20));
  if (overlays.ma50 && n >= 50) seriesForRange.push(smaSeries(closes, 50));
  if (overlays.bb && n >= 20) {
    const bb = bollingerSeries(closes, 20, 2);
    seriesForRange.push(bb.upper, bb.lower);
  }
  ({ lo, hi } = extendLoHi(lo, hi, seriesForRange));
  if (hi === lo) {
    lo *= 0.995;
    hi *= 1.005;
  }
  const padY = (hi - lo) * 0.06;
  lo -= padY;
  hi += padY;
  const xAt = (i) => padL + (innerW * i) / (n - 1);
  const yAt = (v) => padT + innerH * (1 - (v - lo) / (hi - lo));
  for (let g = 0; g <= 4; g++) {
    const y = padT + (innerH * g) / 4;
    ctx.strokeStyle = grid;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padL, y);
    ctx.lineTo(padL + innerW, y);
    ctx.stroke();
    const val = hi - ((hi - lo) * g) / 4;
    ctx.fillStyle = axis;
    ctx.font = "11px system-ui,sans-serif";
    ctx.textAlign = "right";
    ctx.fillText(val < 100 ? val.toFixed(2) : val.toFixed(1), padL - 6, y + 4);
  }
  drawStudyOverlays(ctx, { padL, innerW, padT, innerH, lo, hi, n, isLight }, closes, overlays);
  ctx.beginPath();
  ctx.moveTo(xAt(0), yAt(closes[0]));
  for (let i = 1; i < n; i++) ctx.lineTo(xAt(i), yAt(closes[i]));
  ctx.strokeStyle = line;
  ctx.lineWidth = 2;
  ctx.stroke();
  ctx.beginPath();
  ctx.moveTo(xAt(0), padT + innerH);
  ctx.lineTo(xAt(0), yAt(closes[0]));
  for (let i = 1; i < n; i++) ctx.lineTo(xAt(i), yAt(closes[i]));
  ctx.lineTo(xAt(n - 1), padT + innerH);
  ctx.closePath();
  const grad = ctx.createLinearGradient(0, padT, 0, padT + innerH);
  grad.addColorStop(0, fill0);
  grad.addColorStop(1, "transparent");
  ctx.fillStyle = grad;
  ctx.fill();
  drawTrendlineOverlay(ctx, { padL, innerW, padT, innerH, lo, hi, n, isLight }, closes, overlays);
  ctx.fillStyle = axis;
  ctx.font = "11px system-ui,sans-serif";
  ctx.textAlign = "left";
  const t0 = bars[0]?.t;
  const t1 = bars[n - 1]?.t;
  if (t0 && t1) {
    const d0 = new Date(Number(t0) * 1000);
    const d1 = new Date(Number(t1) * 1000);
    const opt = { month: "short", day: "numeric" };
    ctx.fillText(d0.toLocaleDateString(undefined, opt), padL, h - 8);
    ctx.textAlign = "right";
    ctx.fillText(d1.toLocaleDateString(undefined, opt), w - padR, h - 8);
  }
  ctx.textAlign = "left";
  ctx.fillText(`${rangeLabel} · close (area)`, padL, 12);
}

function drawOhlcCandles(canvas, ohlc, rangeLabel, modeLabel, overlays, studyCloses) {
  const L = histChartLayout(canvas);
  if (!L) return;
  const { ctx, w, h, padL, padR, padT, padB, innerW, innerH, isLight } = L;
  const grid = isLight ? "rgba(0,0,0,0.08)" : "rgba(255,255,255,0.08)";
  const axis = isLight ? "rgba(0,0,0,0.45)" : "rgba(255,255,255,0.45)";
  const bull = getComputedStyle(document.documentElement).getPropertyValue("--a").trim() || "#5eead4";
  const bear = isLight ? "#dc2626" : "#f87171";
  const n = ohlc.length;
  let lo = Infinity;
  let hi = -Infinity;
  for (const b of ohlc) {
    lo = Math.min(lo, b.l);
    hi = Math.max(hi, b.h);
  }
  if (!Number.isFinite(lo) || !Number.isFinite(hi)) return;
  const sc = Array.isArray(studyCloses) && studyCloses.length === n ? studyCloses : null;
  const seriesForRange = [];
  if (sc) {
    if (overlays.ma20 && n >= 20) seriesForRange.push(smaSeries(sc, 20));
    if (overlays.ma50 && n >= 50) seriesForRange.push(smaSeries(sc, 50));
    if (overlays.bb && n >= 20) {
      const bb = bollingerSeries(sc, 20, 2);
      seriesForRange.push(bb.upper, bb.lower);
    }
  }
  ({ lo, hi } = extendLoHi(lo, hi, seriesForRange));
  if (hi === lo) {
    lo *= 0.995;
    hi *= 1.005;
  }
  const padY = (hi - lo) * 0.06;
  lo -= padY;
  hi += padY;
  const xAt = (i) => padL + (innerW * i) / (n - 1);
  const yAt = (v) => padT + innerH * (1 - (v - lo) / (hi - lo));
  for (let g = 0; g <= 4; g++) {
    const y = padT + (innerH * g) / 4;
    ctx.strokeStyle = grid;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(padL, y);
    ctx.lineTo(padL + innerW, y);
    ctx.stroke();
    const val = hi - ((hi - lo) * g) / 4;
    ctx.fillStyle = axis;
    ctx.font = "11px system-ui,sans-serif";
    ctx.textAlign = "right";
    ctx.fillText(val < 100 ? val.toFixed(2) : val.toFixed(1), padL - 6, y + 4);
  }
  if (sc) drawStudyOverlays(ctx, { padL, innerW, padT, innerH, lo, hi, n, isLight }, sc, overlays);
  const slot = n > 1 ? innerW / (n - 1) : innerW;
  const halfW = Math.max(1, Math.min(slot * 0.35, 10));
  const eps = (hi - lo) * 0.0001;
  for (let i = 0; i < n; i++) {
    const b = ohlc[i];
    const cx = xAt(i);
    ctx.strokeStyle = axis;
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(cx, yAt(b.h));
    ctx.lineTo(cx, yAt(b.l));
    ctx.stroke();
    const yTop = yAt(Math.max(b.o, b.c));
    const yBot = yAt(Math.min(b.o, b.c));
    const bullBar = b.c >= b.o;
    ctx.fillStyle = bullBar ? bull : bear;
    ctx.strokeStyle = bullBar ? bull : bear;
    const bh = Math.max(1, yBot - yTop);
    ctx.fillRect(cx - halfW, yTop, halfW * 2, bh);
    if (Math.abs(b.c - b.o) < eps) {
      ctx.beginPath();
      ctx.moveTo(cx - halfW, yTop);
      ctx.lineTo(cx + halfW, yTop);
      ctx.stroke();
    }
  }
  if (sc) drawTrendlineOverlay(ctx, { padL, innerW, padT, innerH, lo, hi, n, isLight }, sc, overlays);
  ctx.fillStyle = axis;
  ctx.font = "11px system-ui,sans-serif";
  const t0 = ohlc[0]?.t;
  const t1 = ohlc[n - 1]?.t;
  if (t0 && t1) {
    const d0 = new Date(Number(t0) * 1000);
    const d1 = new Date(Number(t1) * 1000);
    const opt = { month: "short", day: "numeric" };
    ctx.textAlign = "left";
    ctx.fillText(d0.toLocaleDateString(undefined, opt), padL, h - 8);
    ctx.textAlign = "right";
    ctx.fillText(d1.toLocaleDateString(undefined, opt), w - padR, h - 8);
  }
  ctx.textAlign = "left";
  ctx.fillText(`${rangeLabel} · ${modeLabel}`, padL, 12);
}

function quoteCard(q) {
  if (!q) return `<p class="muted">No data</p>`;
  const pr = esc(q._provider || "");
  const note = esc(q._note || "");
  const ccy = (q.currency || "").toUpperCase();
  const px = Number.isFinite(Number(q.price)) ? fmtMoney(ccy, Number(q.price)) : "—";
  const pc = Number.isFinite(Number(q.previousClose)) ? fmtMoney(ccy, Number(q.previousClose)) : "—";
  const op = Number.isFinite(Number(q.open)) ? fmtMoney(ccy, Number(q.open)) : "—";
  const hi = Number.isFinite(Number(q.dayHigh)) ? fmtMoney(ccy, Number(q.dayHigh)) : "—";
  const lo = Number.isFinite(Number(q.dayLow)) ? fmtMoney(ccy, Number(q.dayLow)) : "—";
  const yHi = Number(q.yearHigh);
  const yLo = Number(q.yearLow);
  const w52 =
    Number.isFinite(yHi) && Number.isFinite(yLo)
      ? `<div class="grid2 sml mt">
          <div>52w high <strong>${esc(fmtMoney(ccy, yHi))}</strong></div>
          <div>52w low <strong>${esc(fmtMoney(ccy, yLo))}</strong></div>
        </div>`
      : `<p class="sml muted mt">52-week range: not returned by this quote source.</p>`;
  const chgPct = Number(q.changePercent);
  const chgRow =
    Number.isFinite(chgPct) && chgPct !== 0
      ? `<p class="sml mt">Session move <strong class="${chgPct >= 0 ? "plp" : "pln"}">${esc(fmtN(Math.abs(chgPct) <= 1 ? chgPct * 100 : chgPct, 2))}%</strong>${Number.isFinite(Number(q.changeAmount)) ? ` <span class="muted">(${esc(fmtMoney(ccy, Number(q.changeAmount)))})</span>` : ""}</p>`
      : "";
  const avgV = Number(q.averageVolume);
  const av3 = Number(q.avgVolume3Mo);
  const av10 = Number(q.avgVolume10d);
  const volCmp =
    Number.isFinite(Number(q.volume)) && (Number.isFinite(avgV) || Number.isFinite(av3))
      ? `<p class="sml muted mt">Volume vs averages: session <strong>${esc(fmtVolShort(Number(q.volume)))}</strong>${Number.isFinite(avgV) ? ` · avg <strong>${esc(fmtVolShort(avgV))}</strong>` : ""}${Number.isFinite(av3) ? ` · 3mo avg <strong>${esc(fmtVolShort(av3))}</strong>` : ""}${Number.isFinite(av10) ? ` · 10d avg <strong>${esc(fmtVolShort(av10))}</strong>` : ""}</p>`
      : "";
  const mcap = Number(q.marketCap);
  const pe = Number(q.peRatio ?? q.trailingPE);
  const tiny =
    Number.isFinite(mcap) || Number.isFinite(pe)
      ? `<p class="sml muted mt">${Number.isFinite(mcap) ? `Mkt cap <strong>${esc(fmtBigNumber(mcap))}</strong>` : ""}${Number.isFinite(mcap) && Number.isFinite(pe) ? " · " : ""}${Number.isFinite(pe) ? `P/E <strong>${esc(fmtN(pe, 2))}</strong>` : ""}</p>`
      : "";
  return `<div class="card2">
    <div class="h2">${esc(q.symbol || "")} ${q.name ? "· " + esc(q.name) : ""}</div>
    <div class="big">${px} <span class="muted">${esc(ccy)}</span></div>
    <div class="sml muted">${pr ? `Source: ${pr}` : ""}${pr && note ? " · " : ""}${note}</div>
    ${chgRow}
    <div class="grid3 sml mt">
      <div>Prev <strong>${esc(pc)}</strong></div>
      <div>Open <strong>${esc(op)}</strong></div>
      <div>Hi/Lo <strong>${esc(hi)}</strong> / <strong>${esc(lo)}</strong></div>
    </div>
    ${w52}
    ${tiny}
    ${volCmp}
  </div>`;
}

/** Exchange label → country (aligns with server `server.py` `_EXCH_COUNTRY_CCY`). */
const PF_EXCH_COUNTRY = {
  NASDAQ: "United States",
  NYSE: "United States",
  "NYSE ARCA": "United States",
  "NYSE AMERICAN": "United States",
  AMEX: "United States",
  NMS: "United States",
  NGM: "United States",
  NCM: "United States",
  NYQ: "United States",
  PCX: "United States",
  BATS: "United States",
  NASDAQGS: "United States",
  "NASDAQ GM": "United States",
  "NASDAQ CM": "United States",
  NSE: "India",
  NSI: "India",
  BSE: "India",
  BOM: "India",
  LSE: "United Kingdom",
  LON: "United Kingdom",
  XETRA: "Germany",
  GER: "Germany",
  ETR: "Germany",
  FRA: "Germany",
  AMS: "Netherlands",
  AS: "Netherlands",
  PAR: "France",
  EPA: "France",
  MIL: "Italy",
  BIT: "Italy",
  SWX: "Switzerland",
  SW: "Switzerland",
  STO: "Sweden",
  HEL: "Finland",
  WSE: "Poland",
  EL: "Greece",
  ISE: "Ireland",
  TSX: "Canada",
  TOR: "Canada",
  V: "Canada",
  CN: "Canada",
};

/** Yahoo-style suffix → country (longest match wins). */
const PF_SYM_SUFFIX = [
  [".NSE", "India"],
  [".NS", "India"],
  [".TW", "Taiwan"],
  [".TO", "Canada"],
  [".KS", "South Korea"],
  [".BO", "India"],
  [".HK", "Hong Kong"],
  [".DE", "Germany"],
  [".SW", "Switzerland"],
  [".PA", "France"],
  [".AS", "Netherlands"],
  [".V", "Canada"],
  [".AX", "Australia"],
  [".SI", "Singapore"],
  [".SA", "Brazil"],
  [".MX", "Mexico"],
  [".WA", "Poland"],
  [".ST", "Sweden"],
  [".OL", "Norway"],
  [".CO", "Denmark"],
  [".HE", "Finland"],
  [".MI", "Italy"],
  [".L", "United Kingdom"],
  [".T", "Japan"],
];

const PF_CCY_REGION = {
  INR: "India (currency)",
  USD: "United States (currency)",
  EUR: "Euro area (currency)",
  GBP: "United Kingdom (currency)",
  CHF: "Switzerland (currency)",
  JPY: "Japan (currency)",
  CAD: "Canada (currency)",
  AUD: "Australia (currency)",
  HKD: "Hong Kong (currency)",
  SGD: "Singapore (currency)",
  KRW: "South Korea (currency)",
  PLN: "Poland (currency)",
  SEK: "Sweden (currency)",
  NOK: "Norway (currency)",
  DKK: "Denmark (currency)",
  CNY: "China (currency)",
  TWD: "Taiwan (currency)",
  ZAR: "South Africa (currency)",
  BRL: "Brazil (currency)",
};

const PF_PIE_PAL = ["#2dd4bf", "#a78bfa", "#fb7185", "#fbbf24", "#38bdf8", "#94a3b8", "#fb923c", "#4ade80", "#818cf8", "#f472b6"];

/** Portfolio row weight for charts: market value, else cost, else qty. */
function pfRowWeight(r) {
  if (isInsuranceAltRow(r)) {
    const cur = num(r.currentValue);
    if (cur > 0) return cur;
    const inv = insuranceInvestedTotal(r);
    return inv > 0 ? inv : 0;
  }
  if (isFdAltRow(r)) {
    const cur = num(r.currentValue);
    if (cur > 0) return cur;
    const pr = num(r.principal);
    return pr > 0 ? pr : 0;
  }
  const qty = num(r.qty);
  const last = num(r.last);
  const avg = num(r.avg);
  const px = last > 0 ? last : avg > 0 ? avg : 0;
  let w = qty * px;
  if (w <= 0 && qty > 0) w = qty;
  return w;
}

function pfExchHead(ex) {
  const u = String(ex || "")
    .trim()
    .toUpperCase();
  if (!u) return "";
  const head = u.split(/\s+/)[0];
  return head;
}

function pfInferCountry(r) {
  const ex = String(r.ex || "").trim().toUpperCase();
  if (ex && PF_EXCH_COUNTRY[ex]) return PF_EXCH_COUNTRY[ex];
  const head = pfExchHead(r.ex);
  if (head && PF_EXCH_COUNTRY[head]) return PF_EXCH_COUNTRY[head];
  const sym = String(r.sym || "").toUpperCase();
  for (const [suf, ctry] of PF_SYM_SUFFIX) {
    if (sym.endsWith(suf.toUpperCase())) return ctry;
  }
  const ccy = String(r.ccy || "").toUpperCase();
  if (ccy && PF_CCY_REGION[ccy]) return PF_CCY_REGION[ccy];
  return "Other / unknown";
}

function pfNormQuoteKind(qt) {
  const u = String(qt || "").toUpperCase();
  if (!u) return "";
  if (/\b(ETF|ETN|ETC)\b/.test(u) || u.includes("MUTUAL") || u.includes("FUND") || u.includes("INDEX")) {
    return "Funds / ETFs / indexes";
  }
  if (/\b(EQUITY|STOCK|ADR)\b/.test(u)) return "Equities & ADRs";
  return u.length <= 24 ? u : `${u.slice(0, 22)}…`;
}

function pfInferKindFromName(sym, nm) {
  const t = `${nm} ${sym}`.toUpperCase();
  if (
    /\bETF\b|ETC|ETN|UCITS|ISHARES|VANGUARD|SPDR|WISDOMTREE|XTRACKERS|LYXOR|AMUNDI|INVESCO|PROSHARES|DIREXION|REIT|TRUST|INDEX|ACCUMULATING|DISTRIBUT|MULTIFACTOR|BOND\s|GOVERNMENT\s|CORP\s|HIGH\sYIELD/.test(
      t,
    )
  ) {
    return "Funds / ETFs (name hint)";
  }
  return "Equities & other";
}

function pfInstrumentBucket(r) {
  const fromQ = pfNormQuoteKind(r.pfKind);
  if (fromQ) return fromQ;
  return pfInferKindFromName(r.sym, r.nm);
}

/**
 * Heuristic: coin held on the **Trading 212** stock tab (positions sometimes land in `j.t212` only).
 * Used to split T212 vs **Crypto (T212)** in the combined € roll-up without double-counting.
 */
function isCryptoLikeT212Row(r) {
  if (!r || typeof r !== "object") return false;
  const ex = String(r.ex || "").toUpperCase();
  if (/\bCRYPTO(CURRENCY)?\b/.test(ex)) return true;
  const pk = String(r.pfKind || "").toUpperCase();
  if (pk.includes("CRYPTO") || (pk.includes("CURRENCY") && /DIGITAL|CRYPTO/.test(pk))) return true;
  const sym0 = String(r.sym || "")
    .trim()
    .toUpperCase();
  const sym = sym0.split(/[-_/]/)[0] || sym0;
  if (
    /^(BTC|XBT|ETH|XRP|ADA|DOGE|SOL|DOT|AVAX|LTC|BCH|ETC|LINK|UNI|ATOM|NEAR|XLM|ALGO|PEPE|SHIB|TRX|TON|AAVE|CRV|LDO|MKR|INJ|RNDR|ARB|OP|IMX|GRT|HBAR|ICP|FIL|APT|STX|SUI|VET|QNT|BONK|WIF|JUP|FET|TAO|PYTH|STRK|TIA|ORDI|RUNE|ETN)$/.test(sym)
  ) {
    return true;
  }
  const nm = String(r.nm || "").toUpperCase();
  if (/\b(BITCOIN|ETHEREUM|DOGECOIN|CRYPTOCURREN|CRYPTO\s*ASSET|SHIBA|LITECOIN|RIPPLE|SOLANA|CARDANO|AVALANCHE|POLKADOT|CHAINLINK|UNISWAP)\b/.test(nm)) return true;
  return false;
}

/** Single listing currency for a ledger, or "" if mixed / missing. */
function pfPrimaryCcy(rows) {
  const g = new Set(rows.map((x) => normalizeCcyForFx(x.ccy)).filter(Boolean));
  return g.size === 1 ? [...g][0] : "";
}

/**
 * NSE/BSE-style rows often ship without `nm` — classify by symbol + scraped name text.
 * Returns `null` when no India-specific rule matches (caller continues with global rules).
 */
function pfStrategyIndianHint(symU, t) {
  const s = symU || "";
  const u = `${s} ${t}`.toUpperCase();
  if (!s && !String(t || "").trim()) return null;
  if (/\b(BITCOIN|CRYPTO|ETHEREUM)\b/.test(u)) return null;

  if (/LIQUID|OVERNIGHT|PARAGPARIK|MONEY.?MARKET|ULTRA.?SHORT|SHORT.?DURATION|GILT\s*FUND|CORP(?:ORATE)?\s*BOND/i.test(u)) {
    return "Cash, liquid & short bond funds (India)";
  }
  if (/\b(BEES|ETF|INDEX\s*FUND|NIFTY\s*50|NIFTY\s*BANK|SENSEX|GOLD\s*BEES|SILVER\s*BEES|LOW\s*VOL|QUALITY|MOMENTUM)\b/i.test(u)) {
    return "Index & thematic ETFs (India)";
  }
  if (/^(INFY|TCS|WIPRO|HCLTECH|TECHM|LTIM|MPHASIS|COFORGE|PERSISTENT|SONATSOFTW|ZENSAR|CYIENT|KPITTECH|ECLERX|NEWGEN|MASTEK|SONATA|ITDCEM)$/i.test(s)) {
    return "IT services & software (India)";
  }
  if (s.endsWith("BANK") || /^(HDFCBANK|ICICIBANK|SBIN|AXISBANK|KOTAKBANK|INDUSINDBK|BANKBARODA|YESBANK|IDFCFIRSTB|FEDERALBNK|UNIONBANK|AUBANK|KARURVYSYA|TMB|PNB|IOB|MAHABANK|CENTRALBK|CORPORATION|INDIANB|SOUTHBANK|RBLBANK|DCBBANK|CSBBANK)$/i.test(s)) {
    return "Banks (India)";
  }
  if (/^(RECLTD|CHOLAHLDNG|BAJFIN|LICHSGFIN|MUTHOOTFIN|MANAPPURAM|SHRIRAMFIN|HDFCAMC|MFSL|NAM-INDIA|POONAWALLA|UTIAMC|CREDITACC|IIFL|LTF)$/i.test(s) || /\b(NBFC|HOUSING\s*FIN|FINANCE\s*&\s*INVEST)\b/i.test(t)) {
    return "NBFC & diversified finance (India)";
  }
  if (/^(PFC|IREDA)$/i.test(s)) return "Infrastructure finance (India)";
  if (/^(SUNPHARMA|DRREDDY|CIPLA|DIVISLAB|AUROPHARMA|LUPIN|BIOCON|TORNTPHARM|ALKEM|GLENMARK|GRANULES|LAURUSLABS|NATCOPHARM|ZYDUSLIFE|APLLTD|MAXHEALTH|FORTIS|APOLLOHOSP)$/i.test(s) || /\b(PHARMA|PHARMACEUT|HOSPITAL)\b/i.test(t)) {
    return "Healthcare & pharma (India)";
  }
  if (/^(MARUTI|TATAMOTORS|EICHERMOT|BAJAJ-AUTO|M&M|HEROMOTOCO|TVSMOTOR|ASHOKLEY|FORCEMOT|OLECTRA|SONACOMS)$/i.test(s) || /\b(MARUTI|TATA\s*MOTORS|EICHER)\b/i.test(t)) {
    return "Automotive (India)";
  }
  if (/^(ITC|HINDUNILVR|NESTLEIND|DABUR|TATACONSUM|BRITANNIA|MARICO|GODREJCP|RADICO|UNITDSPR|UBL|VBL|JUBLFOOD|TRENT|DMART|PAGEIND|TITAN)$/i.test(s)) {
    return "Consumer brands & retail (India)";
  }
  if (/^(RELIANCE|ONGC|OIL|IOC|BPCL|HPCL|GAIL|PETRONET|ATGL|ADANIGREEN|ADANIENT|ADANIPOWER|ADANIGAS|TORNTPOWER|NHPC|SJVN|JSWENERGY)$/i.test(s)) {
    return "Energy & utilities (India)";
  }
  if (/^(COALINDIA|NMDC|VEDL|HINDCOPPER|MOIL)$/i.test(s) || /\b(COAL\s*INDIA|MINING)\b/i.test(t)) {
    return "Mining & materials (India)";
  }
  if (/^(TATASTEEL|JSWSTEEL|SAIL|HINDALCO|NATIONALUM|APLAPOLLO|RATNAMANI|JSL)$/i.test(s)) {
    return "Steel & industrials (India)";
  }
  if (/^(ULTRACEMCO|ACC|AMBUJACEM|SHREECEM|DALMIACEM|RAMCOCEM|JKCEMENT|INDIACEM|ORIENTCEM)$/i.test(s)) {
    return "Cement & building materials (India)";
  }
  if (/^(LT|SIEMENS|ABB|THERMAX|BHEL|GREAVESCOT|CUMMINSIND|VOLTAS|HAL|BEL|BDL|DATAPATTNS|GRSE|MAZDOCK)$/i.test(s) || /\b(LARSEN|L\s*&\s*T|SIEMENS\s*IND)\b/i.test(t)) {
    return "Capital goods & defence (India)";
  }
  if (/^(NTPC|POWERGRID|TATAPOWER|ADANIPOWER)$/i.test(s)) {
    return "Power & grids (India)";
  }
  if (/^(BHARTIARTL|IDEA|INDUSTOWER)$/i.test(s)) {
    return "Telecom (India)";
  }
  if (/^(DLF|OBEROIRLTY|GODREJPROP|PRESTIGE|LODHA|BRIGADE|SOBHA)$/i.test(s)) {
    return "Real estate (India)";
  }
  if (/^(UPL|PIIND|SUMICHEM|NAVINFLUOR)$/i.test(s)) {
    return "Agri & chemicals (India)";
  }
  if (/^(HDFCLIFE|ICICIPRULI|SBILIFE)$/i.test(s)) {
    return "Insurance (India)";
  }
  if (/^NIFTY/i.test(s) || /^CNX/i.test(s) || /^INDIAVIX$/i.test(s)) {
    return "Index & benchmarks (India)";
  }
  return null;
}

/**
 * High-level “strategy sector” for portfolio lens (keyword + name heuristics — not GICS).
 * Order: first match wins. Tuned for common T212 / global listings.
 */
function pfStrategySector(r) {
  const symRaw = String(r.sym || "").trim();
  const symU = symRaw.toUpperCase().replace(/^NSE:/i, "").replace(/^BSE:/i, "");
  const t = `${String(r.nm || "")} ${symRaw}`.toUpperCase();

  const bucket = pfInstrumentBucket(r);
  if (/Funds|ETF|ETN|INDEX|MUTUAL/i.test(bucket)) {
    return "Funds, ETFs & indexes";
  }

  const india =
    String(r.ccy || "").toUpperCase() === "INR" || /\b(NSE|BSE)\b/i.test(String(r.ex || ""));
  if (india) {
    const hi = pfStrategyIndianHint(symU, t);
    if (hi) return hi;
  }
  if (
    /\b(BITCOIN|ETHEREUM|CRYPTO|BLOCKCHAIN|COINBASE|IBIT|GBTC|ETHE|GRAYSCALE|SOLANA|DEFI)\b/.test(t) ||
    /\b(BTC|ETH|WBTC)\b/.test(t)
  ) {
    return "Crypto & digital assets";
  }
  if (
    /\b(PHARMA|BIOTECH|HEALTH ?CARE|THERAPEUT|VACCIN|GENENTECH|NOVO|PFIZER|ROCHE|NOVARTIS|MERCK|JOHNSON|ABBVIE|ABBV\b|GSK\b|SANOFI|REGENERON|AMGEN|BIO\b|DRUG|LILLY\b|ASTRAZENECA|MODERNA)\b/.test(t)
  ) {
    return "Medical & pharmaceuticals";
  }
  if (
    /\b(SOFTWARE|SEMICONDUCT|SEMICONDUCTOR|TSMC|NVIDIA|INTEL|MICRON|ASML|AMD\b|CLOUD|CYBER|SAAS|DATACENTER|DATA CENTER|APPLE|MICROSOFT|GOOGLE|ALPHABET|AMAZON|ORACLE|SAP\b|SALESFORCE|SERVICENOW|ADOBE|CRM\b|NOW\b|SNOWFLAKE|PALANTIR|SERVICES)\b/.test(t) ||
    /\b(MSFT|AAPL|NVDA|TSM|AVGO|CRM|NOW|SNOW|META|GOOGL|GOOG)\b/.test(t)
  ) {
    return "Information technology";
  }
  if (
    /\b(NESTLE|NESTLÉ|COCA|PEPSI|STARBUCK|MCDONALD|KEURIG|RESTAURANT|FOOD|BEVERAGE|DANONE|KRAFT|HEINZ|UNILEVER|CONSUMER STAPLES|GROCERY|BREW|SPIRITS)\b/.test(t)
  ) {
    return "Consumer & gastronomy";
  }
  if (/\b(GOLD|SILVER|MINING|COPPER|STEEL|ALUMIN|METALS|ORE|LITHIUM)\b/.test(t) || /\b(BHP|RIO\b|VALE|FCX|SCCO)\b/.test(t)) {
    return "Metals & materials";
  }
  if (/\b(OIL|PETRO|EXXON|CHEVRON|SHELL|TOTALENERGIES|CONOCOPHILLIPS|ENI\b|SCHLUMBERGER|SLB\b|REPSOL|BP\b|ENERGY)\b/.test(t)) {
    return "Energy";
  }
  if (
    /\b(BANK|JPMORGAN|JPM\b|GOLDMAN|MORGAN STANLEY|BLACKROCK|SCHWAB|INSURANCE|BERKSHIRE|ALLIANZ|AIG\b|CITIGROUP|HSBC|BARCLAYS|UBS\b|CREDIT SUISSE)\b/.test(t)
  ) {
    return "Financial services";
  }
  if (
    /\b(TESLA|TSLA|RIVIAN|LUCID|FORD|TOYOTA|VOLKSWAGEN|STELLANTIS|AUTO|VEHICLE|AEROSPACE|BOEING|AIRBUS|CATERPILLAR|DEERE|SIEMENS|INDUSTRIAL|MANUFACTUR|3M\b|GE\b|GENERAL ELECTRIC)\b/.test(t)
  ) {
    return "Manufacturing & mobility";
  }
  return "Other / unclassified";
}

/**
 * @param {false | number} [mergeSmall] — `false` keeps every sector slice for tables; default `0.02` merges tiny pie slices into “Other”.
 */
function pfAggregate(rows, labelFn, mergeSmall) {
  /** @type {Map<string, number>} */
  const m = new Map();
  for (const r of rows) {
    const w = pfRowWeight(r);
    if (w <= 0) continue;
    const lab = labelFn(r);
    const k = lab || "—";
    m.set(k, (m.get(k) || 0) + w);
  }
  const arr = [...m.entries()]
    .map(([label, value]) => ({ label, value }))
    .sort((a, b) => b.value - a.value);
  if (mergeSmall === false) return arr;
  const minShare = typeof mergeSmall === "number" && Number.isFinite(mergeSmall) ? mergeSmall : 0.02;
  return pfMergeSmallSlices(arr, minShare);
}

/** Merge slices below `minShare` of total into "Other". */
function pfMergeSmallSlices(slices, minShare) {
  const tot = slices.reduce((s, x) => s + x.value, 0);
  if (tot <= 0 || slices.length <= 1) return slices;
  const min = tot * minShare;
  const big = [];
  let other = 0;
  for (const sl of slices) {
    if (sl.value >= min) big.push(sl);
    else other += sl.value;
  }
  if (other > 0) big.push({ label: "Other", value: other });
  return big.sort((a, b) => b.value - a.value);
}

/** (Optional) rule-based text — turned off to keep the portfolio view uncluttered. */
function buildPfIntelligenceHtml(/** @type {unknown[]} */ _rows, /** @type {string} */ _broker) {
  return "";
}

function readThemeCssVar(name, fallback) {
  try {
    const v = getComputedStyle(document.documentElement).getPropertyValue(name).trim();
    return v || fallback;
  } catch {
    return fallback;
  }
}

function drawPfPieCanvas(canvas, slices) {
  const ctx = canvas.getContext("2d");
  if (!ctx) return;
  const dpr = Math.min(2, window.devicePixelRatio || 1);
  const rawW = Number(canvas.getAttribute("width")) || 120;
  const css = Math.min(200, Math.max(96, rawW));
  canvas.width = Math.floor(css * dpr);
  canvas.height = Math.floor(css * dpr);
  canvas.style.width = `${css}px`;
  canvas.style.height = `${css}px`;
  ctx.setTransform(1, 0, 0, 1, 0, 0);
  ctx.scale(dpr, dpr);
  ctx.clearRect(0, 0, css, css);
  const cx = css / 2;
  const cy = css / 2;
  const R = css * 0.37;
  const r0 = css * 0.21;
  const holeRgb = readThemeCssVar("--card", "#1e293b");
  const edgeRgb = readThemeCssVar("--bd", "rgba(148,163,184,0.35)");
  const tot = slices.reduce((s, x) => s + x.value, 0);
  if (tot <= 0) {
    ctx.fillStyle = readThemeCssVar("--m", "#888");
    ctx.font = "12px system-ui,sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("No values", cx, cy);
    return;
  }
  let ang = -Math.PI / 2;
  slices.forEach((sl, i) => {
    const a = (sl.value / tot) * Math.PI * 2;
    ctx.beginPath();
    ctx.arc(cx, cy, R, ang, ang + a, false);
    ctx.arc(cx, cy, r0, ang + a, ang, true);
    ctx.closePath();
    ctx.fillStyle = PF_PIE_PAL[i % PF_PIE_PAL.length];
    ctx.fill();
    ctx.strokeStyle = edgeRgb;
    ctx.lineWidth = 1.25;
    ctx.stroke();
    ang += a;
  });
  ctx.beginPath();
  ctx.arc(cx, cy, r0 - 0.5, 0, Math.PI * 2);
  ctx.fillStyle = holeRgb;
  ctx.fill();
  ctx.strokeStyle = edgeRgb;
  ctx.lineWidth = 1;
  ctx.stroke();
}

/** Full sector list with % and optional single-currency value column. */
function pfSectorBreakdownHtml(rows, detailSlices) {
  const tot = detailSlices.reduce((s, x) => s + x.value, 0);
  if (tot <= 0) {
    return `<details class="pfMetaDetails pfSectorDetails mt"><summary class="sml muted">Strategy sector weights (detail)</summary><p class="sml muted mt">No sector weights yet — use <strong>Refresh prices</strong> and ensure rows have qty and last (or avg).</p></details>`;
  }
  const ccy = pfPrimaryCcy(rows);
  const ccyD = ccy ? formatCcyLabel(ccy) : "";
  const valTh = ccyD ? `Value (${esc(ccyD)})` : "Value";
  const rowsHtml = detailSlices
    .map((sl) => {
      const pct = (sl.value / tot) * 100;
      const valCell = ccy ? `<td class="pfNum">${esc(fmtMoney(ccy, sl.value))}</td>` : `<td class="muted sml">—</td>`;
      return `<tr><td>${esc(sl.label)}</td>${valCell}<td class="pfNum">${esc(fmtN(pct, 1))}%</td></tr>`;
    })
    .join("");
  return `<details class="pfMetaDetails pfSectorDetails mt">
    <summary class="sml muted">Strategy sector weights (detail)</summary>
    <div class="pfSectorTblWrap mt">
      <table class="pfSectorTbl" role="table" aria-label="Strategy sector breakdown">
        <thead><tr><th>Sector</th><th class="pfNum">${valTh}</th><th class="pfNum">Weight</th></tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      <p class="sml muted mt">Same weights as the pie (qty × last, else avg). Sectors mix <strong>India symbol rules</strong> (INR / NSE / BSE) with global keywords — not exchange GICS.</p>
    </div>
  </details>`;
}

function pfPieLegendHtml(slices) {
  const tot = slices.reduce((s, x) => s + x.value, 0);
  if (tot <= 0) return `<p class="muted sml">No data</p>`;
  return `<ul class="pfPieLeg" role="list">${slices
    .map((sl, i) => {
      const pct = (sl.value / tot) * 100;
      const p = pct >= 10 ? pct.toFixed(0) : pct.toFixed(1);
      const col = PF_PIE_PAL[i % PF_PIE_PAL.length];
      return `<li><span class="pfSwatch" style="background:${col}" aria-hidden="true"></span><span class="pfPieLab">${esc(sl.label)}</span><span class="pfPiePct">${p}%</span></li>`;
    })
    .join("")}</ul>`;
}

function renderPfCharts(rows) {
  const mount = $("pfCharts");
  if (!mount) return;
  if (!rows.length) {
    mount.hidden = true;
    mount.innerHTML = "";
    return;
  }
  mount.hidden = false;
  const cur = pfAggregate(rows, (r) => {
    const c = normalizeCcyForFx(r.ccy);
    return c || "—";
  });
  const ctry = pfAggregate(rows, pfInferCountry);
  const kind = pfAggregate(rows, pfInstrumentBucket);
  const sector = pfAggregate(rows, pfStrategySector);
  const brain = buildPfIntelligenceHtml(rows, getActiveBroker());
  mount.innerHTML = `
    <div class="card2 pfAllocCard">
      <div class="h3">Allocation</div>
      <div class="pfPieGrid">
        <figure class="pfPieCard">
          <figcaption class="pfCap">Currency</figcaption>
          <canvas class="pfPie" width="120" height="120" data-which="ccy" aria-label="Currency allocation"></canvas>
          ${pfPieLegendHtml(cur)}
        </figure>
        <figure class="pfPieCard">
          <figcaption class="pfCap">Country / region</figcaption>
          <canvas class="pfPie" width="120" height="120" data-which="ctry" aria-label="Country allocation"></canvas>
          ${pfPieLegendHtml(ctry)}
        </figure>
        <figure class="pfPieCard">
          <figcaption class="pfCap">Instrument type</figcaption>
          <canvas class="pfPie" width="120" height="120" data-which="kind" aria-label="Instrument type allocation"></canvas>
          ${pfPieLegendHtml(kind)}
        </figure>
        <figure class="pfPieCard">
          <figcaption class="pfCap">Strategy sector</figcaption>
          <canvas class="pfPie" width="120" height="120" data-which="sector" aria-label="Strategy sector allocation"></canvas>
          ${pfPieLegendHtml(sector)}
        </figure>
      </div>
    </div>
    ${brain ? `<div class="card2 mt pfBrainCard">${brain}</div>` : ""}`;
  const cans = mount.querySelectorAll("canvas.pfPie");
  const sets = [cur, ctry, kind, sector];
  cans.forEach((c, i) => {
    if (c instanceof HTMLCanvasElement && Array.isArray(sets[i])) drawPfPieCanvas(c, sets[i]);
  });
}

function renderWl() {
  const el = $("wlTbl");
  if (!el) return;
  const rows = loadWl().rows;
  if (!rows.length) {
    el.innerHTML = `<p class="muted">No symbols yet. Open a listing from <a class="backLink" href="#/search">Search</a> and use <strong>Add to watchlist</strong> on the instrument page.</p>`;
    return;
  }
  const showName = rows.some((r) => (r.nm || "").trim());
  const thName = showName ? "<th>Name</th>" : "";
  const body = rows
    .map((r) => {
      const href = instrumentHref(r.sym, r.ex, r.nm || "", r.ccy);
      const nmCell = showName ? `<td class="sml">${esc(r.nm || "")}</td>` : "";
      return `<tr>
        <td><strong><a class="pfSymLink" href="${href}">${esc(r.sym)}</a></strong></td>
        ${nmCell}
        <td>${esc(r.ex || "—")}</td>
        <td>${esc(r.ccy || "—")}</td>
        <td><button type="button" class="btn ghost smlBtn" data-wlrm="1" data-sym="${esc(r.sym)}" data-ex="${esc(r.ex || "")}">Remove</button></td>
      </tr>`;
    })
    .join("");
  el.innerHTML = `<table><thead><tr>
    <th>Sym</th>${thName}<th>Ex</th><th>Ccy</th><th></th>
  </tr></thead><tbody>${body}</tbody></table>`;
  el.querySelectorAll("button[data-wlrm]").forEach((btn) => {
    if (!(btn instanceof HTMLButtonElement)) return;
    btn.addEventListener("click", () => {
      const s = btn.dataset.sym || "";
      const x = btn.dataset.ex || "";
      if (!requireDangerPin(`remove ${s || "symbol"} from the watchlist`)) return;
      removeWlRow(s, x);
      renderWl();
      status("Removed from watchlist");
    });
  });
}

function exportPfCsv() {
  const b = getActiveBroker();
  const bundle = loadPfBundle();
  const escC = (x) => `"${String(x ?? "").replace(/"/g, '""')}"`;
  let rows = bundle.brokers[b].rows;
  let slug = b;
  if (b === PF_INSURANCE) {
    rows = insuranceRowsForCompany(bundle, getPfInsuranceCompany());
    slug = `${b}-${getPfInsuranceCompany()}`;
  }
  if (!rows.length) {
    status("Nothing to export");
    return;
  }
  const lines = [];
  if (b === PF_INSURANCE) {
    lines.push(
      [
        "insCompany",
        "policyName",
        "policyNo",
        "purchaseDate",
        "valueAtPurchase",
        "growthPct",
        "currentValue",
        "currency",
        "paymentsJson",
      ].join(","),
    );
    for (const r of rows) {
      if (!isInsuranceAltRow(r)) continue;
      const pay = JSON.stringify(Array.isArray(r.payments) ? r.payments : []);
      lines.push(
        [
          escC(r.insCompany),
          escC(r.policyName),
          escC(r.policyNo),
          escC(r.purchaseDate),
          escC(fmtN(num(r.valueAtPurchase), 4)),
          escC(fmtN(num(r.growthPct), 4)),
          escC(fmtN(num(r.currentValue), 4)),
          escC(r.ccy),
          escC(pay),
        ].join(","),
      );
    }
  } else if (b === PF_FIXED_DEPOSIT) {
    lines.push(
      ["fdBank", "fdCountry", "fdName", "fdRef", "openDate", "principal", "ratePct", "currentValue", "maturityDate", "currency"].join(","),
    );
    for (const r of rows) {
      if (!isFdAltRow(r)) continue;
      lines.push(
        [
          escC(r.fdBank),
          escC(r.fdCountry),
          escC(r.fdName),
          escC(r.fdRef),
          escC(r.openDate),
          escC(fmtN(num(r.principal), 4)),
          escC(fmtN(num(r.ratePct), 4)),
          escC(fmtN(num(r.currentValue), 4)),
          escC(r.maturityDate),
          escC(r.ccy),
        ].join(","),
      );
    }
  } else {
    const head = ["symbol", "name", "exchange", "currency", "qty", "avg_buy", "last", "market_value", "unrealized_pl"];
    lines.push(head.join(","));
    for (const r of rows) {
      const qty = num(r.qty);
      const avg = num(r.avg);
      const last = num(r.last);
      const val = qty * last;
      const pl = val - qty * avg;
      lines.push(
        [
          escC(r.sym),
          escC(r.nm),
          escC(r.ex),
          escC(r.ccy),
          escC(fmtN(qty, 6)),
          escC(fmtN(avg, 6)),
          escC(fmtN(last, 6)),
          escC(fmtN(val, 4)),
          escC(fmtN(pl, 4)),
        ].join(","),
      );
    }
  }
  if (lines.length < 2 && (b === PF_INSURANCE || b === PF_FIXED_DEPOSIT)) {
    status("Nothing to export (only legacy stock-style rows — add structured rows first)");
    return;
  }
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `portfolio-${slug}-${new Date().toISOString().slice(0, 10)}.csv`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2500);
  status("CSV exported");
}

/** Full v2 bundle — all broker ledgers — for moving data between devices / URLs. */
function exportPfBackupJson() {
  const bundle = loadPfBundle();
  const n = PF_BROKER_IDS.reduce((s, id) => s + (bundle.brokers[id]?.rows?.length || 0), 0);
  if (!n) {
    status("Nothing to back up");
    return;
  }
  const payload = {
    app: "johnsstockapp",
    exportedAt: new Date().toISOString(),
    v: bundle.v,
    brokers: bundle.brokers,
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json;charset=utf-8" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `johnsstockapp-portfolio-all-ledgers-${new Date().toISOString().slice(0, 10)}.json`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2500);
  status("JSON backup downloaded");
}

/** Replace local portfolio from JSON backup (same shape as localStorage). */
async function importPfBackupFromFile(inpEl) {
  if (!(inpEl instanceof HTMLInputElement)) return;
  const f = inpEl.files?.[0];
  inpEl.value = "";
  if (!f) return;
  let raw = "";
  try {
    raw = await f.text();
  } catch {
    status("Could not read file");
    return;
  }
  let x;
  try {
    x = JSON.parse(raw);
  } catch {
    status("Not valid JSON");
    return;
  }
  if (x?.v !== 2 || !x.brokers || typeof x.brokers !== "object") {
    status("Not a v2 portfolio backup");
    return;
  }
  if (!confirm("Replace the entire portfolio on this device with the backup? This overwrites every ledger tab.")) {
    return;
  }
  const bundle = { v: 2, brokers: {} };
  for (const id of PF_BROKER_IDS) {
    const rows = x.brokers[id]?.rows;
    bundle.brokers[id] = { rows: Array.isArray(rows) ? [...rows] : [] };
  }
  ensurePfBrokerShape(bundle.brokers);
  migratePortfolioBundleShape(bundle);
  savePfBundle(bundle);
  renderPf();
  status("Portfolio restored from JSON");
}

/** Summary strip directly under broker tabs — single-currency totals or multi-ccy guidance. */
function updatePfLedgerTotals(rows, rowObjs, multi, oneCcy, codes, tVal, tPl) {
  const el = $("pfLedgerTotals");
  if (!el) return;
  if (!rows.length) {
    el.hidden = true;
    el.innerHTML = "";
    return;
  }
  el.hidden = false;
  const b = getActiveBroker();
  const lab = PF_BROKER_LABEL[b];
  if (!multi && codes.size > 0) {
    let tCost = 0;
    for (const o of rowObjs) tCost += o.cost;
    el.innerHTML = `<div class="pfTotalsStrip" role="region" aria-label="Active ledger totals">
      <div class="pfTotalsStripInner">
        <span class="pfTotalsLedger muted sml">${esc(lab)}</span>
        <div class="pfTotalsGrid">
          <div>
            <div class="muted sml">Market value</div>
            <div class="pfTotalsBig">${esc(fmtMoney(oneCcy, tVal))}</div>
          </div>
          <div>
            <div class="muted sml">Cost basis</div>
            <div class="pfTotalsMid">${esc(fmtMoney(oneCcy, tCost))}</div>
          </div>
          <div>
            <div class="muted sml">Unrealized P/L</div>
            <div class="pfTotalsMid ${tPl >= 0 ? "plp" : "pln"}"><strong>${esc(fmtMoney(oneCcy, tPl))}</strong></div>
          </div>
        </div>
      </div>
    </div>`;
    return;
  }
  if (multi) {
    el.innerHTML = `<div class="pfTotalsStrip pfTotalsStripMulti" role="region" aria-label="Active ledger totals">
    <div class="pfTotalsStripInner">
      <span class="pfTotalsLedger muted sml">${esc(lab)} — multiple currencies</span>
    </div>
  </div>`;
    return;
  }
  el.innerHTML = `<div class="pfTotalsStrip pfTotalsStripMulti" role="region" aria-label="Active ledger totals">
    <div class="pfTotalsStripInner">
      <span class="pfTotalsLedger muted sml">${esc(lab)}</span>
    </div>
  </div>`;
}

function paintPfSubLedgerBar() {
  const bar = $("pfSubLedgerBar");
  if (!bar) return;
  const main = getMainPortfolioTab();
  if (main === PF_MAIN_TAB_MF) {
    bar.hidden = false;
    bar.setAttribute("aria-label", "Mutual fund source");
    const cur = getActiveBroker();
    bar.innerHTML = [
      { id: PF_MF_COIN, lab: "Zerodha Coin MF" },
      { id: PF_MF_KUVERA, lab: "Kuvera MF" },
    ]
      .map(
        ({ id, lab }) =>
          `<button type="button" class="brokerTab pfSubTab" role="tab" data-mf-sub="${esc(id)}" aria-selected="${id === cur ? "true" : "false"}">${esc(lab)}</button>`,
      )
      .join("");
    return;
  }
  if (main === PF_INSURANCE) {
    bar.hidden = false;
    bar.setAttribute("aria-label", "Insurance provider");
    const cur = getPfInsuranceCompany();
    bar.innerHTML = PF_INS_CO_IDS.map(
      (id) =>
        `<button type="button" class="brokerTab pfSubTab insCoTab" role="tab" data-insco="${esc(id)}" aria-selected="${id === cur ? "true" : "false"}">${esc(PF_INS_CO_LABEL[id] || id)}</button>`,
    ).join("");
    return;
  }
  bar.hidden = true;
  bar.innerHTML = "";
}

function paintPfAddFieldsMount() {
  const mount = $("pfAddFieldsMount");
  if (!mount) return;
  const b = getActiveBroker();
  if (b === PF_INSURANCE) {
    const coLab = esc(PF_INS_CO_LABEL[getPfInsuranceCompany()] || "");
    mount.innerHTML = `
      <label class="lbl pfAddField">Policy name
        <input class="in" id="fPolName" placeholder="e.g. Smart Wealth" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Policy number
        <input class="in" id="fPolNo" placeholder="—" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Date of purchase
        <input class="in" id="fPolPur" type="date" />
      </label>
      <label class="lbl pfAddField">Value at purchase
        <input class="in" id="fPolV0" placeholder="Initial premium / corpus" inputmode="decimal" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Avg growth % p.a. <span class="muted sml">(opt.)</span>
        <input class="in" id="fPolGr" placeholder="e.g. 7" inputmode="decimal" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Current possible value
        <input class="in" id="fPolCur" placeholder="Surrender / fund value" inputmode="decimal" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Currency
        <input class="in" id="fPolCcy" placeholder="INR, EUR…" autocomplete="off" />
      </label>`;
    return;
  }
  if (b === PF_CRYPTO) {
    mount.innerHTML = `
      <label class="lbl pfAddField">Symbol
        <input class="in" id="fSym" placeholder="BTC, ETH" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Name <span class="muted sml">(opt.)</span>
        <input class="in" id="fNm" placeholder="Bitcoin" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Exchange <span class="muted sml">(opt.)</span>
        <input class="in" id="fEx" placeholder="leave empty or CRYPTO" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Currency
        <input class="in" id="fCcy" placeholder="USD" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Qty / units
        <input class="in" id="fQty" placeholder="0.5" inputmode="decimal" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Avg buy
        <input class="in" id="fAvg" placeholder="Average in row currency" inputmode="decimal" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Last (live) <span class="muted sml">(opt.)</span>
        <input class="in" id="fLast" placeholder="Filled on Refresh" inputmode="decimal" autocomplete="off" />
      </label>`;
    return;
  }
  if (b === PF_FIXED_DEPOSIT) {
    mount.innerHTML = `
      <label class="lbl pfAddField">Bank / institution
        <input class="in" id="fFdBank" placeholder="SBI, HDFC…" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Deposit name
        <input class="in" id="fFdName" placeholder="e.g. 12-mo FD" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Reference / receipt # <span class="muted sml">(opt.)</span>
        <input class="in" id="fFdRef" placeholder="—" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Open date
        <input class="in" id="fFdOpen" type="date" />
      </label>
      <label class="lbl pfAddField">Principal
        <input class="in" id="fFdPrin" inputmode="decimal" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Rate % p.a. <span class="muted sml">(opt.)</span>
        <input class="in" id="fFdRate" inputmode="decimal" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Current / maturity value
        <input class="in" id="fFdCur" inputmode="decimal" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Maturity date <span class="muted sml">(opt.)</span>
        <input class="in" id="fFdMat" type="date" />
      </label>
      <label class="lbl pfAddField">Currency
        <input class="in" id="fFdCcy" placeholder="INR…" autocomplete="off" />
      </label>
      <label class="lbl pfAddField">Country / region <span class="muted sml">(opt.)</span>
        <input class="in" id="fFdCtry" placeholder="India, Ireland…" autocomplete="off" />
      </label>`;
    return;
  }
  mount.innerHTML = `
    <label class="lbl pfAddField">Symbol / ISIN
      <input class="in" id="fSym" placeholder="TSLA, INF209K01Z…" autocomplete="off" />
    </label>
    <label class="lbl pfAddField">Name <span class="muted sml">(opt.)</span>
      <input class="in" id="fNm" placeholder="—" autocomplete="off" />
    </label>
    <label class="lbl pfAddField">Exchange <span class="muted sml">(opt.)</span>
      <input class="in" id="fEx" placeholder="NSE, NASDAQ…" autocomplete="off" />
    </label>
    <label class="lbl pfAddField">Currency
      <input class="in" id="fCcy" placeholder="EUR, INR…" autocomplete="off" />
    </label>
    <label class="lbl pfAddField">Qty / units
      <input class="in" id="fQty" placeholder="Units" inputmode="decimal" autocomplete="off" />
    </label>
    <label class="lbl pfAddField">Avg buy / avg NAV
      <input class="in" id="fAvg" placeholder="Avg cost" inputmode="decimal" autocomplete="off" />
    </label>
    <label class="lbl pfAddField">Last / NAV
      <input class="in" id="fLast" placeholder="Current" inputmode="decimal" autocomplete="off" />
    </label>`;
}

function pfClearTableMountState() {
  const mt = $("pfT212EurMount");
  if (mt) mt.innerHTML = "";
  const totStrip = $("pfLedgerTotals");
  if (totStrip) {
    totStrip.hidden = true;
    totStrip.innerHTML = "";
  }
}

function handlePfTableClick(ev) {
  const t = ev.target;
  if (!(t instanceof HTMLElement)) return;
  const hdr = t.closest("button.pf-hdr-sort");
  if (hdr instanceof HTMLButtonElement && hdr.dataset.pfHdrKey) {
    if (t.closest("table.pfAltHoldingsTbl")) return;
    const host = ev.currentTarget;
    if (!(host instanceof HTMLElement) || host.id !== "tbl") return;
    const b = getActiveBroker();
    if (b === PF_INSURANCE || b === PF_FIXED_DEPOSIT) return;
    const st0 = pfHoldingSortRead(b);
    const k = String(hdr.dataset.pfHdrKey || "sym");
    const next = st0.k === k ? { k, d: -st0.d } : { k, d: pfHoldingSortDefaultDir(k) };
    pfHoldingSortWrite(b, next);
    renderPf();
    return;
  }
  const rm = t.closest("[data-pf-rm]");
  if (rm instanceof HTMLElement && rm.dataset.pfRm) {
    const rid = rm.dataset.pfRm;
    const b = getActiveBroker();
    if (!requireDangerPin(`remove this ${PF_BROKER_LABEL[b]} row`)) return;
    const bundle = loadPfBundle();
    const rows = bundle.brokers[b].rows;
    const i = rows.findIndex((r) => pfRowId(r) === rid || (!pfRowId(r) && String(r.sym) === rid));
    if (i < 0) return;
    rows.splice(i, 1);
    savePfBundle(bundle);
    renderPf();
    status("Row removed");
    return;
  }
  const pr = t.closest("[data-pf-prem]");
  if (pr instanceof HTMLElement && pr.dataset.pfPrem) {
    const rid = pr.dataset.pfPrem;
    const amtS = window.prompt("Premium / top-up amount (number only)", "");
    if (amtS === null) return;
    const amt = num(amtS);
    if (!(amt > 0)) {
      status("Enter a positive amount");
      return;
    }
    const bundle = loadPfBundle();
    const rows = bundle.brokers[PF_INSURANCE].rows;
    const row = rows.find((r) => pfRowId(r) === rid || String(r.sym || "") === rid);
    if (!row || !isInsuranceAltRow(row)) return;
    if (!Array.isArray(row.payments)) row.payments = [];
    const today = new Date().toISOString().slice(0, 10);
    row.payments.push({ date: today, amount: amt });
    savePfBundle(bundle);
    renderPf();
    status("Payment logged");
  }
}

function renderPfInsuranceTable(el) {
  const bundle = loadPfBundle();
  const co = getPfInsuranceCompany();
  const rows = insuranceRowsForCompany(bundle, co).sort((a, b) =>
    String(a.policyName || a.sym || "").localeCompare(String(b.policyName || b.sym || ""), undefined, {
      sensitivity: "base",
    }),
  );
  if (!rows.length) {
    el.innerHTML = `<p class="muted">No policies for <strong>${esc(PF_INS_CO_LABEL[co])}</strong> yet.</p>`;
    pfClearTableMountState();
    renderPfCharts([]);
    void refreshPfCombinedEur();
    return;
  }
  const rowObjs = rows.map((r) => {
    if (isInsuranceAltRow(r)) {
      const invested = insuranceInvestedTotal(r);
      const val = num(r.currentValue);
      const pl = val - invested;
      const ccy = normalizeCcyForFx(String(r.ccy || "INR").toUpperCase());
      return { r, invested, val, pl, ccy };
    }
    const qty = num(r.qty);
    const avg = num(r.avg);
    const last = num(r.last);
    return { r, invested: qty * avg, val: qty * last, pl: qty * (last - avg), ccy: normalizeCcyForFx(String(r.ccy || "").toUpperCase()) };
  });
  const codes = new Set(rowObjs.map((o) => o.ccy).filter(Boolean));
  const multi = codes.size > 1;
  let tVal = 0;
  let tPl = 0;
  let tInv = 0;
  const oneCcy = !multi && codes.size ? Array.from(codes)[0] : "";
  for (const o of rowObjs) {
    if (!multi && oneCcy) {
      tVal += o.val;
      tPl += o.pl;
      tInv += o.invested;
    }
  }
  const fakeRows = rowObjs.map((o) => ({ ccy: o.ccy, qty: 1, avg: o.invested, last: o.val }));
  const fakeObjs = rowObjs.map((o) => ({
    r: o.r,
    qty: 1,
    avg: o.invested,
    last: o.val,
    cost: o.invested,
    val: o.val,
    pl: o.pl,
  }));
  updatePfLedgerTotals(
    fakeRows,
    fakeObjs,
    multi,
    oneCcy || rowObjs[0]?.ccy || "",
    codes,
    !multi ? tVal : 0,
    !multi ? tPl : 0,
  );
  const body = rowObjs
    .map((o) => {
      const { r, invested, val, pl, ccy } = o;
      const cls = pl >= 0 ? "plp" : "pln";
      const rid = pfRowId(r) || String(r.sym || "");
      if (isInsuranceAltRow(r)) {
        const pc = Array.isArray(r.payments) ? r.payments.length : 0;
        const ps = sumInsurancePayments(r);
        return `<tr>
          <td class="sml"><strong>${esc(r.policyName || "—")}</strong></td>
          <td class="sml">${esc(r.policyNo || "—")}</td>
          <td class="sml">${esc(r.purchaseDate || "—")}</td>
          <td class="pfNum">${esc(fmtMoney(ccy, num(r.valueAtPurchase)))}</td>
          <td class="pfNum">${esc(fmtMoney(ccy, ps))}</td>
          <td class="pfNum">${esc(fmtMoney(ccy, invested))}</td>
          <td class="pfNum">${esc(fmtN(num(r.growthPct), 2))}%</td>
          <td class="pfNum">${esc(fmtMoney(ccy, val))}</td>
          <td class="pfNum pfPlCol ${cls}">${esc(fmtMoney(ccy, pl))}</td>
          <td class="sml">${esc(formatCcyLabel(ccy))}</td>
          <td class="pfActCol"><button type="button" class="btn ghost smlBtn" data-pf-prem="${esc(rid)}">Log premium</button>
            <button type="button" class="btn ghost smlBtn" data-pf-rm="${esc(rid)}">Remove</button>
            <span class="muted sml">${pc} payment(s)</span></td>
        </tr>`;
      }
      const href = instrumentHref(r.sym, r.ex, r.nm || "", r.ccy, { fromPf: true });
      return `<tr>
        <td class="sml" colspan="3"><span class="muted">Legacy row</span> <a class="pfSymLink" href="${href}"><strong>${esc(r.sym)}</strong></a></td>
        <td class="pfNum">—</td><td class="pfNum">—</td>
        <td class="pfNum">${esc(fmtMoney(ccy, invested))}</td>
        <td class="pfNum">—</td>
        <td class="pfNum">${esc(fmtMoney(ccy, val))}</td>
        <td class="pfNum pfPlCol ${cls}">${esc(fmtMoney(ccy, pl))}</td>
        <td>${esc(formatCcyLabel(ccy))}</td>
        <td class="pfActCol"><button type="button" class="btn ghost smlBtn" data-pf-rm="${esc(rid)}">Remove</button></td>
      </tr>`;
    })
    .join("");
  const foot = !multi && oneCcy
    ? `<tr class="tot"><td colspan="6"><strong>Total</strong></td><td class="pfNum"><strong>${esc(fmtMoney(oneCcy, tVal))}</strong></td><td class="pfNum pfPlCol ${tPl >= 0 ? "plp" : "pln"}"><strong>${esc(fmtMoney(oneCcy, tPl))}</strong></td><td></td><td></td></tr>`
    : `<tr class="tot"><td colspan="11" class="muted">Several currencies in this provider — row amounts stay native. See the <strong>By ledger (€)</strong> block at the <strong>top</strong> of the page.</td></tr>`;
  el.innerHTML = `<div class="pfTableWrap" role="region" aria-label="Insurance policies"><table class="pfHoldingsTbl pfAltHoldingsTbl"><thead><tr>
    <th>Policy</th><th>Policy #</th><th>Purchased</th><th class="pfNum">At purchase</th><th class="pfNum">Premiums paid</th><th class="pfNum">Invested total</th><th class="pfNum">Growth %</th><th class="pfNum">Current value</th><th class="pfNum pfPlCol">P/L</th><th>Ccy</th><th></th>
  </tr></thead><tbody>${body}${foot}</tbody></table></div>`;
  el.onclick = (e) => handlePfTableClick(e);
  pfClearTableMountState();
  renderPfCharts([]);
  void refreshPfCombinedEur();
}

function renderPfFdTable(el) {
  const bundle = loadPfBundle();
  const rows = [...bundle.brokers[PF_FIXED_DEPOSIT].rows].sort((a, b) =>
    String(a.fdName || a.sym || "").localeCompare(String(b.fdName || b.sym || ""), undefined, { sensitivity: "base" }),
  );
  if (!rows.length) {
    el.innerHTML = `<p class="muted">No fixed deposits yet.</p>`;
    pfClearTableMountState();
    renderPfCharts([]);
    void refreshPfCombinedEur();
    return;
  }
  const rowObjs = rows.map((r) => {
    if (isFdAltRow(r)) {
      const inv = num(r.principal);
      const val = num(r.currentValue);
      const pl = val - inv;
      const ccy = normalizeCcyForFx(String(r.ccy || "INR").toUpperCase());
      return { r, inv, val, pl, ccy };
    }
    const qty = num(r.qty);
    const avg = num(r.avg);
    const last = num(r.last);
    return { r, inv: qty * avg, val: qty * last, pl: qty * (last - avg), ccy: normalizeCcyForFx(String(r.ccy || "").toUpperCase()) };
  });
  const codes = new Set(rowObjs.map((o) => o.ccy).filter(Boolean));
  const multi = codes.size > 1;
  let tVal = 0;
  let tPl = 0;
  const oneCcy = !multi && codes.size ? Array.from(codes)[0] : "";
  for (const o of rowObjs) {
    if (!multi && oneCcy) {
      tVal += o.val;
      tPl += o.pl;
    }
  }
  const fakeRows = rowObjs.map((o) => ({ ccy: o.ccy, qty: 1, avg: o.inv, last: o.val }));
  const fakeObjs = rowObjs.map((o) => ({
    r: o.r,
    qty: 1,
    avg: o.inv,
    last: o.val,
    cost: o.inv,
    val: o.val,
    pl: o.pl,
  }));
  updatePfLedgerTotals(
    fakeRows,
    fakeObjs,
    multi,
    oneCcy || rowObjs[0]?.ccy || "",
    codes,
    !multi ? tVal : 0,
    !multi ? tPl : 0,
  );
  const body = rowObjs
    .map((o) => {
      const { r, inv, val, pl, ccy } = o;
      const cls = pl >= 0 ? "plp" : "pln";
      const rid = pfRowId(r) || String(r.sym || "");
      if (isFdAltRow(r)) {
        return `<tr>
          <td class="sml">${esc(r.fdBank || "—")}</td>
          <td class="sml">${esc(r.fdCountry || "—")}</td>
          <td class="sml"><strong>${esc(r.fdName || "—")}</strong></td>
          <td class="sml">${esc(r.fdRef || "—")}</td>
          <td class="sml">${esc(r.openDate || "—")}</td>
          <td class="pfNum">${esc(fmtMoney(ccy, num(r.principal)))}</td>
          <td class="pfNum">${esc(fmtN(num(r.ratePct), 2))}%</td>
          <td class="pfNum">${esc(fmtMoney(ccy, val))}</td>
          <td class="sml">${esc(r.maturityDate || "—")}</td>
          <td>${esc(formatCcyLabel(ccy))}</td>
          <td class="pfNum pfPlCol ${cls}">${esc(fmtMoney(ccy, pl))}</td>
          <td class="pfActCol"><button type="button" class="btn ghost smlBtn" data-pf-rm="${esc(rid)}">Remove</button></td>
        </tr>`;
      }
      const href = instrumentHref(r.sym, r.ex, r.nm || "", r.ccy, { fromPf: true });
      return `<tr>
        <td colspan="5"><span class="muted">Legacy</span> <a class="pfSymLink" href="${href}"><strong>${esc(r.sym)}</strong></a></td>
        <td class="pfNum">${esc(fmtMoney(ccy, inv))}</td><td>—</td>
        <td class="pfNum">${esc(fmtMoney(ccy, val))}</td><td>—</td><td>${esc(formatCcyLabel(ccy))}</td>
        <td class="pfNum pfPlCol ${cls}">${esc(fmtMoney(ccy, pl))}</td>
        <td class="pfActCol"><button type="button" class="btn ghost smlBtn" data-pf-rm="${esc(rid)}">Remove</button></td>
      </tr>`;
    })
    .join("");
  const foot = !multi && oneCcy
    ? `<tr class="tot"><td colspan="7"><strong>Total current value</strong></td><td class="pfNum"><strong>${esc(fmtMoney(oneCcy, tVal))}</strong></td><td colspan="2"></td><td class="pfNum pfPlCol ${tPl >= 0 ? "plp" : "pln"}"><strong>${esc(fmtMoney(oneCcy, tPl))}</strong></td><td></td></tr>`
    : `<tr class="tot"><td colspan="12" class="muted">Multiple currencies and countries — each row keeps its own currency; see the € total under the page title.</td></tr>`;
  el.innerHTML = `<div class="pfTableWrap" role="region" aria-label="Fixed deposits"><table class="pfHoldingsTbl pfAltHoldingsTbl"><thead><tr>
    <th>Bank</th><th>Country</th><th>Deposit</th><th>Ref #</th><th>Open</th><th class="pfNum">Principal</th><th class="pfNum">Rate</th><th class="pfNum">Current value</th><th>Maturity</th><th>Ccy</th><th class="pfNum pfPlCol">P/L vs principal</th><th></th>
  </tr></thead><tbody>${body}${foot}</tbody></table></div>`;
  el.onclick = (e) => handlePfTableClick(e);
  pfClearTableMountState();
  renderPfCharts([]);
  void refreshPfCombinedEur();
}

function pfHoldingSortRead(brokerId) {
  try {
    const raw = sessionStorage.getItem(PF_HOLDING_SORT_KEY);
    const o = raw ? JSON.parse(raw) : {};
    if (o && typeof o === "object" && o[brokerId] && o[brokerId].k) {
      const d = Number(o[brokerId].d);
      return { k: String(o[brokerId].k), d: d === -1 ? -1 : 1 };
    }
  } catch {
    /* ignore */
  }
  return { k: "sym", d: 1 };
}

function pfHoldingSortWrite(brokerId, st) {
  try {
    const raw = sessionStorage.getItem(PF_HOLDING_SORT_KEY);
    const o = raw ? JSON.parse(raw) : {};
    if (typeof o !== "object" || o === null) return;
    o[brokerId] = st;
    sessionStorage.setItem(PF_HOLDING_SORT_KEY, JSON.stringify(o));
  } catch {
    /* ignore */
  }
}

function pfCcySortIndex(ccy) {
  const n = normalizeCcyForFx(String(ccy || ""));
  if (!n) return 10000;
  const i = PF_CCY_SORT_ORDER.indexOf(n);
  if (i >= 0) return i;
  return 500 + n.charCodeAt(0) * 3 + (n.charCodeAt(1) || 0) % 40;
}

function pfHoldingSortDefaultDir(key) {
  return key === "sym" || key === "nm" || key === "ex" || key === "ccy" ? 1 : -1;
}

/**
 * @param {{ r: object, qty: number, avg: number, last: number, val: number, pl: number, wPct: number }} A
 * @param {{ k: string, d: number }} sst
 */
function pfHoldingRowCompare(A, B, sst, multi) {
  const k0 = sst.k;
  const d = sst.d;
  const k = multi && k0 === "w" ? "sym" : k0;
  const an = (x) => String(x ?? "").toLowerCase();
  let raw = 0;
  switch (k) {
    case "sym":
      raw = an(A.r.sym).localeCompare(an(B.r.sym), undefined, { sensitivity: "base" });
      break;
    case "nm":
      raw = an(A.r.nm || A.r.sym).localeCompare(an(B.r.nm || B.r.sym), undefined, { sensitivity: "base" });
      break;
    case "ex":
      raw = an(A.r.ex).localeCompare(an(B.r.ex), undefined, { sensitivity: "base" });
      break;
    case "ccy":
      raw = pfCcySortIndex(A.r.ccy) - pfCcySortIndex(B.r.ccy);
      break;
    case "qty":
      raw = (A.qty - B.qty) || 0;
      break;
    case "last":
      raw = (A.last - B.last) || 0;
      break;
    case "avg":
      raw = (A.avg - B.avg) || 0;
      break;
    case "val":
      raw = (A.val - B.val) || 0;
      break;
    case "w":
      raw = (A.wPct - B.wPct) || 0;
      break;
    case "pl":
      raw = (A.pl - B.pl) || 0;
      break;
    default:
      raw = 0;
  }
  if (d < 0) raw = -raw;
  if (raw !== 0) return raw;
  return an(A.r.sym).localeCompare(an(B.r.sym), undefined, { sensitivity: "base" });
}

function pfHoldingTheadBtn(key, label, isNum, st, noSort) {
  if (noSort) {
    return `<th class="${isNum ? "pfNum" : ""} pfThPlain" scope="col">${esc(label)}</th>`;
  }
  const active = st.k === key;
  const dir = st.d;
  const caret = active ? (dir > 0 ? " \u25B2" : " \u25BC") : "";
  const ariaS = active ? (dir > 0 ? ' aria-sort="ascending"' : ' aria-sort="descending"') : "";
  return `<th class="${isNum ? "pfNum" : ""} pfThSort" scope="col"${ariaS}><button type="button" class="pf-hdr-sort${isNum ? " pf-hdr-sortNum" : ""}" data-pf-hdr-key="${key}" aria-label="Sort by ${esc(label)}"${active ? ' aria-pressed="true"' : ""}>${esc(label)}<span class="pfSortInd" aria-hidden="true">${caret}</span></button></th>`;
}

/** Extra copy when a ledger has no rows (private window vs family link). */
function portfolioLedgerEmptyHintHtml() {
  let isFamilyUrl = false;
  try {
    const { sp } = parseLocationHash();
    const v = (sp.get("view") || "").trim().toLowerCase();
    const tok = familyReadTokenFromUrl(sp);
    isFamilyUrl = v === "family" && Boolean(tok);
  } catch {
    /* ignore */
  }
  if (isFamilyUrl) {
    return `<p class="sml muted pfEmptyLedgerHint mt">If you expected holdings here, confirm you’re using the <strong>exact</strong> family link (with <code>view=family</code> and <code>token=…</code>) and that the owner has published. A normal <code>#/portfolio</code> link has no data in a private / new browser.</p>`;
  }
  return `<div class="card2 mt pfEmptyLedgerHint" role="note"><p class="sml" style="margin:0;line-height:1.55;max-width:42rem">Your portfolio in this app is stored in <strong>this browser only</strong> (not on the server). A <strong>private / incognito</strong> window, or another device, starts with an <strong>empty</strong> ledger — that is why you see no totals or rows.</p><p class="sml muted mt" style="margin:0 0 0 0;line-height:1.55;max-width:42rem">For <strong>family</strong> (read-only) access to data the owner published, open the <strong>full</strong> URL they sent, which must include <code>view=family</code> and <code>token=…</code> in the address bar (not only <code>#/portfolio</code>).</p></div>`;
}

/* portfolio table (uses `loadPfBundle` + active broker) */
function renderPf() {
  const el = $("tbl");
  if (!el) return;
  el.onclick = null;
  paintPfSubLedgerBar();
  syncPortfolioTabAria();
  paintPfAddFieldsMount();
  updatePfBrokerCaption();
  const b = getActiveBroker();
  if (b === PF_INSURANCE) {
    renderPfInsuranceTable(el);
    paintPfT212SyncTimestamp();
    return;
  }
  if (b === PF_FIXED_DEPOSIT) {
    renderPfFdTable(el);
    paintPfT212SyncTimestamp();
    return;
  }
  const rowsRaw = loadPfBundle().brokers[b].rows;
  const rows = [...rowsRaw];
  if (!rows.length) {
    el.innerHTML = `<p class="muted">No rows in this ledger yet.</p>${portfolioLedgerEmptyHintHtml()}`;
    el.onclick = null;
    pfClearTableMountState();
    renderPfCharts([]);
    void refreshPfCombinedEur();
    paintPfT212SyncTimestamp();
    return;
  }
  const codes = new Set(rows.map((r) => normalizeCcyForFx(r.ccy)).filter(Boolean));
  const multi = codes.size > 1;
  const showName = rows.some((r) => (r.nm || "").trim());
  const rowObjs = rows.map((r) => {
    const qty = num(r.qty);
    const avg = num(r.avg);
    const last = num(r.last);
    const cost = qty * avg;
    const val = qty * last;
    const pl = val - cost;
    return { r, qty, avg, last, cost, val, pl, wPct: 0 };
  });
  let tVal = 0;
  let tPl = 0;
  for (const o of rowObjs) {
    if (!multi) {
      tVal += o.val;
      tPl += o.pl;
    }
  }
  for (const o of rowObjs) {
    o.wPct = !multi && tVal > 0 ? (o.val / tVal) * 100 : 0;
  }
  const sst = pfHoldingSortRead(b);
  rowObjs.sort((A, B) => pfHoldingRowCompare(A, B, sst, multi));
  const t212CsvHint =
    b === PF_T212 &&
    rows.some((r) => !r.t212Ticker && /_EQ$/i.test(String(r.sym || "").replace(/\s+/g, "")))
      ? `<div class="card2 mt" role="status"><p class="sml"><strong>CSV / manual rows:</strong> symbols and the <strong>CCY</strong> column match your import (often account USD). <strong>Sync from Trading 212</strong> replaces this ledger with live open positions: cleaned tickers, exchange, and listing currency — requires your server’s Trading 212 API keys on Render.</p></div>`
      : "";
  const thSt = sst;
  const thWeight = !multi ? pfHoldingTheadBtn("w", "Weight", true, thSt, false) : pfHoldingTheadBtn("w", "Weight", true, thSt, true);
  const body = rowObjs
    .map((o) => {
      const { r, qty, avg, last, val, pl } = o;
      const cls = pl >= 0 ? "plp" : "pln";
      const nmCell = showName ? `<td class="sml">${esc(r.nm || "")}</td>` : "";
      const href = instrumentHref(r.sym, r.ex, r.nm || "", r.ccy, { fromPf: true });
      const wCell =
        !multi && tVal > 0
          ? `<td class="sml pfNum">${esc(fmtN((val / tVal) * 100, 1))}%</td>`
          : !multi
            ? `<td class="sml pfNum">—</td>`
            : "";
      return `<tr>
        <td class="pfSymCol"><strong><a class="pfSymLink" href="${href}">${esc(r.sym)}</a></strong></td>
        ${nmCell}
        <td class="pfExCol">${esc(r.ex)}</td>
        <td>${esc(formatCcyLabel(r.ccy))}</td>
        <td class="pfNum">${esc(fmtN(qty, 4))}</td>
        <td class="pfNum">${esc(fmtMoney(r.ccy, avg))}</td>
        <td class="pfNum">${esc(fmtMoney(r.ccy, last))}</td>
        <td class="pfNum">${esc(fmtMoney(r.ccy, val))}</td>
        ${wCell}
        <td class="pfNum pfPlCol ${cls}">${esc(fmtMoney(r.ccy, pl))}</td>
      </tr>`;
    })
    .join("");
  const oneCcy = !multi && codes.size ? Array.from(codes)[0] : rows[0]?.ccy || "";
  updatePfLedgerTotals(rows, rowObjs, multi, oneCcy, codes, tVal, tPl);
  const nc = showName ? 1 : 0;
  const foot = !multi
    ? `<tr class="tot"><td colspan="${6 + nc}"><strong>Total</strong></td><td class="pfNum"><strong>${esc(fmtMoney(oneCcy, tVal))}</strong></td><td class="sml pfNum"><strong>100%</strong></td><td class="pfNum pfPlCol ${tPl >= 0 ? "plp" : "pln"}"><strong>${esc(fmtMoney(oneCcy, tPl))}</strong></td></tr>`
    : `<tr class="tot"><td colspan="${8 + nc}" class="muted">Multiple currencies in this table.</td></tr>`;
  const thName = showName ? pfHoldingTheadBtn("nm", "Name", false, thSt, false) : "";
  const colg = showName
    ? `<colgroup><col class="pfCol pfColSym" /><col class="pfCol pfColNm" /><col class="pfCol pfColEx" /><col class="pfCol pfColCcy" /><col class="pfCol pfColQty" /><col class="pfCol pfColAvg" /><col class="pfCol pfColLast" /><col class="pfCol pfColVal" />${!multi ? `<col class="pfCol pfColWt" />` : ""}<col class="pfCol pfColPl" /></colgroup>`
    : `<colgroup><col class="pfCol pfColSym" /><col class="pfCol pfColEx" /><col class="pfCol pfColCcy" /><col class="pfCol pfColQty" /><col class="pfCol pfColAvg" /><col class="pfCol pfColLast" /><col class="pfCol pfColVal" />${!multi ? `<col class="pfCol pfColWt" />` : ""}<col class="pfCol pfColPl" /></colgroup>`;
  el.innerHTML = `${t212CsvHint}<div class="pfTableWrap" role="region" aria-label="Holdings table"><table class="pfHoldingsTbl">${colg}<thead><tr>
    ${pfHoldingTheadBtn("sym", "Sym", false, thSt, false)}${thName}${pfHoldingTheadBtn("ex", "Ex", false, thSt, false)}${pfHoldingTheadBtn("ccy", "Ccy", false, thSt, false)}${pfHoldingTheadBtn("qty", "Qty", true, thSt, false)}${pfHoldingTheadBtn("avg", "Avg", true, thSt, false)}${pfHoldingTheadBtn("last", "Last", true, thSt, false)}${pfHoldingTheadBtn("val", "Value", true, thSt, false)}${thWeight}${pfHoldingTheadBtn("pl", "P/L", true, thSt, false)}
  </tr></thead><tbody>${body}${foot}</tbody></table></div>`;
  el.onclick = (e) => handlePfTableClick(e);
  const t212Mount = $("pfT212EurMount");
  if (t212Mount) {
    if ((b === PF_T212 || b === PF_CRYPTO) && rows.length) {
      t212Mount.innerHTML = `<div id="pfT212Eur" class="pfT212Eur card2 mt" role="region" aria-live="polite"><p class="muted sml">Loading EUR conversion…</p></div>`;
      const eurTitle = b === PF_CRYPTO ? "Crypto (T212) — totals in EUR" : "Trading 212 — totals in EUR";
      void refreshPfT212Euro(rows, multi, oneCcy, tVal, tPl, eurTitle);
    } else {
      t212Mount.innerHTML = "";
    }
  }
  renderPfCharts(isStockLikePfBroker(b) ? rowsRaw : []);
  void refreshPfCombinedEur();
  paintPfT212SyncTimestamp();
}

/** Fetch ECB reference table via this app’s `/api/fx-eur` (retries + safe JSON parse). */
async function fetchEurFxTable() {
  let lastErr = /** @type {unknown} */ (null);
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const r = await fetch("/api/fx-eur", { cache: "no-store" });
      const raw = await r.text();
      let j;
      try {
        j = JSON.parse(raw);
      } catch {
        lastErr = new SyntaxError("not json");
        continue;
      }
      if (!r.ok) {
        lastErr = new Error(String(j?.detail || j?.error || r.status));
        continue;
      }
      if (!j.eur_per_unit || typeof j.eur_per_unit !== "object") {
        lastErr = new Error("no eur_per_unit");
        continue;
      }
      const ep = { ...j.eur_per_unit };
      if (typeof ep.EUR !== "number" || !Number.isFinite(ep.EUR) || ep.EUR <= 0) ep.EUR = 1;
      j.eur_per_unit = ep;
      return j;
    } catch (e) {
      lastErr = e;
    }
    await new Promise((res) => setTimeout(res, 450 * (attempt + 1)));
  }
  throw lastErr instanceof Error ? lastErr : new Error(String(lastErr));
}

/** Map insurance / FD alt rows into qty×avg / qty×last so ECB EUR rollup matches the table. */
function brokerRowsForEurLedger(brokerId, rows) {
  if (brokerId === PF_INSURANCE) {
    return rows.map((r) => {
      if (isInsuranceAltRow(r)) {
        const inv = insuranceInvestedTotal(r);
        const last = num(r.currentValue);
        const ccy = normalizeCcyForFx(String(r.ccy || "INR").toUpperCase());
        const px = last > 0 ? last : inv;
        return { ccy, qty: 1, avg: inv, last: px };
      }
      return r;
    });
  }
  if (brokerId === PF_FIXED_DEPOSIT) {
    return rows.map((r) => {
      if (isFdAltRow(r)) {
        const inv = num(r.principal);
        const last = num(r.currentValue);
        const ccy = normalizeCcyForFx(String(r.ccy || "INR").toUpperCase());
        const px = last > 0 ? last : inv;
        return { ccy, qty: 1, avg: inv, last: px };
      }
      return r;
    });
  }
  return rows;
}

function portfolioEurFromRows(rows, eurPer) {
  let totE = 0;
  let costE = 0;
  const miss = new Set();
  for (const row of rows) {
    const raw = String(row.ccy || "").trim();
    const ccy = normalizeCcyForFx(raw);
    const qty = num(row.qty);
    const avg = num(row.avg);
    const last = num(row.last);
    const rate = ccy ? eurPerUnitToEur(eurPer, ccy) : undefined;
    if (rate == null) {
      if (raw) miss.add(raw);
      else if (ccy) miss.add(ccy);
      continue;
    }
    totE += qty * last * rate;
    costE += qty * avg * rate;
  }
  return { totE, costE, plE: totE - costE, miss };
}

async function refreshPfCombinedEur() {
  const box = $("pfCombinedEur");
  const grand = $("pfGrandTotalMount");
  if (!box) return;
  const bundle = loadPfBundle();
  if (pfBrokersEmpty(bundle.brokers)) {
    box.hidden = true;
    box.innerHTML = "";
    if (grand) {
      grand.hidden = true;
      grand.innerHTML = "";
    }
    return;
  }
  box.hidden = false;
  box.innerHTML = `<div class="card2 pfCombinedEurInner"><p class="muted sml">Loading combined EUR (all ledgers)…</p></div>`;
  if (grand) {
    grand.hidden = false;
    grand.innerHTML = `<div class="card2 pfGrandTotalInner"><p class="muted sml">Loading total (€)…</p></div>`;
  }
  try {
    const j = await fetchEurFxTable();
    const eurPer = j.eur_per_unit;
    const t212All = Array.isArray(bundle.brokers[PF_T212]?.rows) ? bundle.brokers[PF_T212].rows : [];
    const t212Crypto = t212All.filter(isCryptoLikeT212Row);
    const t212Stocks = t212All.filter((r) => !isCryptoLikeT212Row(r));
    const cryptoStandalone = Array.isArray(bundle.brokers[PF_CRYPTO]?.rows) ? bundle.brokers[PF_CRYPTO].rows : [];
    const cryptoCombined = [...cryptoStandalone, ...t212Crypto];
    const legs = PF_BROKER_IDS.map((id) => {
      let rawForCount;
      let eurInputRows;
      if (id === PF_T212) {
        eurInputRows = brokerRowsForEurLedger(PF_T212, t212Stocks);
        rawForCount = t212All;
      } else if (id === PF_CRYPTO) {
        eurInputRows = brokerRowsForEurLedger(PF_CRYPTO, cryptoCombined);
        rawForCount = cryptoCombined;
      } else {
        const rawRows = bundle.brokers[id].rows;
        eurInputRows = brokerRowsForEurLedger(id, rawRows);
        rawForCount = rawRows;
      }
      const p = portfolioEurFromRows(eurInputRows, eurPer);
      return {
        id,
        label: PF_BROKER_LABEL[id],
        rowCount: Array.isArray(rawForCount) ? rawForCount.length : 0,
        ...p,
      };
    });
    const combTot = legs.reduce((s, L) => s + L.totE, 0);
    const combCost = legs.reduce((s, L) => s + L.costE, 0);
    const combPl = combTot - combCost;
    const miss = new Set();
    for (const L of legs) for (const m of L.miss) miss.add(m);
    const missTxt =
      miss.size > 0
        ? `<p class="muted sml" style="margin:8px 0 0">Missing FX: <strong>${esc([...miss].join(", "))}</strong></p>`
        : "";
    const src = esc(j.source || "ECB reference");
    const dt = esc(j.date || "");
    const cell = (label, totE, costE, plE, nRows, legMiss) => {
      if (nRows <= 0) {
        return `<div class="pfEurCell"><span class="muted sml">${esc(label)}</span><div class="pfEurBig muted">—</div><p class="sml muted">No rows</p></div>`;
      }
      if (!totE && !costE && !plE) {
        const m =
          legMiss && legMiss.size > 0
            ? `No ECB rate: ${[...legMiss].join(", ")}`
            : "—";
        return `<div class="pfEurCell"><span class="muted sml">${esc(label)}</span><div class="pfEurBig muted">${esc(fmtMoney("EUR", 0))}</div>
        <div class="sml muted">Cost ${esc(fmtMoney("EUR", 0))}</div>
        <div class="sml">P/L <strong>${esc(fmtMoney("EUR", 0))}</strong></div>
        <p class="sml muted mt" style="margin-bottom:0;opacity:0.85">${esc(m)}</p></div>`;
      }
      const plCls = plE >= 0 ? "plp" : "pln";
      return `<div class="pfEurCell"><span class="muted sml">${esc(label)}</span><div class="pfEurBig">${esc(fmtMoney("EUR", totE))}</div>
        <div class="sml muted">Cost ${esc(fmtMoney("EUR", costE))}</div>
        <div class="sml ${plCls}">P/L <strong>${esc(fmtMoney("EUR", plE))}</strong></div></div>`;
    };
    const legCells = legs
      .map((L) => cell(L.label, L.totE, L.costE, L.plE, L.rowCount, L.miss))
      .join("");
    const combCell = `<div class="pfEurCell pfEurCellHighlight"><span class="muted sml">All ledgers combined</span><div class="pfEurBig">${esc(fmtMoney("EUR", combTot))}</div>
      <div class="sml muted">Cost ${esc(fmtMoney("EUR", combCost))}</div>
      <div class="sml ${combPl >= 0 ? "plp" : "pln"}">P/L <strong>${esc(fmtMoney("EUR", combPl))}</strong></div></div>`;
    box.innerHTML = `<div class="card2 pfCombinedEurInner">
      <div class="h3">By ledger (€) — all accounts</div>
      <p class="sml muted" style="margin:0 0 2px">ECB ref · <strong>${src}</strong> · <strong>${dt}</strong></p>
      <div class="pfEurGrid pfEurGridWide mt">${legCells}${combCell}</div>${missTxt}</div>`;
    if (grand) {
      const missGrand =
        miss.size > 0
          ? `<p class="sml muted mt" style="margin-bottom:0">Some rows omitted: missing rates above.</p>`
          : "";
      const plCls = combPl >= 0 ? "plp" : "pln";
      grand.hidden = false;
      grand.innerHTML = `<div class="card2 pfGrandTotalInner" role="region" aria-label="Total portfolio in euro">
        <div class="sml muted">${src} · ${dt}</div>
        <div class="pfGrandTotalFigRow mt"><span class="pfGrandTotalLab">Total market value</span><span class="pfGrandTotalFig">${esc(fmtMoney("EUR", combTot))}</span></div>
        <div class="pfGrandTotalSubRow sml muted"><span>Cost basis (same FX)</span><span>${esc(fmtMoney("EUR", combCost))}</span></div>
        <div class="pfGrandTotalSubRow sml ${plCls}"><span>Unrealized P/L</span><span><strong>${esc(fmtMoney("EUR", combPl))}</strong></span></div>
        ${missGrand}
      </div>`;
    }
  } catch {
    box.innerHTML = `<div class="card2 pfCombinedEurInner"><p class="err">Could not load EUR reference.</p>
      <p class="sml muted">Check that <code>server.py</code> is running and try <a href="/api/fx-eur" target="_blank" rel="noopener"><code>/api/fx-eur</code></a>. The server retries Frankfurter if the first request fails.</p></div>`;
    if (grand) {
      grand.hidden = false;
      grand.innerHTML = `<div class="card2 pfGrandTotalInner"><p class="err sml">Could not load the portfolio total in €.</p></div>`;
    }
  }
}

/** Trading 212 & Crypto (T212) — portfolio value & P/L converted to EUR (ECB reference via `/api/fx-eur`). */
async function refreshPfT212Euro(rows, multi, oneCcy, tVal, tPl, eurTitle) {
  const title = eurTitle || "Trading 212 — totals in EUR";
  const box = $("pfT212Eur");
  if (!box || !rows.length) return;
  const allEur = rows.every((r) => normalizeCcyForFx(String(r.ccy || "")) === "EUR");
  if (allEur && !multi) {
    let cost = 0;
    for (const row of rows) cost += num(row.qty) * num(row.avg);
    const pl = tVal - cost;
    box.innerHTML = `<div class="h3">${esc(title)}</div>
      <div class="pfEurGrid mt">
        <div><span class="muted sml">Market value</span><div class="pfEurBig">${esc(fmtMoney("EUR", tVal))}</div></div>
        <div><span class="muted sml">Cost basis</span><div>${esc(fmtMoney("EUR", cost))}</div></div>
        <div><span class="muted sml">P/L</span><div class="${pl >= 0 ? "plp" : "pln"}"><strong>${esc(fmtMoney("EUR", pl))}</strong></div></div>
      </div>`;
    return;
  }
  try {
    const j = await fetchEurFxTable();
    const eurPer = j.eur_per_unit;
    const { totE, costE, plE, miss } = portfolioEurFromRows(rows, eurPer);
    const missTxt =
      miss.size > 0
        ? `<p class="muted sml">No ECB rate for <strong>${esc([...miss].join(", "))}</strong> — EUR totals exclude those row(s).</p>`
        : "";
    const src = esc(j.source || "ECB reference");
    const dt = esc(j.date || "");
    if (!totE && !costE && miss.size) {
      box.innerHTML = `<p class="err">Could not convert any row to EUR (missing FX codes).</p>${missTxt}`;
      return;
    }
    box.innerHTML = `<div class="h3">${esc(title)}</div>
      <p class="sml muted" style="margin:0 0 6px">${src} · ${dt}</p>
      <div class="pfEurGrid mt">
        <div><span class="muted sml">Market value</span><div class="pfEurBig">${esc(fmtMoney("EUR", totE))}</div></div>
        <div><span class="muted sml">Cost basis (same FX)</span><div>${esc(fmtMoney("EUR", costE))}</div></div>
        <div><span class="muted sml">P/L</span><div class="${plE >= 0 ? "plp" : "pln"}"><strong>${esc(fmtMoney("EUR", plE))}</strong></div></div>
      </div>${missTxt}`;
  } catch {
    box.innerHTML = `<p class="err">Could not load EUR conversion.</p>`;
  }
}

function val(id) {
  const e = $(id);
  return e instanceof HTMLInputElement ? e.value.trim() : "";
}

function num(x) {
  const n = Number(String(x).replace(/,/g, ""));
  return Number.isFinite(n) ? n : 0;
}

/** Map a user symbol to a Yahoo-style pair for `/api/quote` (crypto in USD). */
function yahooCryptoPairSymbol(sym) {
  const s = String(sym || "")
    .trim()
    .toUpperCase();
  if (!s) return "";
  if (s.includes("-")) return s;
  if (s === "BTC" || s === "XBT") return "BTC-USD";
  if (s === "ETH") return "ETH-USD";
  if (s === "SOL") return "SOL-USD";
  if (s === "XRP") return "XRP-USD";
  if (s === "ADA") return "ADA-USD";
  if (s === "DOGE") return "DOGE-USD";
  if (/^[A-Z0-9]{1,14}$/.test(s)) return `${s}-USD`;
  return s;
}

function buildQuoteUrlParams(sym, ex, brokerId) {
  const u = new URLSearchParams();
  if (brokerId === PF_CRYPTO) {
    u.set("symbol", yahooCryptoPairSymbol(sym));
    const x = String(ex || "").trim();
    if (x && !/^crypto$/i.test(x)) u.set("exchange", x);
  } else {
    u.set("symbol", sym);
    if (ex) u.set("exchange", ex);
  }
  return u;
}

/** T212: one quote per **open position** (`t212Ticker`); the same display sym can list on more than one venue. */
function t212QuoteKey(row) {
  const t = String(row.t212Ticker || "").trim();
  if (t) return t;
  return `${String(row.sym || "").trim().toUpperCase()}|${String(row.ex || "").trim().toUpperCase()}`;
}

function rowMatchesT212QuoteKey(row, key) {
  const t = String(row.t212Ticker || "").trim();
  if (t) return t === key;
  return t212QuoteKey(row) === key;
}

/** Update `last` (and optional `pfKind`) from `/api/quote` for all rows in a broker. */
async function applyLiveQuotesToRowsForBroker(bundle, b) {
  const rows = bundle.brokers[b].rows;
  if (!rows.length) return 0;
  /** Yahoo `*-USD` is USD / coin; some rows get USD→row-ccy using ECB/Frankfurter. */
  let fxEur = null;
  if (b === PF_CRYPTO || b === PF_T212) {
    try {
      fxEur = await fetchEurFxTable();
    } catch {
      fxEur = null;
    }
  }
  const eurPer = fxEur?.eur_per_unit;
  const usdToEur = eurPer ? eurPerUnitToEur(eurPer, "USD") : undefined;

  /** @type { { k: string, row0: (typeof rows)[0] }[] } */
  const iter = [];
  if (b === PF_T212) {
    const seen = new Set();
    for (const row0 of rows) {
      const k = t212QuoteKey(row0);
      if (!k || seen.has(k)) continue;
      if (!String(row0.sym || "").trim()) continue;
      seen.add(k);
      iter.push({ k, row0 });
    }
  } else {
    for (const s of new Set(rows.map((r) => r.sym).filter(Boolean))) {
      const row0 = rows.find((r) => r.sym === s) || null;
      iter.push({ k: s, row0 });
    }
  }
  if (!iter.length) return 0;
  let n = 0;
  for (let i = 0; i < iter.length; i++) {
    const { k, row0 } = iter[i];
    const s = b === PF_T212 && row0 ? String(row0.sym || "").trim() : k;
    if (!s) continue;
    const ex = row0 ? String(row0.ex || "") : rows.find((r) => r.sym === s)?.ex || "";
    status(
      b === PF_T212 && row0
        ? `Quote ${i + 1}/${iter.length}: ${s}${ex ? " · " + ex : ""}`
        : `Quote ${i + 1}/${iter.length}: ${s}`,
    );
    const u =
      b === PF_T212 && row0 && isCryptoLikeT212Row(row0)
        ? buildQuoteUrlParams(s, ex, PF_CRYPTO)
        : buildQuoteUrlParams(s, ex, b);
    try {
      const r = await fetch(`/api/quote?${u}`);
      if (!r.ok) continue;
      const j = await r.json();
      const q = Array.isArray(j) ? j[0] : null;
      const px = num(q?.price);
      if (px > 0) {
        const qCcy = String(q?.currency || "").trim();
        rows.forEach((row) => {
          const same = b === PF_T212 && row0 ? rowMatchesT212QuoteKey(row, t212QuoteKey(row0)) : row.sym === s;
          if (!same) return;
          const ccyN = normalizeCcyForFx(row.ccy);
          let v = px;
          if (b === PF_T212) {
            if (isCryptoLikeT212Row(row) && eurPer && usdToEur != null && ccyN === "EUR") v = px * usdToEur;
            else v = eurPer ? quotePriceToRowCcy(px, qCcy || ccyN, row.ccy, eurPer) : px;
          } else if (eurPer && usdToEur != null && ccyN === "EUR" && b === PF_CRYPTO) {
            v = px * usdToEur;
          }
          row.last = v;
        });
        n++;
      }
      const qt = String(q?.quoteType || q?.instrumentType || "").trim();
      if (qt) {
        rows.forEach((row) => {
          const same = b === PF_T212 && row0 ? rowMatchesT212QuoteKey(row, t212QuoteKey(row0)) : row.sym === s;
          if (same) row.pfKind = qt;
        });
      }
    } catch {
      /* ignore */
    }
    await new Promise((res) => setTimeout(res, 300));
  }
  return n;
}

function normalizeT212SyncedRow(r) {
  if (!r || typeof r !== "object") return r;
  const t = String(r.t212Ticker || r.sym || "").trim();
  const o = { ...r };
  if (t) o.pfRowId = `t212:${t}`.replace(/[^\w.:-]+/g, "_");
  else ensurePfRowId(o);
  if (o.ccy) o.ccy = normalizeCcyForFx(String(o.ccy).toUpperCase());
  return o;
}

function applyT212RowList(rows) {
  return (Array.isArray(rows) ? rows : []).map((x) => normalizeT212SyncedRow({ ...x }));
}

async function fetchTrading212Payload() {
  const r = await fetch("/api/t212/rows", { cache: "no-store" });
  let j = {};
  try {
    j = await r.json();
  } catch {
    j = {};
  }
  if (!r.ok) {
    throw new Error(String(j.detail || j.error || `HTTP ${r.status}`));
  }
  if (!j.ok) {
    throw new Error(String(j.detail || j.error || "t212_error"));
  }
  return j;
}

function formatServerIsoLocal(iso) {
  const s = String(iso || "").trim();
  if (!s) return "—";
  try {
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return s;
    return d.toLocaleString(undefined, { dateStyle: "medium", timeStyle: "short" });
  } catch {
    return s;
  }
}

function paintPfT212SyncTimestamp() {
  const el = $("pfT212SyncTime");
  if (!(el instanceof HTMLElement)) return;
  try {
    const raw = sessionStorage.getItem(K.pfT212LastFetch);
    if (!raw) {
      el.textContent = "";
      el.hidden = true;
      return;
    }
    el.hidden = false;
    el.textContent = `Last T212 server fetch: ${formatServerIsoLocal(raw)}`;
  } catch {
    el.textContent = "";
    el.hidden = true;
  }
}

async function applyTrading212SyncToBundle() {
  const j = await fetchTrading212Payload();
  const bundle = loadPfBundle();
  bundle.brokers[PF_T212].rows = applyT212RowList(j.t212);
  bundle.brokers[PF_CRYPTO].rows = applyT212RowList(j.t212_crypto);
  savePfBundle(bundle);
  try {
    if (typeof j.fetched_at === "string" && j.fetched_at.trim()) {
      sessionStorage.setItem(K.pfT212LastFetch, j.fetched_at.trim());
    }
  } catch {
    /* ignore */
  }
  return j;
}

async function refreshPf() {
  const bundle = loadPfBundle();
  const b = getActiveBroker();
  if (b === PF_CRYPTO) {
    status("Crypto: Trading 212 sync + live USD prices…");
    try {
      const j = await applyTrading212SyncToBundle();
      const ns = Number(j.n_t212 ?? 0) || 0;
      const nc = Number(j.n_t212_crypto ?? 0) || 0;
      status(`T212 sync ok · ${ns} stock/ETF, ${nc} broker crypto — fetching last prices…`);
    } catch (e) {
      status(e instanceof Error ? e.message : String(e));
    }
    const b2 = loadPfBundle();
    const nQ = await applyLiveQuotesToRowsForBroker(b2, PF_CRYPTO);
    savePfBundle(b2);
    renderPf();
    status(nQ > 0 ? `Crypto updated · live prices for ${nQ} symbol(s). EUR cards use your row currency. · ${PF_BROKER_LABEL[PF_CRYPTO]}` : "Crypto: sync done — add BTC/ETH rows or check quote keys if prices stay 0");
    return;
  }
  if (b === PF_T212) {
    status("Syncing from Trading 212…");
    try {
      const j = await applyTrading212SyncToBundle();
      const b2 = loadPfBundle();
      if (b2.brokers[PF_T212].rows.length) await applyLiveQuotesToRowsForBroker(b2, PF_T212);
      if (b2.brokers[PF_CRYPTO].rows.length) await applyLiveQuotesToRowsForBroker(b2, PF_CRYPTO);
      savePfBundle(b2);
      renderPf();
      const ns = Number(j.n_t212 ?? 0) || 0;
      const nc = Number(j.n_t212_crypto ?? 0) || 0;
      status(`T212 sync · ${ns} stock/ETF position(s), ${nc} crypto position(s)`);
    } catch (e) {
      status(e instanceof Error ? e.message : String(e));
    }
    return;
  }
  const b3 = loadPfBundle();
  if (!b3.brokers[b].rows.length) {
    status("Nothing to refresh");
    return;
  }
  const nQ = await applyLiveQuotesToRowsForBroker(b3, b);
  savePfBundle(b3);
  renderPf();
  status(
    nQ > 0 ? `Refresh done · live prices for ${nQ} symbol(s) · ${PF_BROKER_LABEL[b]}` : `Refresh done · no prices updated · ${PF_BROKER_LABEL[b]}`,
  );
}

/** Split one line on comma or semicolon with RFC-style quoted fields (`sep` is `,` or `;`). */
function splitCsvDelim(line, sep) {
  const o = [];
  let c = "";
  let q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"' && line[i + 1] === '"') {
      c += '"';
      i++;
      continue;
    }
    if (ch === '"') {
      q = !q;
      continue;
    }
    if (ch === sep && !q) {
      o.push(c);
      c = "";
      continue;
    }
    c += ch;
  }
  o.push(c);
  return o;
}

function splitCsv(line) {
  return splitCsvDelim(line, ",");
}

/**
 * Pick field separator for a CSV-like file. Numbers (DE) often writes `;` — Excel US then
 * shows one column unless you use Data → Text to columns. We sniff from the header row.
 */
function detectDelimiter(headerLine) {
  if (!headerLine) return ",";
  const tabs = headerLine.split("\t").length;
  if (tabs >= 3) return "\t";
  const byComma = splitCsv(headerLine);
  const bySemi = splitCsvDelim(headerLine, ";");
  if (bySemi.length > byComma.length && bySemi.length >= 2) return ";";
  return ",";
}

function splitRowWithDelim(line, delim) {
  if (delim === "\t") {
    return line.split("\t").map((x) => x.trim().replace(/^"|"$/g, ""));
  }
  if (delim === ";") return splitCsvDelim(line, ";");
  return splitCsv(line);
}

/** Normalize CSV header for matching (Zerodha / Excel add spaces, ₹, dots). */
function hdrKey(x) {
  return String(x ?? "")
    .toLowerCase()
    .replace(/^\ufeff/, "")
    .replace(/₹/g, "inr")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/** First column index whose normalized header equals or starts with any key. */
function hdrPick(Hnorm, keys, mode) {
  const loose = mode === "loose";
  for (const k of keys) {
    const kl = hdrKey(k);
    const i = Hnorm.findIndex((cell) => cell === kl);
    if (i >= 0) return i;
  }
  if (loose) {
    for (const k of keys) {
      const kl = hdrKey(k);
      const i = Hnorm.findIndex((cell) => cell.includes(kl) || (kl.length >= 4 && kl.includes(cell)));
      if (i >= 0) return i;
    }
  }
  return -1;
}

/**
 * Zerodha Console → Portfolio → Holdings → Download XLSX, then Save As CSV in Excel.
 * Headers vary slightly; we match common English column names and default INR.
 * Returns `null` if this does not look like a Zerodha-style holdings sheet.
 */
function tryParseZerodhaHoldings(text) {
  const lines = text.replace(/\r/g, "").split("\n").map((l) => l.trim()).filter(Boolean);
  if (lines.length < 2) return null;
  const delim = detectDelimiter(lines[0]);
  const rawHdr = splitRowWithDelim(lines[0], delim);
  const Hnorm = rawHdr.map(hdrKey);
  if (!Hnorm.some(Boolean)) return null;
  /* App template / generic CSV uses an explicit currency column; Zerodha holdings usually do not. */
  if (Hnorm.includes("currency") || Hnorm.includes("ccy")) return null;

  const iSym = hdrPick(Hnorm, ["symbol", "tradingsymbol", "instrument", "scrip name", "scrip"], "loose");
  const iQty = hdrPick(Hnorm, ["quantity", "qty", "qty net", "net qty", "quantity net", "holding quantity"], "loose");
  if (iSym < 0 || iQty < 0) return null;

  const iEx = hdrPick(Hnorm, ["exchange", "segment", "series"], "loose");
  const iAvg = hdrPick(
    Hnorm,
    [
      "average price",
      "avg price",
      "avgprice",
      "avg cost",
      "average cost",
      "buy average",
      "price average",
      "authorised price",
    ],
    "loose",
  );
  const iLtp = hdrPick(
    Hnorm,
    [
      "ltp",
      "close price",
      "closing price",
      "last price",
      "lastprice",
      "prev close",
      "close",
      "last traded price",
    ],
    "loose",
  );
  const iNm = hdrPick(Hnorm, ["name", "company name", "company", "security name"], "loose");
  const iIsin = hdrPick(Hnorm, ["isin"]);

  const hasPricing = iAvg >= 0 || iLtp >= 0;
  const looksIndian =
    iIsin >= 0 ||
    iEx >= 0 ||
    Hnorm.some((x) => x.includes("nse") || x.includes("bse") || x.includes("isin")) ||
    (hasPricing && hdrPick(Hnorm, ["currency", "ccy"]) < 0);

  if (!looksIndian) return null;

  const out = [];
  for (const ln of lines.slice(1)) {
    const c = splitRowWithDelim(ln, delim);
    if (!c.length || c.every((cell) => !String(cell).trim())) continue;
    let sym = String(c[iSym] || "").trim();
    let ex = iEx >= 0 ? String(c[iEx] || "").trim().toUpperCase() : "";
    const m = sym.match(/^(NSE|BSE|NFO|BFO|CDS|MCX)[:\s]+(.+)$/i);
    if (m) {
      ex = ex || m[1].toUpperCase();
      sym = m[2].trim();
    }
    sym = sym.replace(/\s+(EQ|BE|BZ|SM|ST|ETF|MF|DR)$/i, "").trim();
    const qty = num(c[iQty]);
    if (!sym || qty <= 0) continue;
    const avg = iAvg >= 0 ? num(c[iAvg]) : 0;
    const last = iLtp >= 0 ? num(c[iLtp]) : avg;
    const nm = iNm >= 0 ? String(c[iNm] || "").trim() : "";
    let ccy = "INR";
    if (iIsin >= 0) {
      const isin = String(c[iIsin] || "").trim().toUpperCase();
      if (isin.startsWith("GB")) ccy = "GBP";
      else if (isin.startsWith("US")) ccy = "USD";
      else if (isin.startsWith("DE") || isin.startsWith("FR") || isin.startsWith("NL")) ccy = "EUR";
    }
    out.push({
      sym,
      ex,
      ccy,
      qty,
      avg,
      last: last > 0 ? last : avg,
      nm,
    });
  }
  return out.length ? out : [];
}

/** Map CSV cell to `PF_INS_CO_IDS` entry, or "". */
function normalizeInsCompanyImport(cell) {
  const s = hdrKey(cell);
  if (!s) return "";
  if (s.includes("sbi") && s.includes("life")) return "sbi_life";
  if (s.includes("aditya") || s.includes("birla")) return "aditya_birla";
  if (s.includes("allianz")) return "allianz_retirement";
  if (s.includes("vrk")) return "vrk_retirement";
  if (s === "other" || s.includes("5th") || s.includes("fifth")) return "other";
  return "";
}

function parseInsuranceCsv(text) {
  const lines = text.replace(/\r/g, "").split("\n").map((l) => l.trim()).filter(Boolean);
  if (lines.length < 2) return [];
  const delim = detectDelimiter(lines[0]);
  const row = (ln) => splitRowWithDelim(ln, delim);
  const Hnorm = row(lines[0]).map(hdrKey);
  if (!Hnorm.some(Boolean)) return [];
  const iPn = hdrPick(Hnorm, ["policy name", "policyname", "plan name", "plan"], "loose");
  const iCcy = hdrPick(Hnorm, ["currency", "ccy"], "loose");
  if (iPn < 0 || iCcy < 0) return [];
  const iCo = hdrPick(Hnorm, ["inscompany", "insurance company", "provider", "company"], "loose");
  const iNo = hdrPick(Hnorm, ["policy number", "policyno", "policy no", "policynumber"], "loose");
  const iPd = hdrPick(Hnorm, ["purchase date", "date of purchase", "purchased", "start date"], "loose");
  const iV0 = hdrPick(Hnorm, ["value at purchase", "premium at start", "initial value"], "loose");
  const iGr = hdrPick(Hnorm, ["growth", "avg growth", "growth rate", "average growth"], "loose");
  const iCur = hdrPick(Hnorm, ["current possible value", "current value", "surrender", "fund value"], "loose");
  const iPay = hdrPick(Hnorm, ["payments json", "payments"], "loose");
  const fallbackCo = getPfInsuranceCompany();
  const out = [];
  for (const ln of lines.slice(1)) {
    const c = row(ln);
    if (!c.length || c.every((cell) => !String(cell).trim())) continue;
    const pn = String(c[iPn] || "").trim();
    const ccy = normalizeCcyForFx(String(c[iCcy] || "").trim().toUpperCase());
    if (!pn || !ccy) continue;
    let insCompany = iCo >= 0 ? normalizeInsCompanyImport(c[iCo]) : "";
    if (!isPfInsCoId(insCompany)) insCompany = fallbackCo;
    let payments = [];
    if (iPay >= 0 && String(c[iPay] || "").trim()) {
      try {
        const p = JSON.parse(String(c[iPay]).trim());
        if (Array.isArray(p)) payments = p;
      } catch {
        payments = [];
      }
    }
    const rowObj = {
      policyName: pn,
      policyNo: iNo >= 0 ? String(c[iNo] || "").trim() : "",
      purchaseDate: iPd >= 0 ? String(c[iPd] || "").trim() : "",
      valueAtPurchase: iV0 >= 0 ? num(c[iV0]) : 0,
      growthPct: iGr >= 0 ? num(c[iGr]) : 0,
      currentValue: iCur >= 0 ? num(c[iCur]) : 0,
      ccy,
      insCompany,
      payments,
    };
    ensurePfRowId(rowObj);
    out.push(rowObj);
  }
  return out;
}

function parseFdCsv(text) {
  const lines = text.replace(/\r/g, "").split("\n").map((l) => l.trim()).filter(Boolean);
  if (lines.length < 2) return [];
  const delim = detectDelimiter(lines[0]);
  const row = (ln) => splitRowWithDelim(ln, delim);
  const Hnorm = row(lines[0]).map(hdrKey);
  const iBank = hdrPick(Hnorm, ["fdbank", "bank", "institution"], "loose");
  const iName = hdrPick(Hnorm, ["fdname", "deposit name", "deposit"], "loose");
  const iCcy = hdrPick(Hnorm, ["currency", "ccy"], "loose");
  const iPr = hdrPick(Hnorm, ["principal", "amount", "deposit amount"], "loose");
  if (iBank < 0 || iName < 0 || iCcy < 0 || iPr < 0) return [];
  const iRef = hdrPick(Hnorm, ["fdref", "reference", "receipt"], "loose");
  const iOpen = hdrPick(Hnorm, ["opendate", "open date", "start date"], "loose");
  const iRate = hdrPick(Hnorm, ["ratepct", "rate", "interest rate"], "loose");
  const iCur = hdrPick(Hnorm, ["currentvalue", "current value", "maturity value"], "loose");
  const iMat = hdrPick(Hnorm, ["maturitydate", "maturity date", "maturity"], "loose");
  const iCtry = hdrPick(Hnorm, ["fdcountry", "country", "region", "nation"], "loose");
  const out = [];
  for (const ln of lines.slice(1)) {
    const c = row(ln);
    if (!c.length || c.every((cell) => !String(cell).trim())) continue;
    const bank = String(c[iBank] || "").trim();
    const nm = String(c[iName] || "").trim();
    const ccy = normalizeCcyForFx(String(c[iCcy] || "").trim().toUpperCase());
    const pr = num(c[iPr]);
    if (!bank || !nm || !ccy || pr <= 0) continue;
    const cur = iCur >= 0 ? num(c[iCur]) : pr;
    const rowObj = {
      fdBank: bank,
      fdName: nm,
      fdRef: iRef >= 0 ? String(c[iRef] || "").trim() : "",
      openDate: iOpen >= 0 ? String(c[iOpen] || "").trim() : "",
      principal: pr,
      ratePct: iRate >= 0 ? num(c[iRate]) : 0,
      currentValue: cur > 0 ? cur : pr,
      maturityDate: iMat >= 0 ? String(c[iMat] || "").trim() : "",
      ccy,
      fdCountry: iCtry >= 0 ? String(c[iCtry] || "").trim() : "",
    };
    ensurePfRowId(rowObj);
    out.push(rowObj);
  }
  return out;
}

function parseCsvSmart(text) {
  const lines = text.replace(/\r/g, "").split("\n").map((l) => l.trim()).filter(Boolean);
  if (lines.length < 2) return [];
  const delim = detectDelimiter(lines[0]);
  const row = (ln) => splitRowWithDelim(ln, delim);
  const h = row(lines[0]).map((x) => x.toLowerCase().replace(/^\ufeff/, ""));
  const idx = (n) => h.indexOf(n);
  const t212 =
    h.includes("action") && h.includes("ticker") && h.includes("no. of shares") && h.includes("price / share");
  if (t212) return parseT212(text);
  const zd = tryParseZerodhaHoldings(text);
  if (zd !== null) return zd;
  const iS = idx("symbol");
  const iC = idx("currency");
  const iQ = idx("qty");
  if (iS < 0 || iC < 0 || iQ < 0) return [];
  const iE = idx("exchange");
  const iA = idx("avgprice");
  const iL = idx("lastprice");
  const iN = idx("name");
  const out = [];
  for (const ln of lines.slice(1)) {
    const c = row(ln);
    const sym = (c[iS] || "").trim();
    if (!sym) continue;
    out.push({
      sym,
      ex: iE >= 0 ? (c[iE] || "").trim() : "",
      ccy: normalizeCcyForFx((c[iC] || "").trim().toUpperCase()),
      qty: num(c[iQ]),
      avg: iA >= 0 ? num(c[iA]) : 0,
      last: iL >= 0 ? num(c[iL]) : 0,
      nm: iN >= 0 ? (c[iN] || "").trim() : "",
    });
  }
  return out.filter((r) => r.qty > 0);
}

function parseT212(text) {
  const lines = text.replace(/\r/g, "").split("\n").map((l) => l.trim()).filter(Boolean);
  if (lines.length < 2) return [];
  const delim = detectDelimiter(lines[0]);
  const hdr = splitRowWithDelim(lines[0], delim);
  const I = (n) => hdr.findIndex((x) => x.toLowerCase() === n.toLowerCase());
  const ia = I("Action");
  const it = I("Ticker");
  const is = I("No. of shares");
  const ip = I("Price / share");
  const ic = I("Currency (Price / share)");
  const inm = I("Name");
  const iisin = I("ISIN");
  if (ia < 0 || it < 0 || is < 0 || ip < 0) return [];
  const M = new Map();
  for (const ln of lines.slice(1)) {
    const c = splitRowWithDelim(ln, delim);
    const act = (c[ia] || "").toLowerCase();
    const sym = (c[it] || "").trim();
    if (!sym) continue;
    const buy = act.includes("buy");
    const sell = act.includes("sell");
    if (!buy && !sell) continue;
    const sh = num(c[is]);
    const px = num(c[ip]);
    const ccy = ic >= 0 ? normalizeCcyForFx((c[ic] || "").trim().toUpperCase()) : "";
    if (sh <= 0 || px <= 0) continue;
    const isin = iisin >= 0 ? (c[iisin] || "").trim() : "";
    const k = `${sym}__${ccy}__${isin || "NOISIN"}`;
    const cur = M.get(k) || { sym, ex: "", ccy, qty: 0, avg: 0, last: px };
    if (inm >= 0 && c[inm]) cur.nm = String(c[inm]).trim();
    if (buy) {
      const nq = cur.qty + sh;
      cur.avg = nq > 0 ? (cur.qty * cur.avg + sh * px) / nq : cur.avg;
      cur.qty = nq;
    } else cur.qty = Math.max(0, cur.qty - sh);
    cur.last = px;
    M.set(k, cur);
  }
  return [...M.values()].filter((r) => r.qty > 1e-8);
}

_themeUserPref = loadThemePref();
applyThemePref(_themeUserPref);
if (window.matchMedia) {
  window.matchMedia("(prefers-color-scheme: light)").addEventListener("change", () => {
    if (_themeUserPref === "auto") applyThemePref("auto");
  });
}
$("btnTheme")?.addEventListener("click", () => {
  const next = _themeUserPref === "auto" ? "light" : _themeUserPref === "light" ? "dark" : "auto";
  applyThemePref(next);
});
setupInstall();
setupSw();
window.addEventListener("hashchange", route);
route();
