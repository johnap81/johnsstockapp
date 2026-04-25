#!/usr/bin/env python3
"""
John’sStockApp — minimal API server + static file host.
Endpoints: GET /api/search, GET /api/quote, GET /api/history, GET /api/news, GET /api/corporate, GET /api/fx-eur,
GET/POST /api/ai-commentary, GET /api/ai-ask, GET /api/t212/rows, GET /api/t212/instruments,
PUT /api/shared/portfolio (owner publish), GET /api/shared/portfolio?token= (family read-only)
"""
from __future__ import annotations

import base64
import copy
import errno
import json
import os
import subprocess
import re
import ssl
import xml.etree.ElementTree as ET
import threading
import time
from collections import OrderedDict
from datetime import datetime, timedelta, timezone
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from urllib.error import HTTPError
from urllib.parse import parse_qs, quote as urlquote, urlencode, urlparse
from urllib.request import Request, urlopen

APP_ROOT = os.path.dirname(os.path.abspath(__file__))


def collect_lan_ipv4s() -> list[str]:
    """Best-effort LAN IPs for same-Wi-Fi phone access (macOS `ipconfig` on en* / bridge)."""
    names = [f"en{i}" for i in range(16)] + ["bridge0", "bridge100"]
    seen: set[str] = set()
    out: list[str] = []
    for name in names:
        try:
            r = subprocess.run(
                ["ipconfig", "getifaddr", name],
                capture_output=True,
                text=True,
                timeout=2,
            )
            ip = (r.stdout or "").strip()
            if not ip or ip.startswith("127.") or ip in seen:
                continue
            seen.add(ip)
            out.append(ip)
        except (OSError, subprocess.TimeoutExpired):
            continue
    return out


def write_mobile_url_file(port: int, ips: list[str]) -> str | None:
    """Write copy-paste URLs for Safari on the same Wi-Fi. Returns path or None."""
    path = os.path.join(APP_ROOT, "MOBILE_URL.txt")
    lines = [
        "John'sStockApp — same Wi-Fi as this Mac. Keep Terminal (server) running.",
        "",
        "Paste ONE of these into Safari on your iPhone:",
        "",
    ]
    if ips:
        for ip in ips:
            lines.append(f"http://{ip}:{port}")
    else:
        lines.append(
            "(No LAN IP auto-detected. Use Mac Wi-Fi IP from System Settings → Network → Wi-Fi → Details, then:)"
        )
        lines.append("")
        lines.append(f"http://<YOUR_MAC_IP>:{port}")
    lines.extend(
        [
            "",
            "If Safari cannot connect: System Settings → Network → Firewall → Options →",
            "allow incoming for python3 (or turn firewall off briefly to test).",
        ]
    )
    try:
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return path
    except OSError:
        return None


def load_dotenv(path: str) -> None:
    try:
        with open(path, "r", encoding="utf-8") as f:
            for raw in f.read().splitlines():
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                if k and (k not in os.environ or os.environ.get(k, "") == ""):
                    os.environ[k] = v
    except FileNotFoundError:
        pass


load_dotenv(os.path.join(APP_ROOT, ".env"))


def env(name: str, default: str | None = None) -> str | None:
    v = os.environ.get(name)
    return v if v is not None and v != "" else default


def _truthy(name: str) -> bool:
    v = os.environ.get(name)
    if v is None:
        return False
    return str(v).strip().lower() in ("1", "true", "yes", "on")


def redact_for_json_detail(msg: str) -> str:
    """Strip API keys and tokens from text returned to the browser (errors must not leak .env)."""
    out = str(msg)
    secrets: list[str] = []
    for name in (
        "ALPHAVANTAGE_API_KEY",
        "TWELVE_DATA_API_KEY",
        "EODHD_API_TOKEN",
        "EODHD_API_KEY",
        "FMP_API_KEY",
        "MARKETSTACK_ACCESS_KEY",
        "MARKETSTACK_API_KEY",
        "OPENAI_API_KEY",
        "ANTHROPIC_API_KEY",
        "GOOGLE_AI_API_KEY",
        "GEMINI_API_KEY",
        "XAI_API_KEY",
        "TRADING212_API_KEY",
        "TRADING212_API_SECRET",
    ):
        v = env(name)
        if v and len(v) >= 8:
            secrets.append(v)
    secrets.sort(key=len, reverse=True)
    for v in secrets:
        if v in out:
            out = out.replace(v, "[redacted]")
    # Alpha Vantage: "We have detected your API key as XXXXX ..."
    out = re.sub(r"(?i)(api key as)\s*([A-Z0-9]{10,})\b", r"\1 [redacted]", out)
    out = re.sub(r"([?&])(apikey|api_token|access_key)=[A-Za-z0-9._-]+", r"\1\2=[redacted]", out)
    # AV / vendors sometimes echo a bare key in parentheses or after "apikey="
    out = re.sub(r"\(([A-Z0-9]{16,30})\)", "([redacted])", out)
    out = re.sub(r"(?i)\b(apikey|api key)\s*[=:]\s*([A-Z0-9]{12,30})\b", r"\1 [redacted]", out)
    return out


def _rate_limit_hint_from_message(msg: str) -> str | None:
    u = msg.lower()
    if any(
        x in u
        for x in (
            "429",
            "402",
            "too many requests",
            "rate limit",
            "api credits",
            "requests per day",
            "run out of",
            "premium plans",
            "payment required",
        )
    ):
        return (
            "Every configured source refused this call (free-tier limits: per-minute credits, daily caps, or Yahoo throttling). "
            "Wait a few minutes, try again later in the day, or upgrade one provider you rely on."
        )
    return None


def _use_yahoo() -> bool:
    if os.environ.get("USE_YAHOO_FINANCE") is None:
        return True
    return _truthy("USE_YAHOO_FINANCE")


def json_response(handler: SimpleHTTPRequestHandler, status: int, payload: object) -> None:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Cache-Control", "no-store")
    handler.send_header("Content-Length", str(len(data)))
    handler.end_headers()
    handler.wfile.write(data)


def fetch_json(url: str, timeout_s: float = 15.0, extra: dict[str, str] | None = None) -> object:
    """GET JSON from URL. Retries on 429/503 (rate limit / overload) with backoff."""
    h = {
        "User-Agent": "JohnsStockApp/2.0",
        "Accept": "application/json",
    }
    if extra:
        h.update(extra)
    insecure = (env("INSECURE_SSL", "0") or "0").strip() == "1"
    delays_429 = (1.2, 3.0, 6.0, 12.0, 20.0, 32.0)
    attempt = 0
    while True:
        req = Request(url, headers=h)
        try:
            ctx = ssl._create_unverified_context() if insecure else ssl.create_default_context()
            with urlopen(req, timeout=timeout_s, context=ctx) as resp:
                return json.loads(resp.read().decode("utf-8", errors="replace"))
        except HTTPError as e:
            if e.code in (429, 503) and attempt < len(delays_429):
                try:
                    e.read()
                except Exception:  # noqa: BLE001
                    pass
                time.sleep(delays_429[attempt])
                attempt += 1
                continue
            try:
                body = e.read().decode("utf-8", errors="replace")[:500]
                raise RuntimeError(f"{e} • {body}") from e
            except RuntimeError:
                raise
            except Exception:  # noqa: BLE001
                raise
        except ssl.SSLCertVerificationError:
            ctx = ssl._create_unverified_context()
            with urlopen(req, timeout=timeout_s, context=ctx) as resp:
                return json.loads(resp.read().decode("utf-8", errors="replace"))
        except Exception as e:  # noqa: BLE001
            reason = getattr(e, "reason", None)
            if isinstance(reason, ssl.SSLCertVerificationError):
                ctx = ssl._create_unverified_context()
                with urlopen(req, timeout=timeout_s, context=ctx) as resp:
                    return json.loads(resp.read().decode("utf-8", errors="replace"))
            raise


def _t212_float(x: object) -> float:
    try:
        return float(x)  # type: ignore[arg-type]
    except (TypeError, ValueError):
        return 0.0


def trading212_base_url() -> str:
    m = (env("TRADING212_ENV", "live") or "live").strip().lower()
    if m in ("demo", "paper", "test", "practice"):
        return "https://demo.trading212.com"
    return "https://live.trading212.com"


def trading212_auth_header() -> str | None:
    k = env("TRADING212_API_KEY")
    s = env("TRADING212_API_SECRET")
    if not k or not s:
        return None
    pair = f"{k}:{s}".encode("utf-8")
    return "Basic " + base64.b64encode(pair).decode("ascii")


def _t212_normalize_positions_list(data: object) -> list[dict]:
    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]
    if isinstance(data, dict):
        items = data.get("items")
        if isinstance(items, list):
            return [x for x in items if isinstance(x, dict)]
    return []


def _t212_display_sym(ticker: str) -> str:
    t = (ticker or "").strip()
    if not t:
        return "—"
    parts = t.split("_")
    if len(parts) >= 3 and parts[-1].upper() == "EQ" and parts[0]:
        return parts[0]
    return t


def _t212_infer_exchange(ticker: str) -> str:
    u = (ticker or "").upper()
    if "_US_EQ" in u or u.endswith("_US_EQ"):
        return "NASDAQ"
    if "_GB_EQ" in u or u.endswith("_GB_EQ"):
        return "LSE"
    if "_DE_EQ" in u or u.endswith("_DE_EQ"):
        return "XETRA"
    if "_EU_EQ" in u or u.endswith("_EU_EQ"):
        return "EU"
    return ""


def _t212_position_is_crypto(inst: dict, p: dict) -> bool:
    t = str((inst or {}).get("type") or "").upper()
    if t and ("CRYPT" in t or t in ("COIN", "CRYPTOCURRENCY")):
        return True
    if t in ("STOCK", "EQUITY", "STOCKS", "ETF", "ETFS", "FUND", "EQUITIES"):
        return False
    tick = str((inst or {}).get("ticker") or p.get("ticker") or "").upper()
    if "CRYPTO" in tick or "_CRY" in tick:
        return True
    return False


def handle_t212_rows(handler: SimpleHTTPRequestHandler) -> None:
    """Read-only: GET /api/t212/rows — open positions split into stock-style vs crypto (instrument.type / ticker)."""
    auth = trading212_auth_header()
    if not auth:
        return json_response(
            handler,
            503,
            {
                "ok": False,
                "error": "trading212_not_configured",
                "detail": "Set TRADING212_API_KEY and TRADING212_API_SECRET in .env (optional TRADING212_ENV=live|demo).",
            },
        )
    base = trading212_base_url().rstrip("/")
    url = f"{base}/api/v0/equity/positions"
    try:
        data = fetch_json(url, timeout_s=25.0, extra={"Authorization": auth})
    except Exception as e:  # noqa: BLE001
        return json_response(
            handler,
            502,
            {
                "ok": False,
                "error": "t212_fetch_failed",
                "detail": redact_for_json_detail(str(e))[:800],
            },
        )
    rows_raw = _t212_normalize_positions_list(data)
    account_ccy = _t212_account_ccy_for_positions(auth)
    eur_per = _t212_eur_per_for_t212()
    stocks: list[dict] = []
    crypto: list[dict] = []
    for p in rows_raw:
        if not isinstance(p, dict):
            continue
        inst = p.get("instrument")
        if not isinstance(inst, dict):
            inst = {}
        is_c = _t212_position_is_crypto(inst, p)
        row = _t212_build_position_row(p, inst, is_c, account_ccy, eur_per)
        if is_c:
            crypto.append(row)
        else:
            stocks.append(row)
    env_tag = "demo" if "demo" in base else "live"
    return json_response(
        handler,
        200,
        {
            "ok": True,
            "t212": stocks,
            "t212_crypto": crypto,
            "n_t212": len(stocks),
            "n_t212_crypto": len(crypto),
            "env": env_tag,
        },
    )


# --- Trading 212 instruments cache (read-only) — 1 request / 50s per T212; warm in background for search ---
T212I_LOCK = threading.Lock()
T212I_ITEMS: list[dict] = []
T212I_STATUS: dict[str, object] = {
    "pages": 0,
    "total_rows": 0,
    "complete": False,
    "loading": False,
    "err": None,
    "last_page_at": 0.0,
}
T212I_WARMER_STARTED = False


def _t212_ticker_looks_eu(ticker: str) -> bool:
    u = (ticker or "").upper()
    for m in (
        "_DE_EQ",
        "_FR_EQ",
        "_NL_EQ",
        "_ES_EQ",
        "_IT_EQ",
        "_AT_EQ",
        "_BE_EQ",
        "_IE_EQ",
        "_PL_EQ",
        "_SE_EQ",
        "_DK_EQ",
        "_FI_EQ",
        "_PT_EQ",
        "_CH_EQ",  # often used for SIX; still European venue
    ):
        if m in u:
            return True
    return False


def _t212_instrument_to_search_row(o: dict) -> dict:
    t = str(o.get("ticker") or "").strip()
    name = str(o.get("name") or o.get("shortName") or "").strip()
    ccy = str(o.get("currencyCode") or "").strip()
    qtyp = str(o.get("type") or "").strip()
    isin = str(o.get("isin") or "").strip()
    return {
        "symbol": _t212_display_sym(t) or t or "—",
        "name": name,
        "exchangeShortName": "T212",
        "country": "EU" if _t212_ticker_looks_eu(t) else ("" if t.endswith("_US_EQ") else "—"),
        "currency": ccy,
        "quoteType": qtyp,
        "t212Ticker": t,
        "isin": isin,
    }


def _t212_filter_cached_instruments(
    q: str,
    region: str,
    itype: str,
    limit: int,
) -> tuple[list[dict], int, bool, int]:
    """Return (matches, total_cache_size, cache_complete, matches_before_limit)."""
    ql = (q or "").strip().lower()
    itype_u = (itype or "").strip().upper()
    reg = (region or "all").strip().lower()
    lim = min(max(1, limit), 80)
    with T212I_LOCK:
        total = len(T212I_ITEMS)
        complete = bool(T212I_STATUS.get("complete"))
        copy_items = [x for x in T212I_ITEMS if isinstance(x, dict)]
    out: list[dict] = []
    for o in copy_items:
        t = str(o.get("ticker") or "")
        ttyp = str(o.get("type") or "").upper()
        if itype_u and ttyp != itype_u:
            continue
        if reg == "eu" and not _t212_ticker_looks_eu(t):
            continue
        if ql:
            name = str(o.get("name") or o.get("shortName") or "")
            isin = str(o.get("isin") or "")
            blob = f"{t} {name} {isin}".lower()
            words = [w for w in re.split(r"\s+", ql) if w]
            if words and not all(w in blob for w in words):
                continue
        out.append(_t212_instrument_to_search_row(o))
    raw_n = len(out)
    return out[:lim], total, complete, raw_n


def _t212_preload_instruments_enabled() -> bool:
    v = (env("TRADING212_PRELOAD_INSTRUMENTS", "1") or "1").strip().lower()
    return v not in ("0", "false", "off", "no")


def t212_instruments_warmer_start(*, from_boot: bool = False) -> None:
    """Background thread: page /api/v0/equity/metadata/instruments (T212 rate limit: ~1 call / 50s)."""
    global T212I_WARMER_STARTED, T212I_ITEMS, T212I_STATUS
    if from_boot and not _t212_preload_instruments_enabled():
        return
    auth = trading212_auth_header()
    if not auth:
        return
    with T212I_LOCK:
        if T212I_STATUS.get("loading"):
            return
        if T212I_STATUS.get("complete"):
            return
        if T212I_WARMER_STARTED:
            return
        if T212I_STATUS.get("err"):
            T212I_ITEMS.clear()
            T212I_STATUS["pages"] = 0
            T212I_STATUS["total_rows"] = 0
        T212I_WARMER_STARTED = True
        T212I_STATUS["loading"] = True
        T212I_STATUS["err"] = None

    def _run() -> None:
        global T212I_WARMER_STARTED, T212I_ITEMS, T212I_STATUS
        base = trading212_base_url().rstrip("/")
        next_path: str | None = f"{base}/api/v0/equity/metadata/instruments?limit=50"
        try:
            while next_path is not None:
                u = next_path
                with T212I_LOCK:
                    T212I_STATUS["last_page_at"] = time.time()
                data = fetch_json(u, timeout_s=90.0, extra={"Authorization": auth})
                items: list[dict] = []
                nxt: str | None = None
                if isinstance(data, dict):
                    it = data.get("items")
                    nxt = data.get("nextPagePath")
                    if isinstance(it, list):
                        items = [x for x in it if isinstance(x, dict)]
                elif isinstance(data, list):
                    items = [x for x in data if isinstance(x, dict)]
                with T212I_LOCK:
                    T212I_ITEMS.extend(items)
                    T212I_STATUS["pages"] = int(T212I_STATUS.get("pages") or 0) + 1
                    T212I_STATUS["total_rows"] = len(T212I_ITEMS)
                if nxt in (None, ""):
                    with T212I_LOCK:
                        T212I_STATUS["complete"] = True
                    break
                nst = str(nxt)
                if nst.startswith("http://") or nst.startswith("https://"):
                    next_path = nst
                else:
                    p = nst[1:] if nst.startswith("/") else nst
                    next_path = f"{base}/" + p
                time.sleep(51.0)
        except Exception as e:  # noqa: BLE001
            with T212I_LOCK:
                T212I_STATUS["err"] = redact_for_json_detail(str(e))[:500]
        finally:
            with T212I_LOCK:
                T212I_STATUS["loading"] = False
                if T212I_STATUS.get("err"):
                    T212I_WARMER_STARTED = False

    threading.Thread(target=_run, name="t212-instruments", daemon=True).start()


def handle_t212_instruments(handler: SimpleHTTPRequestHandler, qs: dict[str, list[str]]) -> None:
    """
    Read-only: GET /api/t212/instruments — filter cached T212 metadata (q, region=eu|all, type, limit).
    T212’s API is paginated; first useful results may appear only after the background warm-up runs for a while.
    """
    q = (qs.get("q", [""])[0] or qs.get("query", [""])[0] or "").strip()
    region = (qs.get("region", ["all"])[0] or "all").strip().lower()
    itype = (qs.get("type", [""])[0] or qs.get("instrument", [""])[0] or "").strip()
    try:
        lim = int((qs.get("limit", ["25"])[0] or "25").strip() or "25")
    except ValueError:
        lim = 25
    for bad in (region, itype, q):
        if not isinstance(bad, str) or any(c in bad for c in "\n\r\x00") or len(bad) > 200:
            return json_response(
                handler,
                400,
                {"ok": False, "error": "invalid_query", "detail": "q/region/type too long or invalid"},
            )
    auth = trading212_auth_header()
    with T212I_LOCK:
        empty_cache = not T212I_ITEMS
    if not auth and empty_cache:
        return json_response(
            handler,
            200,
            {
                "ok": True,
                "matches": [],
                "n": 0,
                "n_cached": 0,
                "cache_complete": False,
                "hint": "Set TRADING212_API_KEY and TRADING212_API_SECRET. Cache starts after preload is enabled and server is restarted.",
            },
        )
    if auth:
        t212_instruments_warmer_start()
    m, n_cached, comp, raw_m = _t212_filter_cached_instruments(q, region, itype, lim)
    with T212I_LOCK:
        loading = bool(T212I_STATUS.get("loading"))
        err = T212I_STATUS.get("err")
    hint: str | None = None
    if not m and not comp and n_cached < 200:
        hint = (
            "Trading 212’s instruments list is loading in the background (strict rate limit: 1 call / 50s). "
            "Search again in a few minutes, or use Yahoo / FMP results for tickers, then use T212’s ticker in your portfolio sync."
        )
    return json_response(
        handler,
        200,
        {
            "ok": True,
            "matches": m,
            "n": len(m),
            "n_matched_raw": raw_m,
            "n_cached": n_cached,
            "cache_complete": comp,
            "cache_loading": loading,
            "cache_error": str(err) if err else None,
            "q": q,
            "region": region,
            "type_filter": itype,
            "hint": hint,
        },
    )


YAHOO_EXTRA = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://finance.yahoo.com/",
}


def _price_ok(row: dict) -> bool:
    if not isinstance(row, dict):
        return False
    for k in ("price", "close", "last", "regularMarketPrice", "previousClose", "previous_close"):
        v = row.get(k)
        if v in (None, ""):
            continue
        try:
            n = float(str(v).replace(",", ""))
        except (TypeError, ValueError):
            continue
        if n == n and abs(n) < 1e15:
            return True
    return False


def yahoo_resolve(symbol: str, exchange: str | None) -> str:
    s = (symbol or "").strip()
    if not s or "." in s.upper():
        return s
    ex = (exchange or "").strip().upper().replace(" ", "")
    if ex in ("NSE", "NS"):
        return f"{s}.NS"
    if ex in ("BSE", "BOM", "BO"):
        return f"{s}.BO"
    if ex in ("LSE", "LON", "L"):
        return f"{s}.L"
    if ex in ("XETRA", "ETR", "GER", "DE", "FRA", "DUS", "MUN", "STU", "GETTEX"):
        return f"{s}.DE"
    if ex in ("SW", "SWX"):
        return f"{s}.SW"
    if ex in ("PA", "EPA"):
        return f"{s}.PA"
    if ex in ("AS", "AMS"):
        return f"{s}.AS"
    if ex in ("TO", "TSX", "TSXV", "V", "CN"):
        return f"{s}.TO"
    return s


def yahoo_chart_candidates(symbol: str, exchange: str | None) -> list[str]:
    base = (symbol or "").strip()
    if not base:
        return []
    seen: set[str] = set()
    out: list[str] = []
    for c in (yahoo_resolve(base, exchange), base, f"{base}.NS", f"{base}.BO", f"{base}.L", f"{base}.DE"):
        x = (c or "").strip()
        if x and x.upper() not in seen:
            seen.add(x.upper())
            out.append(x)
    return out


def _yahoo_meta_quote_extras(meta: dict) -> dict[str, object]:
    """Optional valuation / liquidity fields from Yahoo chart `meta` (when Yahoo sends them)."""
    out: dict[str, object] = {}
    if not isinstance(meta, dict):
        return out
    pairs = (
        ("marketCap", "marketCap"),
        ("trailingPE", "trailingPE"),
        ("forwardPE", "forwardPE"),
        ("priceToBook", "priceToBook"),
        ("beta", "beta"),
        ("bookValue", "bookValue"),
        ("epsTrailingTwelveMonths", "epsTrailing"),
        ("epsForward", "epsForward"),
        ("dividendYield", "dividendYield"),
        ("dividendRate", "dividendRate"),
        ("fiftyDayAverage", "avgPrice50d"),
        ("twoHundredDayAverage", "avgPrice200d"),
        ("averageDailyVolume3Month", "avgVolume3Mo"),
        ("averageDailyVolume10Day", "avgVolume10d"),
        ("regularMarketChange", "changeAmount"),
        ("regularMarketChangePercent", "changePercent"),
    )
    for src, dst in pairs:
        v = meta.get(src)
        if v in (None, ""):
            continue
        try:
            if isinstance(v, bool):
                out[dst] = v
            elif isinstance(v, (int, float)):
                fv = float(v)
                if fv == fv:
                    out[dst] = fv
            elif isinstance(v, str):
                s = v.strip().replace(",", "")
                if s:
                    out[dst] = float(s) if re.match(r"^-?\d", s) else v
        except (TypeError, ValueError):
            out[dst] = str(v)
    return out


def yahoo_quote(symbol: str, exchange: str | None) -> dict:
    err: Exception | None = None
    for ys in yahoo_chart_candidates(symbol, exchange):
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{urlquote(ys, safe='-.')}?" + urlencode(
            {"range": "5d", "interval": "1d"}
        )
        try:
            payload = fetch_json(url, extra=YAHOO_EXTRA)
            chart = payload.get("chart") if isinstance(payload, dict) else None
            if not isinstance(chart, dict) or chart.get("error"):
                raise RuntimeError(str(chart.get("error") if isinstance(chart, dict) else "bad chart"))
            res = (chart.get("result") or [{}])[0]
            meta = res.get("meta") or {}
            price = meta.get("regularMarketPrice")
            prev = meta.get("chartPreviousClose") or meta.get("previousClose")
            if price in (None, "") and isinstance(res.get("indicators"), dict):
                q = (res["indicators"].get("quote") or [{}])[0]
                for x in reversed((q.get("close") or []) if isinstance(q, dict) else []):
                    if x is not None:
                        price = x
                        break
            qt = str(meta.get("instrumentType") or meta.get("quoteType") or "").strip()
            row = {
                "symbol": str(meta.get("symbol") or ys),
                "name": str(meta.get("longName") or meta.get("shortName") or ""),
                "price": price,
                "previousClose": prev,
                "open": meta.get("regularMarketOpen"),
                "dayHigh": meta.get("regularMarketDayHigh"),
                "dayLow": meta.get("regularMarketDayLow"),
                "yearHigh": meta.get("fiftyTwoWeekHigh"),
                "yearLow": meta.get("fiftyTwoWeekLow"),
                "volume": meta.get("regularMarketVolume"),
                "exchange": str(meta.get("exchangeName") or ""),
                "currency": str(meta.get("currency") or ""),
                "timestamp": meta.get("regularMarketTime"),
                "_provider": "yahoo",
                "_note": "Yahoo chart (unofficial, no key).",
            }
            if qt:
                row["quoteType"] = qt
            row.update(_yahoo_meta_quote_extras(meta))
            if _price_ok(row):
                return row
            err = RuntimeError("no price")
        except Exception as e:  # noqa: BLE001
            err = e
    raise RuntimeError(f"Yahoo: {err}")


_HISTORY_RANGES = frozenset({"5d", "1mo", "3mo", "6mo", "1y", "2y", "5y", "max"})
# Symbols reused for many RS55 benchmark fetches — longer OK-cache reduces duplicate upstream calls.
_HISTORY_INDEX_CACHE_SYMBOLS = frozenset({"SPY", "^NSEI", "^GSPC", "QQQ", "DIA", "IWM"})


def _history_max_bars(range_key: str) -> int:
    return {"5d": 8, "1mo": 24, "3mo": 70, "6mo": 140, "1y": 300, "2y": 550, "5y": 1300, "max": 3500}.get(
        range_key, 300
    )


def yahoo_history(symbol: str, exchange: str | None, range_key: str) -> tuple[list[dict], str]:
    """Daily OHLCV from Yahoo chart API (same unofficial endpoint as quotes)."""
    rk = range_key if range_key in _HISTORY_RANGES else "1y"
    interval = "1d"
    last_err: Exception | None = None
    for ys in yahoo_chart_candidates(symbol, exchange):
        url = f"https://query1.finance.yahoo.com/v8/finance/chart/{urlquote(ys, safe='-.')}?" + urlencode(
            {"range": rk, "interval": interval, "includePrePost": "false"}
        )
        try:
            payload = fetch_json(url, extra=YAHOO_EXTRA, timeout_s=24.0)
            chart = payload.get("chart") if isinstance(payload, dict) else None
            if not isinstance(chart, dict) or chart.get("error"):
                raise RuntimeError(str(chart.get("error") if isinstance(chart, dict) else "bad chart"))
            res = (chart.get("result") or [None])[0]
            if not isinstance(res, dict):
                raise RuntimeError("no result")
            ts = res.get("timestamp")
            if not isinstance(ts, list):
                raise RuntimeError("no timestamps")
            ind = res.get("indicators") if isinstance(res.get("indicators"), dict) else {}
            qrows = (ind.get("quote") or [None])[0]
            if not isinstance(qrows, dict):
                raise RuntimeError("no quote")
            opens = qrows.get("open") or []
            highs = qrows.get("high") or []
            lows = qrows.get("low") or []
            closes = qrows.get("close") or []
            vols = qrows.get("volume") or []
            bars: list[dict] = []
            for i, t in enumerate(ts):
                if t is None:
                    continue
                c = closes[i] if i < len(closes) else None
                if c is None:
                    continue
                o = opens[i] if i < len(opens) else c
                h = highs[i] if i < len(highs) else c
                l = lows[i] if i < len(lows) else c
                v = vols[i] if i < len(vols) else None
                bars.append(
                    {
                        "t": int(t),
                        "o": float(o) if o is not None else float(c),
                        "h": float(h) if h is not None else float(c),
                        "l": float(l) if l is not None else float(c),
                        "c": float(c),
                        "v": float(v) if v is not None else None,
                    }
                )
            if len(bars) >= 2:
                return bars, "yahoo"
            last_err = RuntimeError("too few bars")
        except Exception as e:  # noqa: BLE001
            last_err = e
            continue
    raise RuntimeError(f"Yahoo history: {last_err}" if last_err else "Yahoo history failed")


# Short-lived memo for Yahoo search only (personal use: fewer repeat calls → fewer 429s).
_YAHOO_SRCH_CACHE: dict[tuple[str, int], tuple[float, list[dict]]] = {}
_YAHOO_SRCH_MAX_KEYS = 200


def _yahoo_search_cache_ttl_s() -> float:
    """Seconds to reuse identical Yahoo search results; 0 disables. Default 300, max 600."""
    try:
        v = float((env("YAHOO_SEARCH_CACHE_SECONDS", "300") or "300").strip())
    except ValueError:
        v = 300.0
    return max(0.0, min(v, 600.0))


def _yahoo_search_prune(now: float, ttl: float) -> None:
    for k in list(_YAHOO_SRCH_CACHE.keys()):
        if now - _YAHOO_SRCH_CACHE[k][0] > ttl:
            del _YAHOO_SRCH_CACHE[k]
    while len(_YAHOO_SRCH_CACHE) > _YAHOO_SRCH_MAX_KEYS:
        oldest = min(_YAHOO_SRCH_CACHE.items(), key=lambda kv: kv[1][0])[0]
        del _YAHOO_SRCH_CACHE[oldest]


def _yahoo_search_fetch(q: str, limit: int) -> list[dict]:
    url = "https://query1.finance.yahoo.com/v1/finance/search?" + urlencode(
        {"q": q.strip(), "quotesCount": str(min(max(limit, 1), 50)), "newsCount": "0"}
    )
    payload = fetch_json(url, extra=YAHOO_EXTRA)
    quotes = payload.get("quotes") if isinstance(payload, dict) else None
    if not isinstance(quotes, list):
        return []
    out: list[dict] = []
    for r in quotes:
        if not isinstance(r, dict):
            continue
        sym = str(r.get("symbol") or "").strip()
        if not sym:
            continue
        ex_disp = str(r.get("exchDisp") or "").strip()
        ex_mic = str(r.get("exchange") or "").strip()
        ex_show = ex_disp or ex_mic
        qt = str(r.get("quoteType") or r.get("typeDisp") or "").strip()
        out.append(
            {
                "symbol": sym,
                "name": str(r.get("longname") or r.get("shortname") or r.get("longName") or ""),
                "exchangeShortName": ex_show,
                "exchangeMic": ex_mic,
                "currency": str(r.get("currency") or ""),
                "quoteType": qt,
            }
        )
    return finalize_search_results(out)


def yahoo_search(q: str, limit: int) -> list[dict]:
    ttl = _yahoo_search_cache_ttl_s()
    lim = min(max(limit, 1), 50)
    key = (q.strip().lower()[:120], lim)
    now = time.monotonic()
    if ttl > 0:
        hit = _YAHOO_SRCH_CACHE.get(key)
        if hit and now - hit[0] < ttl:
            return [dict(row) for row in hit[1]]
    out = _yahoo_search_fetch(q, lim)  # already finalize_search_results inside fetch
    if ttl > 0 and out:
        _YAHOO_SRCH_CACHE[key] = (now, out)
        _yahoo_search_prune(now, ttl)
    return out


# --- Trading 212: account currency, ISIN → Yahoo, venue → row ccy (API row prices are in account ccy) ---
T212Y2_LOCK = threading.Lock()
T212_ISIN_PICK: dict[str, tuple[float, dict | None]] = {}


def _t212_account_ccy_for_positions(auth: str) -> str:
    try:
        base = trading212_base_url().rstrip("/")
        data = fetch_json(
            f"{base}/api/v0/equity/account/summary",
            timeout_s=12.0,
            extra={"Authorization": auth},
        )
        if isinstance(data, dict):
            c = str(
                data.get("currency")
                or data.get("currencyCode")
                or data.get("accountCurrency")
                or ""
            ).strip().upper()
            if c and len(c) == 3 and c.isalpha():
                return c
    except Exception:  # noqa: BLE001
        pass
    return "USD"


def _t212_eur_per_for_t212() -> dict[str, float] | None:
    try:
        p = build_fx_eur_payload()
        m = p.get("eur_per_unit")
        if not isinstance(m, dict):
            return None
        out: dict[str, float] = {}
        for k, v in m.items():
            c = str(k).upper()
            if not c:
                continue
            try:
                f = float(v)  # type: ignore[arg-type]
            except (TypeError, ValueError):
                continue
            if f == f and f > 0:
                out[c] = f
        return out or None
    except Exception:  # noqa: BLE001
        return None


def _t212_fx_convert_amount(amt: float, fr: str, to: str, eur_per: dict[str, float] | None) -> float:
    a = (fr or "").strip().upper()
    b = (to or "").strip().upper()
    if eur_per is None or a == b or not a or not b or not (amt and amt == amt):
        return amt
    p1 = eur_per.get(a)
    p2 = eur_per.get(b)
    if p1 is None or p2 is None or p2 == 0.0:
        return amt
    return (float(amt) * p1) / p2


def _t212_find_instrument_in_cache(ticker: str) -> dict:
    t = (ticker or "").strip()
    if not t:
        return {}
    with T212I_LOCK:
        for o in T212I_ITEMS:
            if isinstance(o, dict) and str(o.get("ticker") or "").strip() == t:
                return dict(o)
    return {}


def _t212_venue_code_from_ticker(ticker: str) -> str:
    t = (ticker or "").strip().upper()
    if t.endswith("_EQ") and t.count("_") >= 2:
        parts = t.rsplit("_", 2)
        if len(parts) == 3 and parts[2] == "EQ" and (2 <= len(parts[1]) <= 4) and parts[0]:
            return parts[1]
    return ""


def _t212_listing_ccy_from_venue(venue: str) -> str | None:
    v = (venue or "").strip().upper()
    m = {
        "US": "USD",
        "GB": "GBP",
        "CH": "CHF",
        "SE": "SEK",
        "NO": "NOK",
        "DK": "DKK",
        "PL": "PLN",
        "CZ": "CZK",
        "HU": "HUF",
        "RO": "RON",
        "DE": "EUR",
        "FR": "EUR",
        "IT": "EUR",
        "ES": "EUR",
        "NL": "EUR",
        "AT": "EUR",
        "BE": "EUR",
        "IE": "EUR",
        "PT": "EUR",
        "FI": "EUR",
        "EU": "EUR",
        "GR": "EUR",
        "LU": "EUR",
    }
    return m.get(v)


def _t212_ex_from_ticker_venue(venue: str) -> str:
    v = (venue or "").strip().upper()
    m = {
        "US": "NASDAQ",
        "GB": "LSE",
        "DE": "XETRA",
        "AT": "VIE",
        "FR": "EPA",
        "NL": "AMS",
        "IT": "MIL",
        "ES": "BME",
        "SE": "STO",
        "NO": "OSL",
        "CH": "SWX",
        "BE": "BRU",
        "IE": "LSE",
        "EU": "XETRA",
        "PL": "WAR",
    }
    if v in m:
        return m[v]
    if v in ("CZ", "HU", "RO", "PL"):
        return "XETRA"
    return "NASDAQ"


def _t212_yahoo_ex_from_listing_symbol(ys: str) -> str:
    u = (ys or "").strip().upper()
    for suf, ex in (
        (".DE", "XETRA"),
        (".F", "FRA"),
        (".L", "LSE"),
        (".PA", "EPA"),
        (".AS", "AMS"),
        (".MI", "MIL"),
        (".MC", "BME"),
        (".SW", "SWX"),
        (".ST", "STO"),
        (".OL", "OSL"),
        (".VI", "VIE"),
        (".IR", "LSE"),
        (".BR", "BRU"),
        (".WA", "WAR"),
    ):
        if u.endswith(suf):
            return ex
    if "." not in (ys or ""):
        return "NASDAQ"
    return "NASDAQ"


def _t212_display_base_from_yahoo(ys: str) -> str:
    yf = (ys or "").strip()
    if not yf:
        return yf
    u = yf.upper()
    for suf in (".DE", ".L", ".PA", ".AS", ".MI", ".MC", ".SW", ".ST", ".OL", ".VI", ".F", ".IR", ".CO", ".TO", ".NS", ".BO", ".BR", ".WA"):
        if u.endswith(suf) and len(yf) > len(suf) + 1:
            return yf[: -len(suf)]
    if "." in yf:
        return yf.rsplit(".", 1)[0]
    return yf


def _t212_yahoo_isin_cache_get(key: str) -> dict | None:
    with T212Y2_LOCK:
        ent = T212_ISIN_PICK.get(key)
    if not ent:
        return None
    ts, row = ent
    if time.monotonic() - ts > 4 * 3600.0:  # 4h
        with T212Y2_LOCK:
            T212_ISIN_PICK.pop(key, None)
        return None
    return row if isinstance(row, dict) else None


def _t212_yahoo_isin_cache_set(key: str, row: dict | None) -> None:
    with T212Y2_LOCK:
        T212_ISIN_PICK[key] = (time.monotonic(), row)


def _t212_yahoo_isin_row(isin: str, venue: str) -> dict | None:
    q = re.sub(r"\s+", "", (isin or "")).upper()
    if not q or len(q) < 9:
        return None
    ven = (venue or "").upper()
    ckey = f"{q}|{ven}"
    hit = _t212_yahoo_isin_cache_get(ckey)
    if hit is not None:
        return hit

    def score_row(r: dict) -> int:
        symu = (str(r.get("symbol") or "")).upper()
        qt = (str(r.get("quoteType") or "")).upper()
        if qt and "CRYPT" in qt:
            return -1000
        sc = 0
        if "EQU" in qt or "ETF" in qt or qt in ("EQUITY", "ETF") or "STOCK" in qt:
            sc += 1
        if ven in ("DE", "XETRA", "FRA", "DUS", "MUN", "STU", "ETR", "GETTEX"):
            if symu.endswith((".DE", ".F", ".DUS", ".FRA", ".F")):
                sc += 20
        elif ven in ("GB", "LON", "LSE"):
            if symu.endswith(".L"):
                sc += 20
        elif ven in ("FR", "EPA", "PA"):
            if symu.endswith(".PA"):
                sc += 20
        elif ven in ("NL", "EAM", "AS", "AMS"):
            if symu.endswith(".AS"):
                sc += 20
        elif ven in ("US", "NY", "NYS", "NAS"):
            if not ("." in symu and not symu.endswith((".BIO", "WS"))):
                if "." not in symu and symu and symu.isalnum() and 1 < len(symu) < 6:
                    sc += 8
        elif ven in ("EU",) and (symu.endswith((".AS", ".DE", ".L", ".PA")) or "." not in symu):
            sc += 2
        else:
            sc += 0
        ccyv = (str(r.get("currency") or "")).upper()
        if ccyv in ("EUR", "GBP", "CHF", "SEK", "NOK", "PLN", "CZK", "HUF", "DKK", "USD"):
            sc += 1
        return sc

    try:
        items = yahoo_search(q, 8)
    except Exception:  # noqa: BLE001
        return None
    cands = [x for x in (items or []) if isinstance(x, dict)]
    if not cands:
        return None
    best = max(cands, key=score_row)
    if not isinstance(best, dict):
        return None
    _t212_yahoo_isin_cache_set(ckey, best)
    return best


def _t212_build_position_row(
    p: dict,
    inst0: dict,
    is_crypto: bool,
    account_ccy: str,
    eur_per: dict[str, float] | None,
) -> dict:
    inst = {**_t212_find_instrument_in_cache(str((inst0 or {}).get("ticker") or "")), **(inst0 or {})}
    ticker = str((inst or {}).get("ticker") or p.get("ticker") or "").strip()
    name = str((inst or {}).get("name") or p.get("name") or "").strip()
    raw_isin = str((inst or {}).get("isin") or "").strip()
    isin = re.sub(r"\s+", "", raw_isin).upper() or None
    qty = _t212_float(p.get("quantity"))
    avg0 = _t212_float(p.get("averagePricePaid") if p.get("averagePricePaid") is not None else p.get("averagePrice"))
    last0 = _t212_float(p.get("currentPrice"))
    acc = (account_ccy or "USD").strip().upper() or "USD"

    if is_crypto:
        ccy = str(
            (inst or {}).get("currencyCode")
            or (inst or {}).get("currency")
            or p.get("currencyCode")
            or "USD"
        )
        ccy = (ccy or "USD").strip().upper() or "USD"
        return {
            "sym": _t212_display_sym(ticker),
            "t212Ticker": ticker,
            "nm": name,
            "ex": _t212_infer_exchange(ticker) or "crypto",
            "ccy": ccy,
            "qty": qty,
            "avg": avg0,
            "last": last0,
            "pfSource": "t212",
        }

    venue = _t212_venue_code_from_ticker(ticker)
    ccy_m = str((inst or {}).get("currencyCode") or (inst or {}).get("currency") or "").strip().upper()
    ccy_venue = _t212_listing_ccy_from_venue(venue)
    yh: dict | None = _t212_yahoo_isin_row(isin, venue) if isin else None
    ysym = str((yh or {}).get("symbol") or "").strip()

    if yh and ysym:
        sym = _t212_display_base_from_yahoo(ysym)
        yccy = str((yh or {}).get("currency") or "").strip().upper()
        ex_row = _t212_yahoo_ex_from_listing_symbol(ysym)
        ccy2 = yccy or ccy_venue or ccy_m or "USD"
    else:
        sym = _t212_display_sym(ticker)
        ccy2 = ccy_venue
        if not ccy2 and ccy_m and len(ccy_m) == 3 and ccy_m.isalpha():
            ccy2 = ccy_m
        ccy2 = ccy2 or acc
        ex_row = (venue and _t212_ex_from_ticker_venue(venue)) or _t212_infer_exchange(ticker) or ""

    ccy2 = (ccy2 or "USD").strip().upper() or "USD"
    if not (yh and ysym) and ccy_venue and ccy2 == acc and ccy_venue != ccy2:
        ccy2 = ccy_venue

    if acc != ccy2:
        avg1 = _t212_fx_convert_amount(avg0, acc, ccy2, eur_per)
        last1 = _t212_fx_convert_amount(last0, acc, ccy2, eur_per)
    else:
        avg1, last1 = avg0, last0

    ex_row = (ex_row or (venue and _t212_ex_from_ticker_venue(venue)) or _t212_infer_exchange(ticker) or "")

    row: dict = {
        "sym": sym,
        "t212Ticker": ticker,
        "nm": name,
        "ex": ex_row,
        "ccy": ccy2,
        "qty": qty,
        "avg": avg1,
        "last": last1,
        "pfSource": "t212",
    }
    if isin:
        row["isin"] = isin
    if acc != ccy2:
        row["t212AccountCcy"] = acc
    return row


class _ApiResponseCache:
    """Thread-safe LRU + TTL for successful JSON payloads (repeat clicks skip upstream)."""

    def __init__(self, max_keys: int = 220) -> None:
        self._max = max_keys
        self._lock = threading.Lock()
        self._store: OrderedDict[str, tuple[float, object]] = OrderedDict()

    def get(self, key: str, now: float) -> object | None:
        with self._lock:
            ent = self._store.get(key)
            if not ent:
                return None
            exp, val = ent
            if now >= exp:
                del self._store[key]
                return None
            self._store.move_to_end(key)
            return copy.deepcopy(val)

    def set(self, key: str, val: object, ttl_s: float, now: float) -> None:
        if ttl_s <= 0:
            return
        with self._lock:
            self._store[key] = (now + ttl_s, copy.deepcopy(val))
            self._store.move_to_end(key)
            while len(self._store) > self._max:
                self._store.popitem(last=False)


_QUOTE_OK_CACHE = _ApiResponseCache()
_HISTORY_OK_CACHE = _ApiResponseCache()
_NEWS_OK_CACHE = _ApiResponseCache()
_CORP_OK_CACHE = _ApiResponseCache()
_FX_OK_CACHE = _ApiResponseCache()


def _quote_cache_ttl_s() -> float:
    """Seconds to reuse identical /api/quote responses; 0 disables. Default 45, max 600."""
    try:
        v = float((env("API_QUOTE_CACHE_SECONDS", "45") or "45").strip())
    except ValueError:
        v = 45.0
    return max(0.0, min(v, 600.0))


def _history_cache_ttl_s() -> float:
    """Seconds to reuse identical /api/history (symbol+exchange+range); 0 disables. Default 180, max 3600."""
    try:
        v = float((env("API_HISTORY_CACHE_SECONDS", "180") or "180").strip())
    except ValueError:
        v = 180.0
    return max(0.0, min(v, 3600.0))


def _history_index_cache_ttl_s() -> float:
    """Extra-long TTL for shared index/ETF history (SPY, ^NSEI, …). Default 7200s, max 86400."""
    try:
        v = float((env("API_HISTORY_BENCHMARK_CACHE_SECONDS", "7200") or "7200").strip())
    except ValueError:
        v = 7200.0
    return max(0.0, min(v, 86400.0))


def _history_effective_cache_ttl_s(sym: str) -> float:
    base = _history_cache_ttl_s()
    if base <= 0:
        return 0.0
    if sym.strip().upper() not in _HISTORY_INDEX_CACHE_SYMBOLS:
        return base
    return max(base, _history_index_cache_ttl_s())


def _quote_cache_key(sym: str, ex: str | None) -> str:
    return f"{sym.strip().upper()}|{(ex or '').strip().upper()}"


def _history_cache_key(sym: str, ex: str | None, rng: str) -> str:
    return f"{sym.strip().upper()}|{(ex or '').strip().upper()}|{rng.strip().lower()}"


def _news_cache_ttl_s() -> float:
    try:
        v = float((env("API_NEWS_CACHE_SECONDS", "120") or "120").strip())
    except ValueError:
        v = 120.0
    return max(0.0, min(v, 900.0))


def _corp_cache_ttl_s() -> float:
    try:
        v = float((env("API_CORPORATE_CACHE_SECONDS", "300") or "300").strip())
    except ValueError:
        v = 300.0
    return max(0.0, min(v, 3600.0))


def _extras_cache_key(sym: str, ex: str | None, suffix: str) -> str:
    return f"{suffix}|{sym.strip().upper()}|{(ex or '').strip().upper()}"


def _fx_cache_ttl_s() -> float:
    """Frankfurter/ECB daily-ish data — cache several hours by default."""
    try:
        v = float((env("API_FX_CACHE_SECONDS", "28800") or "28800").strip())
    except ValueError:
        v = 28800.0
    return max(120.0, min(v, 86400.0))


def _parse_frankfurter_eur_payload(data: object) -> dict:
    if not isinstance(data, dict):
        raise RuntimeError("fx: non-object response")
    base = str(data.get("base") or "").upper()
    if base != "EUR":
        raise RuntimeError(f"fx: unexpected base {base}")
    rates = data.get("rates")
    if not isinstance(rates, dict):
        raise RuntimeError("fx: missing rates")
    date = str(data.get("date") or "")[:10]
    eur_per: dict[str, float] = {"EUR": 1.0}
    for ccy, raw in rates.items():
        c = str(ccy).upper()
        if not c:
            continue
        try:
            f = float(raw)
        except (TypeError, ValueError):
            continue
        if f <= 0:
            continue
        eur_per[c] = 1.0 / f
    return {
        "date": date,
        "eur_per_unit": eur_per,
        "source": "Frankfurter.app (ECB reference rates)",
        "disclaimer": "Indicative mid reference — not executable prices; bank and broker spreads differ.",
    }


def build_fx_eur_payload() -> dict:
    """ECB reference FX via Frankfurter (no API key). `eur_per_unit[USD]` = EUR value of 1 USD."""
    urls = (
        "https://api.frankfurter.app/latest",
        "https://api.frankfurter.dev/latest",
    )
    last_err: Exception | None = None
    for attempt in range(3):
        for url in urls:
            try:
                data = fetch_json(url, timeout_s=18.0)
                return _parse_frankfurter_eur_payload(data)
            except Exception as e:  # noqa: BLE001
                last_err = e
        if attempt < 2:
            time.sleep(0.45 * (attempt + 1))
    raise RuntimeError(f"fx: {last_err}" if last_err else "fx failed")


def _norm_api_path(raw: str) -> str:
    s = raw or "/"
    if len(s) > 1 and s.endswith("/"):
        return s.rstrip("/")
    return s


def _fmp_search_rows(data: object) -> list[dict]:
    """Normalize FMP stable search JSON (list or occasional wrapper dict)."""
    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]
    if isinstance(data, dict):
        if any(
            isinstance(data.get(k), str) and data.get(k)
            for k in ("Error Message", "error", "message", "Error")
        ):
            return []
        for key in ("data", "results", "content", "items"):
            inner = data.get(key)
            if isinstance(inner, list):
                return [x for x in inner if isinstance(x, dict)]
    return []


def _normalize_fmp_search_row(row: dict) -> dict:
    """Map alternate FMP / stable field names onto the keys the UI and merge logic expect."""
    r = dict(row)
    sym = str(r.get("symbol") or r.get("Symbol") or r.get("ticker") or "").strip()
    if sym:
        r["symbol"] = sym
    ex = str(
        r.get("exchangeShortName")
        or r.get("stockExchangeShortName")
        or r.get("stockExchange")
        or r.get("exchange")
        or r.get("Exchange")
        or ""
    ).strip()
    if ex:
        r["exchangeShortName"] = ex
    nm = str(r.get("name") or r.get("companyName") or r.get("Name") or "").strip()
    if nm:
        r["name"] = nm
    qt = str(r.get("quoteType") or r.get("type") or r.get("Type") or "").strip()
    if qt:
        r["quoteType"] = qt
    cur = str(r.get("currency") or r.get("Currency") or "").strip()
    if cur:
        r["currency"] = cur
    return r


# (country name, ISO currency) by Yahoo MIC / exchDisp / FMP exchangeShortName (uppercased keys).
_EXCH_COUNTRY_CCY: dict[str, tuple[str, str]] = {
    # United States
    "NASDAQ": ("United States", "USD"),
    "NYSE": ("United States", "USD"),
    "NYSE ARCA": ("United States", "USD"),
    "NYSE AMERICAN": ("United States", "USD"),
    "AMEX": ("United States", "USD"),
    "NMS": ("United States", "USD"),
    "NGM": ("United States", "USD"),
    "NCM": ("United States", "USD"),
    "NYQ": ("United States", "USD"),
    "PCX": ("United States", "USD"),
    "BATS": ("United States", "USD"),
    "NASDAQGS": ("United States", "USD"),
    "NASDAQ GM": ("United States", "USD"),
    "NASDAQ CM": ("United States", "USD"),
    # India
    "NSE": ("India", "INR"),
    "NSI": ("India", "INR"),
    "BSE": ("India", "INR"),
    "BOM": ("India", "INR"),
    # Europe (selected)
    "LSE": ("United Kingdom", "GBP"),
    "LON": ("United Kingdom", "GBP"),
    "XETRA": ("Germany", "EUR"),
    "GER": ("Germany", "EUR"),
    "ETR": ("Germany", "EUR"),
    "FRA": ("Germany", "EUR"),
    "AMS": ("Netherlands", "EUR"),
    "AS": ("Netherlands", "EUR"),
    "PAR": ("France", "EUR"),
    "EPA": ("France", "EUR"),
    "MIL": ("Italy", "EUR"),
    "BIT": ("Italy", "EUR"),
    "SWX": ("Switzerland", "CHF"),
    "SW": ("Switzerland", "CHF"),
    "STO": ("Sweden", "SEK"),
    "HEL": ("Finland", "EUR"),
    "WSE": ("Poland", "PLN"),
    "EL": ("Greece", "EUR"),
    "ISE": ("Ireland", "EUR"),
    # Americas
    "TSX": ("Canada", "CAD"),
    "TOR": ("Canada", "CAD"),
    "V": ("Canada", "CAD"),
    "CN": ("Canada", "CAD"),
    "SAO": ("Brazil", "BRL"),
    "B3": ("Brazil", "BRL"),
    "MEX": ("Mexico", "MXN"),
    # Asia / Pacific
    "HKG": ("Hong Kong", "HKD"),
    "HKEX": ("Hong Kong", "HKD"),
    "JPX": ("Japan", "JPY"),
    "TYO": ("Japan", "JPY"),
    "KRX": ("South Korea", "KRW"),
    "KOE": ("South Korea", "KRW"),
    "SGX": ("Singapore", "SGD"),
    "SES": ("Singapore", "SGD"),
    "ASX": ("Australia", "AUD"),
    "NZ": ("New Zealand", "NZD"),
    "NZE": ("New Zealand", "NZD"),
    "TW": ("Taiwan", "TWD"),
    "TWO": ("Taiwan", "TWD"),
    "SHH": ("China", "CNY"),
    "SHZ": ("China", "CNY"),
    "SHE": ("China", "CNY"),
    # Other
    "TASE": ("Israel", "ILS"),
    "TLV": ("Israel", "ILS"),
    "JSE": ("South Africa", "ZAR"),
    "JO": ("South Africa", "ZAR"),
}

# Suffix on Yahoo symbol → (country, currency, exchange hint). Longest suffix first.
_SYMBOL_SUFFIX_META: list[tuple[str, str, str, str]] = [
    (".NSE", "India", "INR", "NSE"),
    (".NS", "India", "INR", "NSE"),
    (".TW", "Taiwan", "TWD", "TPE"),
    (".TO", "Canada", "CAD", "TSX"),
    (".KS", "South Korea", "KRW", "KRX"),
    (".BO", "India", "INR", "BSE"),
    (".HK", "Hong Kong", "HKD", "HKEX"),
    (".DE", "Germany", "EUR", "XETRA"),
    (".SW", "Switzerland", "CHF", "SWX"),
    (".PA", "France", "EUR", "EPA"),
    (".AS", "Netherlands", "EUR", "AMS"),
    (".V", "Canada", "CAD", "TSXV"),
    (".AX", "Australia", "AUD", "ASX"),
    (".SI", "Singapore", "SGD", "SGX"),
    (".SA", "Brazil", "BRL", "B3"),
    (".MX", "Mexico", "MXN", "BMV"),
    (".WA", "Poland", "PLN", "WSE"),
    (".ST", "Sweden", "SEK", "STO"),
    (".OL", "Norway", "NOK", "OSE"),
    (".CO", "Denmark", "DKK", "CPH"),
    (".HE", "Finland", "EUR", "HEL"),
    (".MI", "Italy", "EUR", "BIT"),
    (".L", "United Kingdom", "GBP", "LSE"),
    (".T", "Japan", "JPY", "TSE"),
]


def _lookup_country_currency(exchange_label: str, exchange_mic: str) -> tuple[str, str] | None:
    for raw in (exchange_label, exchange_mic):
        u = (raw or "").strip().upper()
        if not u:
            continue
        if u in _EXCH_COUNTRY_CCY:
            return _EXCH_COUNTRY_CCY[u]
        # first segment e.g. "NASDAQ" from "NASDAQ NMS"
        head = u.split()[0] if u else ""
        if head in _EXCH_COUNTRY_CCY:
            return _EXCH_COUNTRY_CCY[head]
    return None


def enrich_search_metadata(row: dict) -> None:
    """Add country + normalize currency / exchange fields for UI filters."""
    sym = str(row.get("symbol") or "").strip().upper()
    ex_label = str(row.get("exchangeShortName") or row.get("exchange") or "").strip()
    ex_mic = str(row.get("exchangeMic") or "").strip()
    if not ex_mic:
        alt = row.get("stockExchange") or row.get("stockExchangeName")
        if alt:
            ex_mic = str(alt).strip()
            row["exchangeMic"] = ex_mic
    cur = str(row.get("currency") or "").strip().upper()

    meta = _lookup_country_currency(ex_label, ex_mic)
    country = str(row.get("country") or "").strip()
    if meta and not country:
        country = meta[0]
        if not cur:
            cur = meta[1]
    if not country or not cur:
        for suf, ctry, ccy, _ in _SYMBOL_SUFFIX_META:
            if sym.endswith(suf.upper()):
                if not country:
                    country = ctry
                if not cur:
                    cur = ccy
                if not ex_label:
                    ex_label = _
                    row["exchangeShortName"] = ex_label
                break

    if country:
        row["country"] = country
    if cur:
        row["currency"] = cur
    if ex_mic and ex_mic not in ex_label:
        row["exchangeMic"] = ex_mic


def _search_row_is_indian(row: dict) -> bool:
    c = str(row.get("country") or "").lower()
    ex = str(row.get("exchangeShortName") or "").upper()
    sym = str(row.get("symbol") or "").upper()
    if "india" in c:
        return True
    if "NSE" in ex or "BSE" in ex or "BOM" in ex or "NATIONAL STOCK" in ex:
        return True
    return any(sym.endswith(s) for s in (".NS", ".NSE", ".BO", ".BSE", ".BOM"))


def _search_row_indian_base(row: dict) -> str | None:
    if not isinstance(row, dict) or not _search_row_is_indian(row):
        return None
    sym = str(row.get("symbol") or "").strip().upper()
    for suf in (".NS", ".NSE", ".BO", ".BSE", ".BOM"):
        if sym.endswith(suf):
            b = sym[: -len(suf)].strip()
            return b or None
    return sym or None


def _search_row_indian_ex_priority(row: dict) -> int:
    """Lower sorts earlier: NSE first, then BSE, then other."""
    sym = str(row.get("symbol") or "").upper()
    ex = str(row.get("exchangeShortName") or "").upper()
    blob = f"{sym} {ex}"
    if re.search(r"\bNSE\b|\.NS|\.NSE|NSI", blob):
        return 0
    if re.search(r"\bBSE\b|\.BO|\.BSE|BOM", blob):
        return 1
    return 2


def _bubble_nse_before_bse_search(rows: list[dict]) -> list[dict]:
    """Stable bubble so NSE rows sit before BSE for the same Indian base ticker (e.g. BAJFINANCE)."""
    out = list(rows)
    n = len(out)
    for _ in range(max(2, n)):
        changed = False
        for i in range(n - 1):
            a, b = out[i], out[i + 1]
            if not isinstance(a, dict) or not isinstance(b, dict):
                continue
            ba, bb = _search_row_indian_base(a), _search_row_indian_base(b)
            if not ba or ba != bb:
                continue
            if _search_row_indian_ex_priority(a) > _search_row_indian_ex_priority(b):
                out[i], out[i + 1] = b, a
                changed = True
        if not changed:
            break
    return out


def fmp_exchange_variants(symbol_base: str, key: str) -> list[dict]:
    """FMP stable ``search-exchange-variants`` — same instrument on multiple venues (e.g. NSE vs BSE)."""
    sym0 = (symbol_base or "").strip().upper()
    if not sym0:
        return []
    sym = sym0.split(".")[0]
    if sym.startswith("0P") or len(sym) < 2:
        return []
    url = "https://financialmodelingprep.com/stable/search-exchange-variants?" + urlencode({"symbol": sym, "apikey": key})
    try:
        data = fetch_json(url, timeout_s=15.0)
    except Exception:  # noqa: BLE001
        return []
    return _fmp_search_rows(data)


def _fmp_indian_equity_bases_missing_nse(merged: list[dict], max_bases: int = 12) -> list[str]:
    """Bases that already have a BSE-class row but no NSE-class row (FMP name search often returns only one)."""
    has_nse: dict[str, bool] = {}
    has_bse: dict[str, bool] = {}
    order: list[str] = []
    for r in merged:
        if not isinstance(r, dict) or not _search_row_is_indian(r):
            continue
        sym = str(r.get("symbol") or "").strip().upper()
        qt = str(r.get("quoteType") or "").upper()
        if sym.startswith("0P") or "MUTUAL" in qt:
            continue
        base = _search_row_indian_base(r)
        if not base:
            continue
        if base not in order:
            order.append(base)
        has_nse.setdefault(base, False)
        has_bse.setdefault(base, False)
        p = _search_row_indian_ex_priority(r)
        if p == 0:
            has_nse[base] = True
        elif p == 1:
            has_bse[base] = True
    out: list[str] = []
    for b in order:
        if has_bse.get(b) and not has_nse.get(b):
            out.append(b)
        if len(out) >= max_bases:
            break
    return out


def augment_fmp_search_with_exchange_variants(merged: list[dict], seen: set[str], fmp_key: str) -> None:
    if not fmp_key or not merged:
        return
    for base in _fmp_indian_equity_bases_missing_nse(merged):
        for row in fmp_exchange_variants(base, fmp_key):
            if not isinstance(row, dict):
                continue
            row = _normalize_fmp_search_row(row)
            sym = str(row.get("symbol", "") or "").strip()
            ex = str(row.get("exchangeShortName", row.get("exchange", "")) or "").strip()
            k = f"{sym}__{ex}"
            if sym and k not in seen:
                seen.add(k)
                merged.append(row)


def finalize_search_results(rows: list[dict]) -> list[dict]:
    for r in rows:
        if isinstance(r, dict):
            enrich_search_metadata(r)
    return _bubble_nse_before_bse_search(rows)


def twelve_symbol_search(q: str, key: str, limit: int) -> list[dict]:
    """Twelve Data symbol_search — used when Yahoo search is rate-limited (429)."""
    lim_n = min(max(limit, 1), 50)
    url = "https://api.twelvedata.com/symbol_search?" + urlencode(
        {"symbol": q.strip(), "apikey": key, "outputsize": str(min(lim_n * 2, 120))}
    )
    try:
        payload = fetch_json(url, timeout_s=18.0)
    except Exception:  # noqa: BLE001
        return []
    if not isinstance(payload, dict) or payload.get("status") == "error":
        return []
    data = payload.get("data")
    if not isinstance(data, list):
        return []
    out: list[dict] = []
    for r in data:
        if not isinstance(r, dict):
            continue
        sym = str(r.get("symbol") or "").strip()
        if not sym:
            continue
        out.append(
            {
                "symbol": sym,
                "name": str(r.get("instrument_name") or r.get("name") or ""),
                "exchangeShortName": str(r.get("exchange") or ""),
                "exchangeMic": str(r.get("mic_code") or r.get("mic") or ""),
                "currency": str(r.get("currency") or ""),
                "country": str(r.get("country") or ""),
                "quoteType": str(r.get("instrument_type") or r.get("type") or ""),
                "_searchProvider": "twelvedata",
            }
        )
        if len(out) >= lim_n:
            break
    return finalize_search_results(out)


def alphavantage_symbol_search(q: str, key: str, limit: int) -> list[dict]:
    """Alpha Vantage SYMBOL_SEARCH — fallback when Yahoo / Twelve search fail or return nothing."""
    lim_n = min(max(limit, 1), 50)
    url = "https://www.alphavantage.co/query?" + urlencode(
        {"function": "SYMBOL_SEARCH", "keywords": q.strip()[:120], "apikey": key}
    )
    try:
        payload = fetch_json(url, timeout_s=25.0)
    except Exception:  # noqa: BLE001
        return []
    if not isinstance(payload, dict):
        return []
    if payload.get("Note") or payload.get("Information") or str(payload.get("Error Message", "") or "").strip():
        return []
    bm = payload.get("bestMatches")
    if not isinstance(bm, list):
        return []
    out: list[dict] = []
    for r in bm[:lim_n]:
        if not isinstance(r, dict):
            continue
        sym = str(r.get("1. symbol") or r.get("1. Symbol") or "").strip()
        if not sym:
            continue
        region = str(r.get("4. region") or r.get("4. Region") or "").strip()
        ctry = region.split("/")[0].strip() if region else ""
        out.append(
            {
                "symbol": sym,
                "name": str(r.get("2. name") or r.get("2. Name") or ""),
                "exchangeShortName": region or str(r.get("3. type") or ""),
                "currency": str(r.get("8. currency") or r.get("8. Currency") or ""),
                "country": ctry,
                "quoteType": str(r.get("3. type") or r.get("3. Type") or ""),
                "_searchProvider": "alphavantage",
            }
        )
    return finalize_search_results(out)


def eodhd_symbol_search(q: str, token: str, limit: int) -> list[dict]:
    """EODHD Search API — strong for company names and international listings (e.g. Reliance on NSE)."""
    lim_n = min(max(limit, 1), 50)
    qn = (q or "").strip()
    if not qn:
        return []
    url = "https://eodhd.com/api/search/" + urlquote(qn, safe="") + "?" + urlencode(
        {"api_token": token, "fmt": "json", "limit": str(min(lim_n * 2, 100))}
    )
    try:
        payload = fetch_json(url, timeout_s=22.0)
    except Exception:  # noqa: BLE001
        return []
    rows: list[dict]
    if isinstance(payload, list):
        rows = [x for x in payload if isinstance(x, dict)]
    elif isinstance(payload, dict) and isinstance(payload.get("Code"), str):
        rows = [payload]
    else:
        return []
    out: list[dict] = []
    for r in rows:
        code = str(r.get("Code") or r.get("code") or "").strip()
        if not code:
            continue
        exch = str(r.get("Exchange") or r.get("exchange") or "").strip()
        name = str(r.get("Name") or r.get("name") or "").strip()
        ctry = str(r.get("Country") or r.get("country") or "").strip()
        ccy = str(r.get("Currency") or r.get("currency") or "").strip()
        typ = str(r.get("Type") or r.get("type") or "").strip()
        out.append(
            {
                "symbol": code,
                "name": name,
                "exchangeShortName": exch,
                "currency": ccy,
                "country": ctry,
                "quoteType": typ,
                "_searchProvider": "eodhd",
            }
        )
        if len(out) >= lim_n:
            break
    return finalize_search_results(out)


def _search_fallback_twelve_av(q: str, lim: int) -> list[dict]:
    """Twelve Data symbol search, then EODHD Search API, then Alpha Vantage SYMBOL_SEARCH (no Yahoo)."""
    td_key = env("TWELVE_DATA_API_KEY")
    if td_key:
        try:
            alt = twelve_symbol_search(q, td_key, lim)
            if alt:
                return alt
        except Exception:  # noqa: BLE001
            pass
    eod_tok = env("EODHD_API_TOKEN") or env("EODHD_API_KEY")
    if eod_tok:
        try:
            alt_e = eodhd_symbol_search(q, eod_tok, lim)
            if alt_e:
                return alt_e
        except Exception:  # noqa: BLE001
            pass
    av_key = env("ALPHAVANTAGE_API_KEY")
    if av_key:
        try:
            alt2 = alphavantage_symbol_search(q, av_key, lim)
            if alt2:
                return alt2
        except Exception:  # noqa: BLE001
            pass
    return []


def search_yahoo_then_fallbacks(q: str, lim: int) -> list[dict]:
    """Yahoo search first; on failure or empty quotes, Twelve → EODHD → Alpha Vantage.

    Yahoo often returns HTTP 429 or an empty ``quotes`` list while still answering 200 — fallbacks
    must run in those cases, not only when an exception is raised.
    """
    err: Exception | None = None
    primary: list[dict] = []
    try:
        primary = yahoo_search(q, lim)
    except Exception as e:  # noqa: BLE001
        err = e
        primary = []
    if primary:
        return primary
    fb = _search_fallback_twelve_av(q, lim)
    if fb:
        return fb
    if err:
        raise err
    return []


def _is_indian_quote_context(symbol: str, exchange: str | None) -> bool:
    su = (symbol or "").strip().upper()
    xu = (exchange or "").strip().upper()
    if su.endswith((".NS", ".NSE", ".BO", ".BSE", ".BOM")):
        return True
    if any(s in xu for s in ("NSE", "BSE", "NSI", "BOM", "NATIONAL STOCK", "BOMBAY", "INDIA")):
        return True
    return False


def _split_yahoo_style_listing(symbol: str) -> tuple[str, str | None]:
    """Strip Yahoo-style suffix; return (base_ticker, NSE|BSE|None)."""
    s = (symbol or "").strip()
    u = s.upper()
    for suf, ex in ((".NS", "NSE"), (".NSE", "NSE"), (".BO", "BSE"), (".BSE", "BSE"), (".BOM", "BSE")):
        if u.endswith(suf):
            return s[: -len(suf)].strip(), ex
    return s, None


def _twelve_exchange_codes(exchange: str | None, inferred: str | None) -> list[str]:
    """Twelve Data ``exchange`` query values to try (NSE/BSE aliases, then literal hints)."""
    parts_raw = [str(x).strip() for x in (exchange, inferred) if x and str(x).strip()]
    # Yahoo search rows sometimes use ``NS``; Twelve expects ``NSE`` / ``NSI``.
    parts = ["NSE" if p.upper() == "NS" else p for p in parts_raw]
    blob = " ".join(p.upper() for p in parts)
    out: list[str] = []
    if re.search(r"\bNSE\b|NATIONAL\s+STOCK|NSI", blob):
        for x in ("NSE", "NSI"):
            if x not in out:
                out.append(x)
    if re.search(r"\bBSE\b|BOMBAY|\bBOM\b", blob):
        for x in ("BSE", "BOM"):
            if x not in out:
                out.append(x)
    for raw in parts:
        compact = re.sub(r"[^A-Z0-9]", "", raw.upper())
        for tok in ("NYSE", "NASDAQ", "AMEX", "NSE", "BSE", "NSI", "BOM", "LSE", "XETRA"):
            if compact == tok or (len(tok) >= 3 and compact.endswith(tok)):
                if tok not in out:
                    out.append(tok)
    return out


def twelve_quote(sym: str, key: str, exchange: str | None) -> dict:
    raw_sym = (sym or "").strip()
    if not raw_sym:
        raise RuntimeError("Twelve: empty symbol")
    base, inferred_ex = _split_yahoo_style_listing(raw_sym)
    codes = _twelve_exchange_codes(exchange, inferred_ex)
    tries: list[dict[str, str]] = []
    for excode in codes:
        tries.append({"symbol": base, "apikey": key, "exchange": excode})
    if exchange and exchange.strip():
        tries.append({"symbol": base, "apikey": key, "exchange": exchange.strip()})
    tries.append({"symbol": base, "apikey": key})
    if raw_sym.upper() != base.upper():
        tries.append({"symbol": raw_sym, "apikey": key})
    seen: set[tuple[str, str]] = set()
    uniq: list[dict[str, str]] = []
    for p in tries:
        k = (p.get("symbol") or "", p.get("exchange") or "")
        if k in seen:
            continue
        seen.add(k)
        uniq.append(p)
    last: Exception | None = None
    for p in uniq:
        url = "https://api.twelvedata.com/quote?" + urlencode(p)
        try:
            payload = fetch_json(url)
            if not isinstance(payload, dict) or payload.get("status") == "error":
                raise RuntimeError(str(payload.get("message") if isinstance(payload, dict) else payload))
            row = {
                "symbol": str(payload.get("symbol") or base),
                "name": str(payload.get("name") or ""),
                "price": payload.get("close") or payload.get("price") or payload.get("last"),
                "previousClose": payload.get("previous_close"),
                "open": payload.get("open"),
                "dayHigh": payload.get("high"),
                "dayLow": payload.get("low"),
                "yearHigh": (payload.get("fifty_two_week") or {}).get("high")
                if isinstance(payload.get("fifty_two_week"), dict)
                else None,
                "yearLow": (payload.get("fifty_two_week") or {}).get("low")
                if isinstance(payload.get("fifty_two_week"), dict)
                else None,
                "volume": payload.get("volume"),
                "exchange": str(payload.get("exchange") or ""),
                "currency": str(payload.get("currency") or ""),
                "timestamp": payload.get("datetime"),
                "_provider": "twelvedata",
                "_note": "Twelve Data",
            }
            for src, dst in (
                ("pe", "peRatio"),
                ("pe_ratio", "peRatio"),
                ("dividend_yield", "dividendYield"),
                ("market_cap", "marketCap"),
                ("market_capitalization", "marketCap"),
                ("beta", "beta"),
                ("change", "changeAmount"),
                ("percent_change", "changePercent"),
                ("average_volume", "averageVolume"),
                ("is_market_open", "marketOpen"),
            ):
                v = payload.get(src)
                if v in (None, ""):
                    continue
                if isinstance(v, (int, float)) and v == v:
                    row[dst] = float(v)
                elif isinstance(v, str) and v.strip():
                    try:
                        row[dst] = float(v.replace(",", ""))
                    except ValueError:
                        row[dst] = v
                elif isinstance(v, bool):
                    row[dst] = v
            if _price_ok(row):
                return row
            last = RuntimeError("Twelve: empty price")
        except Exception as e:  # noqa: BLE001
            last = e
    raise RuntimeError(str(last) if last else "Twelve Data failed")


def twelve_history(symbol: str, exchange: str | None, key: str, range_key: str) -> tuple[list[dict], str]:
    """Daily OHLCV from Twelve Data ``time_series`` (works when Yahoo chart history rate-limits)."""
    rk = range_key if range_key in _HISTORY_RANGES else "1y"
    out_sz = min(_history_max_bars(rk) + 8, 5000)
    raw_sym = (symbol or "").strip()
    if not raw_sym:
        raise RuntimeError("Twelve history: empty symbol")
    base, inferred_ex = _split_yahoo_style_listing(raw_sym)
    codes = _twelve_exchange_codes(exchange, inferred_ex)
    tries: list[dict[str, str]] = []
    for excode in codes:
        tries.append(
            {
                "symbol": base,
                "interval": "1day",
                "outputsize": str(out_sz),
                "apikey": key,
                "exchange": excode,
            }
        )
    if exchange and exchange.strip():
        tries.append(
            {
                "symbol": base,
                "interval": "1day",
                "outputsize": str(out_sz),
                "apikey": key,
                "exchange": exchange.strip(),
            }
        )
    tries.append({"symbol": base, "interval": "1day", "outputsize": str(out_sz), "apikey": key})
    if raw_sym.upper() != base.upper():
        tries.append({"symbol": raw_sym, "interval": "1day", "outputsize": str(out_sz), "apikey": key})
    seen: set[tuple[str, str]] = set()
    uniq: list[dict[str, str]] = []
    for p in tries:
        k = (p.get("symbol") or "", p.get("exchange") or "")
        if k in seen:
            continue
        seen.add(k)
        uniq.append(p)
    last_err: Exception | None = None
    for p in uniq:
        url = "https://api.twelvedata.com/time_series?" + urlencode(p)
        try:
            payload = fetch_json(url, timeout_s=26.0)
        except Exception as e:  # noqa: BLE001
            last_err = e
            continue
        if not isinstance(payload, dict) or payload.get("status") == "error":
            last_err = RuntimeError(str(payload.get("message") if isinstance(payload, dict) else payload))
            continue
        vals = payload.get("values")
        if not isinstance(vals, list) or len(vals) < 2:
            last_err = RuntimeError("no values")
            continue
        bars: list[dict] = []
        for row in reversed(vals):
            if not isinstance(row, dict):
                continue
            ds = str(row.get("datetime") or "")[:10]
            if len(ds) < 10:
                continue
            try:
                tsec = int(datetime.strptime(ds, "%Y-%m-%d").replace(tzinfo=timezone.utc).timestamp())
            except ValueError:
                continue
            c = row.get("close")
            if c in (None, ""):
                continue
            o = row.get("open") if row.get("open") not in (None, "") else c
            h = row.get("high") if row.get("high") not in (None, "") else c
            l = row.get("low") if row.get("low") not in (None, "") else c
            v = row.get("volume")
            bars.append(
                {
                    "t": tsec,
                    "o": float(o),
                    "h": float(h),
                    "l": float(l),
                    "c": float(c),
                    "v": float(v) if v not in (None, "") else None,
                }
            )
        if len(bars) >= 2:
            return bars, "twelvedata"
    raise RuntimeError(f"Twelve history: {last_err}" if last_err else "Twelve history failed")


_EXCH_MAP = {
    "NASDAQ": "US",
    "NYSE": "US",
    "AMEX": "US",
    "NSE": "NSE",
    "NS": "NSE",
    "BSE": "BSE",
    "LSE": "LSE",
    "XETRA": "XETRA",
    "AMS": "AS",
    "SW": "SW",
    "PA": "PA",
    "TSX": "TO",
}


def eodhd_candidates(symbol: str, hint: str | None) -> list[str]:
    """EOD tickers use ``.NSE`` / ``.BSE`` — map Yahoo-style ``.NS`` / ``.BO`` so history/quote URLs resolve."""
    s = (symbol or "").strip().upper()
    if not s:
        return []
    seen: set[str] = set()
    out: list[str] = []

    def add(x: str) -> None:
        x = x.strip().upper()
        if x and x not in seen:
            seen.add(x)
            out.append(x)

    if s.endswith(".BO") or s.endswith(".BOM"):
        base = s.rsplit(".", 1)[0]
        add(f"{base}.BSE")
        add(s)
    elif s.endswith(".NS") or s.endswith(".NSE"):
        base = s.rsplit(".", 1)[0]
        add(f"{base}.NSE")
        add(s)
    elif "." in s:
        add(s)
    else:
        for suf, ex in (
            (".NS", "NSE"),
            (".NSE", "NSE"),
            (".BO", "BSE"),
            (".BSE", "BSE"),
            (".BOM", "BSE"),
            (".L", "LSE"),
            (".DE", "XETRA"),
        ):
            if s.endswith(suf):
                add(s[: -len(suf)] + "." + ex)
                break

    blob = (hint or "").strip().upper()
    h_compact = re.sub(r"[^A-Z0-9]", "", blob)
    ex: str | None = _EXCH_MAP.get(h_compact) or _EXCH_MAP.get((hint or "").strip().upper().replace(" ", ""))
    if not ex:
        if re.search(r"\bBSE\b|BOMBAY|\bBOM\b", blob):
            ex = "BSE"
        elif re.search(r"\bNSE\b|NATIONAL\s+STOCK|NSI", blob):
            ex = "NSE"
    if ex and s.count(".") == 0:
        add(f"{s}.{ex}")
        if ex != "US":
            add(f"{s}.US")
    if s.count(".") == 1:
        base = s.rsplit(".", 1)[0]
        if re.search(r"\bNSE\b|NATIONAL|NSI", blob) and not s.endswith((".NS", ".NSE")):
            add(f"{base}.NSE")
        if re.search(r"\bBSE\b|BOMBAY", blob) and not s.endswith((".BO", ".BOM", ".BSE")):
            add(f"{base}.BSE")
    if "." not in s:
        for t in (f"{s}.NSE", f"{s}.BSE", f"{s}.US", f"{s}.L", f"{s}.XETRA", s):
            add(t)
    return out[:18]


def eodhd_eod_fallback(token: str, sym: str) -> dict:
    to_d = datetime.now(timezone.utc).date()
    fr = to_d - timedelta(days=45)
    params = {"api_token": token, "from": fr.isoformat(), "to": to_d.isoformat(), "fmt": "json", "order": "d"}
    url = f"https://eodhd.com/api/eod/{sym}?{urlencode(params)}"
    data = fetch_json(url, timeout_s=18.0)
    if not isinstance(data, list) or not data:
        raise RuntimeError("EODHD EOD empty")
    bar = data[0] if isinstance(data[0], dict) else {}
    c = bar.get("adjusted_close") or bar.get("close")
    row = {
        "symbol": sym,
        "name": "",
        "price": c,
        "previousClose": None,
        "open": bar.get("open"),
        "dayHigh": bar.get("high"),
        "dayLow": bar.get("low"),
        "yearHigh": None,
        "yearLow": None,
        "volume": bar.get("volume"),
        "exchange": "",
        "currency": "",
        "timestamp": bar.get("date"),
        "_provider": "eodhd",
        "_note": "EODHD daily bar",
    }
    if not _price_ok(row):
        raise RuntimeError("EODHD no close")
    return row


def eodhd_quote(symbol: str, token: str, hint: str | None) -> dict:
    last = None
    for sym in eodhd_candidates(symbol, hint):
        url = f"https://eodhd.com/api/real-time/{sym}?{urlencode({'api_token': token, 'fmt': 'json'})}"
        try:
            payload = fetch_json(url)
            if not isinstance(payload, dict) or payload.get("error"):
                raise RuntimeError(str(payload))
            row = {
                "symbol": str(payload.get("code") or sym),
                "name": "",
                "price": payload.get("close") or payload.get("last") or payload.get("price"),
                "previousClose": payload.get("previousClose") or payload.get("previous_close"),
                "open": payload.get("open"),
                "dayHigh": payload.get("high"),
                "dayLow": payload.get("low"),
                "yearHigh": None,
                "yearLow": None,
                "volume": payload.get("volume"),
                "exchange": str(payload.get("exchange_short_name") or ""),
                "currency": str(payload.get("currency") or ""),
                "timestamp": payload.get("timestamp"),
                "_provider": "eodhd",
                "_note": "EODHD real-time",
            }
            if _price_ok(row):
                return row
            last = RuntimeError("EODHD no price")
        except Exception as e:  # noqa: BLE001
            last = e
    try:
        return eodhd_eod_fallback(token, eodhd_candidates(symbol, hint)[0])
    except Exception as e2:  # noqa: BLE001
        raise RuntimeError(f"EODHD: {last}; {e2}") from last


def eodhd_history(symbol: str, exchange: str | None, token: str, max_bars: int) -> tuple[list[dict], str]:
    """Daily OHLCV from EODHD EOD API for the first symbol candidate that returns data."""
    to_d = datetime.now(timezone.utc).date()
    cal_days = max(160, min(max_bars * 3, 4000))
    fr = to_d - timedelta(days=cal_days)
    params = {"api_token": token, "from": fr.isoformat(), "to": to_d.isoformat(), "fmt": "json", "order": "a"}
    last_err: Exception | None = None
    for eod_sym in eodhd_candidates(symbol, exchange):
        url = f"https://eodhd.com/api/eod/{urlquote(eod_sym, safe='-.')}?{urlencode(params)}"
        try:
            data = fetch_json(url, timeout_s=26.0)
        except Exception as e:  # noqa: BLE001
            last_err = e
            continue
        if not isinstance(data, list) or len(data) < 2:
            last_err = RuntimeError("empty eod")
            continue
        use = data[-max_bars:] if len(data) > max_bars else data
        bars: list[dict] = []
        for row in use:
            if not isinstance(row, dict):
                continue
            ds = str(row.get("date") or "")[:10]
            if len(ds) < 10:
                continue
            try:
                day = datetime.strptime(ds, "%Y-%m-%d").replace(tzinfo=timezone.utc)
                tsec = int(day.timestamp())
            except ValueError:
                continue
            c = row.get("adjusted_close") or row.get("close")
            if c in (None, ""):
                continue
            o = row.get("open") if row.get("open") not in (None, "") else c
            h = row.get("high") if row.get("high") not in (None, "") else c
            l = row.get("low") if row.get("low") not in (None, "") else c
            v = row.get("volume")
            bars.append(
                {
                    "t": tsec,
                    "o": float(o),
                    "h": float(h),
                    "l": float(l),
                    "c": float(c),
                    "v": float(v) if v not in (None, "") else None,
                }
            )
        if len(bars) >= 2:
            return bars, "eodhd"
    raise RuntimeError(f"EODHD history: {last_err}" if last_err else "EODHD history failed")


def marketstack_quote(symbol: str, access: str) -> dict:
    params = {"access_key": access, "symbols": symbol, "limit": "2", "sort": "DESC"}
    url = "https://api.marketstack.com/v1/eod/latest?" + urlencode(params)
    payload = fetch_json(url)
    rows = payload.get("data") if isinstance(payload, dict) else None
    if not isinstance(rows, list) or not rows:
        raise RuntimeError("Marketstack no rows")
    latest = rows[0] if isinstance(rows[0], dict) else {}
    prev = rows[1] if len(rows) > 1 and isinstance(rows[1], dict) else {}
    row = {
        "symbol": symbol,
        "name": str(latest.get("name") or ""),
        "price": latest.get("close"),
        "previousClose": prev.get("close") if prev else None,
        "open": latest.get("open"),
        "dayHigh": latest.get("high"),
        "dayLow": latest.get("low"),
        "yearHigh": None,
        "yearLow": None,
        "volume": latest.get("volume"),
        "exchange": str(latest.get("exchange") or ""),
        "currency": str(latest.get("currency") or ""),
        "timestamp": latest.get("date"),
        "_provider": "marketstack",
        "_note": "Marketstack EOD",
    }
    if not _price_ok(row):
        raise RuntimeError("Marketstack no price")
    return row


def av_quote(symbol: str, key: str) -> dict:
    """Alpha Vantage GLOBAL_QUOTE — same symbol variants as daily history (``.NS`` → ``.NSE``, etc.)."""
    errs: list[str] = []

    def pick(gq: dict, *keys: str) -> str:
        for k in keys:
            if k in gq and gq.get(k) not in (None, ""):
                return str(gq[k])
        return ""

    for sym in _alphavantage_daily_symbol_variants(symbol):
        if not sym:
            continue
        try:
            url = "https://www.alphavantage.co/query?" + urlencode({"function": "GLOBAL_QUOTE", "symbol": sym, "apikey": key})
            payload = fetch_json(url)
            if not isinstance(payload, dict):
                raise RuntimeError("AV bad json")
            if payload.get("Note") or payload.get("Information"):
                raise RuntimeError(str(payload.get("Note") or payload.get("Information")))
            gq = payload.get("Global Quote") or {}
            row = {
                "symbol": pick(gq, "01. symbol", "01. Symbol") or sym,
                "name": "",
                "price": pick(gq, "05. price", "05. Price"),
                "previousClose": pick(gq, "08. previous close", "08. Previous Close"),
                "open": pick(gq, "02. open", "02. Open"),
                "dayHigh": pick(gq, "03. high", "03. High"),
                "dayLow": pick(gq, "04. low", "04. Low"),
                "yearHigh": None,
                "yearLow": None,
                "volume": pick(gq, "06. volume", "06. Volume"),
                "exchange": "",
                "currency": "",
                "timestamp": pick(gq, "07. latest trading day", "07. Latest Trading Day"),
                "_provider": "alphavantage",
                "_note": "Alpha Vantage",
            }
            if not _price_ok(row):
                raise RuntimeError("AV no price")
            return row
        except Exception as e:  # noqa: BLE001
            errs.append(f"{sym}: {e}")
    raise RuntimeError(("; ".join(errs)) if errs else "AV no price")


def _alphavantage_daily_symbol_variants(symbol: str) -> list[str]:
    """Try several AV symbol spellings (Indian Yahoo ``.BO`` vs ``.BSE``, etc.)."""
    s = (symbol or "").strip()
    if not s:
        return []
    out: list[str] = []
    seen: set[str] = set()

    def add(x: str) -> None:
        x = x.strip()
        if x and x.upper() not in seen:
            seen.add(x.upper())
            out.append(x)

    u = s.upper()
    add(s)
    if u.endswith(".BO") or u.endswith(".BOM"):
        base = s.rsplit(".", 1)[0]
        add(f"{base}.BSE")
        add(f"{base}.NSE")
    elif u.endswith(".NS") or u.endswith(".NSE"):
        base = s.rsplit(".", 1)[0]
        add(f"{base}.NSE")
        add(f"{base}.BSE")
    elif "." not in u:
        add(f"{s}.NSE")
        add(f"{s}.BSE")
    return out


def alphavantage_history(symbol: str, key: str, range_key: str) -> tuple[list[dict], str]:
    """Alpha Vantage TIME_SERIES_DAILY — last resort; tries symbol variants for India."""
    rk = range_key if range_key in _HISTORY_RANGES else "1y"
    outsize = "full" if rk in ("2y", "5y", "max") else "compact"
    mb = _history_max_bars(rk)
    errs: list[str] = []
    for sym in _alphavantage_daily_symbol_variants(symbol):
        if not sym:
            continue
        try:
            url = "https://www.alphavantage.co/query?" + urlencode(
                {"function": "TIME_SERIES_DAILY", "symbol": sym, "outputsize": outsize, "apikey": key}
            )
            payload = fetch_json(url, timeout_s=40.0)
            if not isinstance(payload, dict):
                raise RuntimeError("not json")
            if payload.get("Note") or payload.get("Information") or payload.get("Error Message"):
                raise RuntimeError(str(payload.get("Note") or payload.get("Information") or payload.get("Error Message")))
            ts = payload.get("Time Series (Daily)")
            if not isinstance(ts, dict):
                raise RuntimeError("no Time Series (Daily)")
            items = sorted(ts.items(), key=lambda kv: str(kv[0]))
            use = items[-mb:] if len(items) > mb else items
            bars: list[dict] = []
            for dstr, row in use:
                if not isinstance(row, dict):
                    continue
                ds = str(dstr)[:10]
                try:
                    day = datetime.strptime(ds, "%Y-%m-%d").replace(tzinfo=timezone.utc)
                    tsec = int(day.timestamp())
                except ValueError:
                    continue
                c = row.get("4. close")
                if c in (None, ""):
                    continue
                o = row.get("1. open") or c
                h = row.get("2. high") or c
                l = row.get("3. low") or c
                v = row.get("5. volume") or row.get("6. volume")
                bars.append(
                    {
                        "t": tsec,
                        "o": float(o),
                        "h": float(h),
                        "l": float(l),
                        "c": float(c),
                        "v": float(v) if v not in (None, "") else None,
                    }
                )
            if len(bars) >= 2:
                return bars, "alphavantage"
            raise RuntimeError("too few bars")
        except Exception as e:  # noqa: BLE001
            errs.append(f"{sym}: {e}")
    raise RuntimeError("alphavantage: " + ("; ".join(errs) if errs else "no candidates"))


def fmp_quote(symbol: str, key: str) -> list[dict]:
    url = "https://financialmodelingprep.com/stable/quote?" + urlencode({"symbol": symbol, "apikey": key})
    data = fetch_json(url)
    if isinstance(data, list) and data and isinstance(data[0], dict) and _price_ok(data[0]):
        return data
    raise RuntimeError("FMP no usable quote")


def best_quote(symbol: str, exchange: str | None) -> list[dict]:
    fmp_k = env("FMP_API_KEY")
    td = env("TWELVE_DATA_API_KEY")
    eod = env("EODHD_API_TOKEN") or env("EODHD_API_KEY")
    ms = env("MARKETSTACK_ACCESS_KEY") or env("MARKETSTACK_API_KEY")
    av = env("ALPHAVANTAGE_API_KEY")
    fmp_first = _truthy("QUOTE_TRY_FMP_FIRST")
    skip_fmp = _truthy("SKIP_FMP_QUOTE")
    use_yh = _use_yahoo()
    indian = _is_indian_quote_context(symbol, exchange)

    order = (
        ["fmp", "yahoo", "twelve", "eodhd", "marketstack", "alphavantage"]
        if fmp_first
        else ["yahoo", "twelve", "eodhd", "marketstack", "alphavantage", "fmp"]
    )
    if indian:
        # Yahoo is often 429; Twelve + EODHD support NSE/BSE. FMP often 402 on free tier; Marketstack often 406 off‑US.
        order = ["twelve", "eodhd", "yahoo", "alphavantage"]
        if fmp_k and not skip_fmp:
            order.append("fmp")
    if skip_fmp:
        order = [x for x in order if x != "fmp"]
    if not use_yh:
        order = [x for x in order if x != "yahoo"]
    if indian:
        order = [x for x in order if x != "marketstack"]

    errs: list[str] = []
    for step in order:
        try:
            if step == "yahoo":
                return [yahoo_quote(symbol, exchange)]
            if step == "twelve" and td:
                return [twelve_quote(symbol, td, exchange)]
            if step == "eodhd" and eod:
                return [eodhd_quote(symbol, eod, exchange)]
            if step == "marketstack" and ms:
                return [marketstack_quote(symbol, ms)]
            if step == "alphavantage" and av:
                return [av_quote(symbol, av)]
            if step == "fmp" and fmp_k:
                return fmp_quote(symbol, fmp_k)
        except Exception as e:  # noqa: BLE001
            errs.append(f"{step}: {e}")
    raise RuntimeError(" | ".join(errs) if errs else "No quote providers")


def best_history(symbol: str, exchange: str | None, range_key: str) -> dict:
    """Daily bars: Yahoo (if enabled), then EODHD, then Twelve and/or Alpha Vantage.

    For **Indian** listings, Alpha Vantage is tried **before** Twelve Data: Twelve often rejects
    NSE/BSE symbols on free tiers (plan / invalid symbol) while AV may still return daily bars.
    """
    rk = range_key if range_key in _HISTORY_RANGES else "1y"
    mb = _history_max_bars(rk)
    errs: list[str] = []
    if _use_yahoo():
        try:
            bars, src = yahoo_history(symbol, exchange, rk)
            if len(bars) >= 2:
                return {
                    "symbol": symbol.strip(),
                    "exchange": (exchange or "").strip(),
                    "range": rk,
                    "interval": "1d",
                    "bars": bars,
                    "source": src,
                }
        except Exception as e:  # noqa: BLE001
            errs.append(f"yahoo: {e}")
    indian = _is_indian_quote_context(symbol, exchange)
    tok = env("EODHD_API_TOKEN") or env("EODHD_API_KEY")
    if tok:
        try:
            bars, src = eodhd_history(symbol, exchange, tok, mb)
            if len(bars) >= 2:
                return {
                    "symbol": symbol.strip(),
                    "exchange": (exchange or "").strip(),
                    "range": rk,
                    "interval": "1d",
                    "bars": bars,
                    "source": src,
                }
        except Exception as e:  # noqa: BLE001
            errs.append(f"eodhd: {e}")
    td = env("TWELVE_DATA_API_KEY")
    avk = env("ALPHAVANTAGE_API_KEY")
    # Non-India: Twelve before AV (AV is slow / strict on call volume).
    if not indian and td:
        try:
            bars, src = twelve_history(symbol, exchange, td, rk)
            if len(bars) >= 2:
                return {
                    "symbol": symbol.strip(),
                    "exchange": (exchange or "").strip(),
                    "range": rk,
                    "interval": "1d",
                    "bars": bars,
                    "source": src,
                }
        except Exception as e:  # noqa: BLE001
            errs.append(f"twelve: {e}")
    if avk:
        try:
            bars, src = alphavantage_history(symbol, avk, rk)
            if len(bars) >= 2:
                return {
                    "symbol": symbol.strip(),
                    "exchange": (exchange or "").strip(),
                    "range": rk,
                    "interval": "1d",
                    "bars": bars,
                    "source": src,
                }
        except Exception as e:  # noqa: BLE001
            errs.append(f"alphavantage: {e}")
    if indian and td:
        try:
            bars, src = twelve_history(symbol, exchange, td, rk)
            if len(bars) >= 2:
                return {
                    "symbol": symbol.strip(),
                    "exchange": (exchange or "").strip(),
                    "range": rk,
                    "interval": "1d",
                    "bars": bars,
                    "source": src,
                }
        except Exception as e:  # noqa: BLE001
            errs.append(f"twelve: {e}")
    raise RuntimeError(" | ".join(errs) if errs else "No history providers")


def fmp_symbol_try_list(symbol: str) -> list[str]:
    """Symbols to try against FMP company endpoints (US base, Indian suffixes, etc.)."""
    s = (symbol or "").strip()
    if not s:
        return []
    out: list[str] = []
    seen: set[str] = set()

    def add(x: str) -> None:
        t = x.strip()
        if not t:
            return
        k = t.upper()
        if k not in seen:
            seen.add(k)
            out.append(t)

    add(s)
    su = s.upper()
    for suf in (".NS", ".NSE", ".BO", ".BSE", ".BOM", ".L", ".DE", ".US"):
        if su.endswith(suf):
            add(s[: -len(suf)])
            break
    if "." not in su and su.replace("-", "").isalnum():
        add(f"{su}.US")
    return out[:8]


def yahoo_rss_headlines(symbol: str, exchange: str | None, limit: int) -> tuple[list[dict], str]:
    """Yahoo Finance RSS (no API key). Returns (items, yahoo_symbol_used_or_empty)."""
    tries: list[str] = []
    y1 = yahoo_resolve(symbol, exchange)
    add = tries.append
    if y1.strip():
        add(y1.strip())
    s0 = symbol.strip()
    if s0 and s0.upper() != y1.strip().upper():
        add(s0)
    for c in yahoo_chart_candidates(symbol, exchange):
        if c.strip() and c.upper() not in {t.upper() for t in tries}:
            add(c.strip())
        if len(tries) >= 6:
            break
    items: list[dict] = []
    used = ""
    for st in tries:
        if not st:
            continue
        qs = urlencode({"s": st, "region": "US", "lang": "en-US"})
        url = f"https://feeds.finance.yahoo.com/rss/2.0/headline?{qs}"
        try:
            req = Request(url, headers=YAHOO_EXTRA)
            with urlopen(req, timeout=16.0) as resp:
                raw = resp.read()
        except Exception:  # noqa: BLE001
            continue
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            continue
        ch = root.find("channel")
        if ch is None:
            continue
        for it in ch.findall("item"):
            if len(items) >= limit:
                break
            title = (it.findtext("title") or "").strip()
            link = (it.findtext("link") or "").strip()
            pub = (it.findtext("pubDate") or "").strip()
            if not title:
                continue
            items.append(
                {
                    "title": title[:500],
                    "url": link[:2000],
                    "published": pub[:80],
                    "source": "Yahoo RSS",
                }
            )
        if items:
            used = st
            break
    return items[:limit], used


def fmp_stock_news(symbol: str, key: str, limit: int) -> list[dict]:
    out: list[dict] = []
    for fs in fmp_symbol_try_list(symbol):
        url = "https://financialmodelingprep.com/stable/news/stock?" + urlencode(
            {"symbols": fs, "limit": str(min(max(limit, 1), 40)), "apikey": key}
        )
        try:
            data = fetch_json(url, timeout_s=18.0)
        except Exception:  # noqa: BLE001
            continue
        if not isinstance(data, list):
            continue
        for row in data:
            if not isinstance(row, dict):
                continue
            title = str(row.get("title") or row.get("headline") or "").strip()
            link = str(row.get("url") or row.get("link") or "").strip()
            pub = str(row.get("publishedDate") or row.get("date") or "").strip()[:80]
            src = str(row.get("site") or row.get("source") or "FMP").strip()[:80]
            if title:
                out.append({"title": title[:500], "url": link[:2000], "published": pub, "source": src})
            if len(out) >= limit:
                return out[:limit]
        if out:
            return out[:limit]
    return out


def build_news_payload(symbol: str, exchange: str | None, fmp_key: str | None, limit: int) -> dict:
    sources: list[str] = []
    items: list[dict] = []
    if fmp_key:
        try:
            fn = fmp_stock_news(symbol, fmp_key, limit)
            if fn:
                items = fn
                sources.append("fmp")
        except Exception:  # noqa: BLE001
            pass
    if not items:
        y_items, y_sym = yahoo_rss_headlines(symbol, exchange, limit)
        items = y_items
        if y_items:
            sources.append("yahoo_rss")
            if y_sym:
                sources.append(f"yahoo_symbol={y_sym}")
    hint: str | None = None
    if not items:
        hint = "No headlines returned. For Indian tickers try the Yahoo suffix form (e.g. RELIANCE.NS). If Yahoo rate-limits, wait a few minutes or set FMP_API_KEY for FMP news."
    return {
        "items": items,
        "sources": sources,
        "hint": hint,
        "disclaimer": "Third-party headlines for study only — not financial advice or an endorsement.",
    }


def _fmt_fmp_div_rows(rows: list[object], cap: int) -> list[dict]:
    out: list[dict] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        d = str(row.get("date") or row.get("paymentDate") or row.get("recordDate") or "")[:10]
        amt = row.get("adjDividend")
        if amt in (None, ""):
            amt = row.get("dividend")
        out.append(
            {
                "date": d,
                "amount": amt,
                "currency": str(row.get("currency") or "").strip(),
                "recordDate": str(row.get("recordDate") or "")[:10],
                "paymentDate": str(row.get("paymentDate") or "")[:10],
            }
        )
        if len(out) >= cap:
            break
    return out


def _fmt_fmp_split_rows(rows: list[object], cap: int) -> list[dict]:
    out: list[dict] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        d = str(row.get("date") or "")[:10]
        num = row.get("numerator")
        den = row.get("denominator")
        ratio = str(row.get("label") or "").strip()
        if not ratio and num not in (None, "") and den not in (None, ""):
            ratio = f"{num}-for-{den}"
        out.append({"date": d, "ratio": ratio or "—"})
        if len(out) >= cap:
            break
    return out


def _fmt_fmp_earnings_rows(rows: list[object], cap: int) -> list[dict]:
    out: list[dict] = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        d = str(row.get("date") or "")[:10]
        out.append(
            {
                "date": d,
                "epsActual": row.get("eps"),
                "epsEstimated": row.get("epsEstimated"),
                "revenueActual": row.get("revenue"),
                "revenueEstimated": row.get("revenueEstimated"),
                "time": str(row.get("time") or "").strip(),
            }
        )
        if len(out) >= cap:
            break
    return out


def fmp_corporate_bundle(symbol: str, key: str) -> tuple[list[dict], list[dict], list[dict]]:
    divs: list[dict] = []
    spls: list[dict] = []
    earns: list[dict] = []
    for fs in fmp_symbol_try_list(symbol):
        if not divs:
            try:
                url = "https://financialmodelingprep.com/stable/dividends?" + urlencode({"symbol": fs, "apikey": key})
                data = fetch_json(url, timeout_s=18.0)
                if isinstance(data, list) and data:
                    divs = _fmt_fmp_div_rows(sorted(data, key=lambda r: str((r or {}).get("date") or ""), reverse=True), 16)
            except Exception:  # noqa: BLE001
                pass
        if not spls:
            try:
                url = "https://financialmodelingprep.com/stable/splits?" + urlencode({"symbol": fs, "apikey": key})
                data = fetch_json(url, timeout_s=18.0)
                if isinstance(data, list) and data:
                    spls = _fmt_fmp_split_rows(sorted(data, key=lambda r: str((r or {}).get("date") or ""), reverse=True), 12)
            except Exception:  # noqa: BLE001
                pass
        if not earns:
            try:
                url = "https://financialmodelingprep.com/stable/earnings?" + urlencode({"symbol": fs, "apikey": key})
                data = fetch_json(url, timeout_s=18.0)
                if isinstance(data, list) and data:
                    earns = _fmt_fmp_earnings_rows(sorted(data, key=lambda r: str((r or {}).get("date") or ""), reverse=True), 12)
            except Exception:  # noqa: BLE001
                pass
    return divs, spls, earns


def eodhd_div_split_history(symbol: str, exchange: str | None, token: str) -> tuple[list[dict], list[dict]]:
    from_d = (datetime.now(timezone.utc) - timedelta(days=365 * 12)).date().isoformat()
    div_out: list[dict] = []
    spl_out: list[dict] = []
    for eod_sym in eodhd_candidates(symbol, exchange):
        if not div_out:
            try:
                url = f"https://eodhd.com/api/div/{urlquote(eod_sym, safe='-.')}?" + urlencode(
                    {"from": from_d, "api_token": token, "fmt": "json"}
                )
                data = fetch_json(url, timeout_s=18.0)
                if isinstance(data, list) and data:
                    for row in sorted(data, key=lambda r: str((r or {}).get("date") or ""), reverse=True):
                        if not isinstance(row, dict):
                            continue
                        div_out.append(
                            {
                                "date": str(row.get("date") or "")[:10],
                                "amount": row.get("value"),
                                "currency": str(row.get("currency") or ""),
                            }
                        )
                        if len(div_out) >= 16:
                            break
            except Exception:  # noqa: BLE001
                pass
        if not spl_out:
            try:
                url = f"https://eodhd.com/api/splits/{urlquote(eod_sym, safe='-.')}?" + urlencode(
                    {"from": from_d, "api_token": token, "fmt": "json"}
                )
                data = fetch_json(url, timeout_s=18.0)
                if isinstance(data, list) and data:
                    for row in sorted(data, key=lambda r: str((r or {}).get("date") or ""), reverse=True):
                        if not isinstance(row, dict):
                            continue
                        ratio = str(row.get("split") or row.get("option") or "")
                        spl_out.append({"date": str(row.get("date") or "")[:10], "ratio": ratio})
                        if len(spl_out) >= 12:
                            break
            except Exception:  # noqa: BLE001
                pass
        if div_out and spl_out:
            break
    return div_out, spl_out


def build_corporate_payload(symbol: str, exchange: str | None, fmp_key: str | None) -> dict:
    dividends: list[dict] = []
    splits: list[dict] = []
    earnings: list[dict] = []
    sources: list[str] = []
    eod_tok = env("EODHD_API_TOKEN") or env("EODHD_API_KEY")
    if eod_tok:
        ed, es = eodhd_div_split_history(symbol, exchange, eod_tok)
        if ed:
            dividends = ed
            sources.append("eodhd:dividends")
        if es:
            splits = es
            sources.append("eodhd:splits")
    if fmp_key:
        fd, fs, fe = fmp_corporate_bundle(symbol, fmp_key)
        if not dividends and fd:
            dividends = fd
            sources.append("fmp:dividends")
        if not splits and fs:
            splits = fs
            sources.append("fmp:splits")
        if fe:
            earnings = fe
            sources.append("fmp:earnings")
    hint: str | None = None
    if not dividends and not splits and not earnings:
        hint = (
            "No dividend, split, or earnings rows returned. "
            "Set EODHD_API_TOKEN and/or FMP_API_KEY in .env — coverage varies by symbol and listing."
        )
    return {
        "dividends": dividends,
        "splits": splits,
        "earnings": earnings,
        "sources": sources,
        "hint": hint,
        "disclaimer": "Corporate calendar data is informational and may be incomplete; verify dates with your broker or the exchange.",
    }


def _http_post_json(url: str, payload: object, headers: dict[str, str], timeout_s: float = 90.0) -> object:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    h = {
        "User-Agent": "JohnsStockApp/2.0",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    h.update(headers)
    insecure = (env("INSECURE_SSL", "0") or "0").strip() == "1"

    def _once(ctx: ssl.SSLContext) -> object:
        req = Request(url, data=data, headers=h, method="POST")
        with urlopen(req, timeout=timeout_s, context=ctx) as resp:
            return json.loads(resp.read().decode("utf-8", errors="replace"))

    ctx0 = ssl._create_unverified_context() if insecure else ssl.create_default_context()
    try:
        return _once(ctx0)
    except HTTPError as e:
        try:
            body = e.read().decode("utf-8", errors="replace")[:2000]
        except Exception:  # noqa: BLE001
            body = ""
        raise RuntimeError(f"HTTP {e.code} from upstream: {body}") from e
    except ssl.SSLCertVerificationError:
        return _once(ssl._create_unverified_context())
    except Exception as e:  # noqa: BLE001
        reason = getattr(e, "reason", None)
        if isinstance(reason, ssl.SSLCertVerificationError):
            return _once(ssl._create_unverified_context())
        raise


def _sanitize_client_quote(q: object) -> dict:
    """Keep quote JSON small and JSON-serializable for prompts (client-supplied)."""
    if not isinstance(q, dict):
        return {}
    out: dict[str, object] = {}
    for k, v in q.items():
        if k in ("hint", "raw", "debug"):
            continue
        if isinstance(v, (str, int, float, bool)) or v is None:
            out[str(k)] = v
        elif isinstance(v, (list, dict)):
            continue
    return out


def _ai_model_system() -> str:
    """Shared system-style instructions; override with AI_MODEL_SYSTEM in .env."""
    return (env("AI_MODEL_SYSTEM") or "").strip() or (
        "You are a careful financial literacy tutor (not a registered investment adviser). "
        "Use clear markdown. Never invent prices, dates, or filings: only use facts present in the user message. "
        "When a field is missing, say it is missing. Keep answers practical and concise."
    )


def _build_ai_commentary_prompt(sym: str, ex: str | None, quote: dict, technical_summary: str) -> str:
    qj = json.dumps(quote, ensure_ascii=False, indent=2)
    if len(qj) > 10000:
        qj = qj[:10000] + "\n…(truncated)"
    ts = (technical_summary or "").strip() or "(Chart technicals not loaded in the browser yet — use quote fields only.)"
    if len(ts) > 6000:
        ts = ts[:6000] + "…(truncated)"
    core = (
        "Task: write commentary for one instrument the user is viewing in a personal dashboard.\n\n"
        f"Instrument: {sym}\n"
        f"Exchange (if any): {ex or 'n/a'}\n\n"
        "Quote JSON (from their server; may be incomplete, delayed, or from a free data tier):\n"
        f"{qj}\n\n"
        "Rule-based technical summary from the same app (not an exchange-official feed):\n"
        f"{ts}\n\n"
        "Use markdown ## headings in this order:\n"
        "## Snapshot — what the numbers show today\n"
        "## Valuation & balance-sheet hints — only if P/E, P/B, market cap, EPS, or dividend fields exist in the JSON\n"
        "## Price action vs moving-average hints — only if avgPrice50d / avgPrice200d or technical summary mention trends\n"
        "## Liquidity & session context — volume vs average volume if present\n"
        "## Risks & limits — data delays, missing fields, why not to treat this as trading advice\n"
        "End with exactly: **Educational only — not investment advice.**"
    )
    extra = (env("AI_COMMENTARY_INSTRUCTIONS") or "").strip()
    if extra:
        if len(extra) > 3500:
            extra = extra[:3500] + "…"
        core += "\n\nAdditional owner instructions (from .env AI_COMMENTARY_INSTRUCTIONS):\n" + extra
    return core


def _openai_commentary(prompt: str) -> tuple[str, str]:
    key = env("OPENAI_API_KEY")
    if not key:
        raise RuntimeError("missing OPENAI_API_KEY")
    model = env("OPENAI_MODEL", "gpt-4o-mini") or "gpt-4o-mini"
    url = "https://api.openai.com/v1/chat/completions"
    sys = _ai_model_system()
    msgs: list[dict[str, str]] = [{"role": "system", "content": sys}, {"role": "user", "content": prompt}]
    payload = {
        "model": model,
        "temperature": 0.35,
        "messages": msgs,
    }
    data = _http_post_json(url, payload, {"Authorization": f"Bearer {key}"})
    choices = data.get("choices") if isinstance(data, dict) else None
    if not isinstance(choices, list) or not choices:
        raise RuntimeError("OpenAI: empty choices")
    msg = choices[0].get("message") if isinstance(choices[0], dict) else None
    content = msg.get("content") if isinstance(msg, dict) else None
    if not isinstance(content, str) or not content.strip():
        raise RuntimeError("OpenAI: no content")
    return content.strip(), model


def _anthropic_commentary(prompt: str) -> tuple[str, str]:
    key = env("ANTHROPIC_API_KEY")
    if not key:
        raise RuntimeError("missing ANTHROPIC_API_KEY")
    model = env("ANTHROPIC_MODEL", "claude-3-5-haiku-20241022") or "claude-3-5-haiku-20241022"
    url = "https://api.anthropic.com/v1/messages"
    payload = {
        "model": model,
        "max_tokens": 4096,
        "system": _ai_model_system(),
        "messages": [{"role": "user", "content": prompt}],
    }
    headers = {"x-api-key": key, "anthropic-version": "2023-06-01"}
    data = _http_post_json(url, payload, headers)
    blocks = data.get("content") if isinstance(data, dict) else None
    if not isinstance(blocks, list):
        raise RuntimeError("Anthropic: empty content")
    parts: list[str] = []
    for b in blocks:
        if isinstance(b, dict) and b.get("type") == "text" and isinstance(b.get("text"), str):
            parts.append(b["text"])
    text = "\n".join(parts).strip()
    if not text:
        raise RuntimeError("Anthropic: no text blocks")
    return text, model


def _google_gemini_commentary(prompt: str) -> tuple[str, str]:
    key = env("GOOGLE_AI_API_KEY") or env("GEMINI_API_KEY")
    if not key:
        raise RuntimeError("missing GOOGLE_AI_API_KEY (or GEMINI_API_KEY)")
    model = env("GOOGLE_AI_MODEL", "gemini-1.5-flash") or "gemini-1.5-flash"
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        + urlquote(model, safe="")
        + ":generateContent?key="
        + urlquote(key, safe="")
    )
    gem_txt = "System:\n" + _ai_model_system() + "\n\nUser:\n" + prompt
    payload = {"contents": [{"parts": [{"text": gem_txt}]}]}
    data = _http_post_json(url, payload, {})
    cands = data.get("candidates") if isinstance(data, dict) else None
    if not isinstance(cands, list) or not cands:
        raise RuntimeError("Gemini: empty candidates")
    parts = (((cands[0] or {}).get("content") or {}).get("parts")) if isinstance(cands[0], dict) else None
    if not isinstance(parts, list) or not parts:
        raise RuntimeError("Gemini: no parts")
    t0 = parts[0].get("text") if isinstance(parts[0], dict) else None
    if not isinstance(t0, str) or not t0.strip():
        raise RuntimeError("Gemini: no text")
    return t0.strip(), model


def _xai_commentary(prompt: str) -> tuple[str, str]:
    key = env("XAI_API_KEY")
    if not key:
        raise RuntimeError("missing XAI_API_KEY")
    model = env("XAI_MODEL", "grok-2-latest") or "grok-2-latest"
    url = "https://api.x.ai/v1/chat/completions"
    sys = _ai_model_system()
    msgs = [{"role": "system", "content": sys}, {"role": "user", "content": prompt}]
    payload = {
        "model": model,
        "temperature": 0.35,
        "messages": msgs,
    }
    data = _http_post_json(url, payload, {"Authorization": f"Bearer {key}"})
    choices = data.get("choices") if isinstance(data, dict) else None
    if not isinstance(choices, list) or not choices:
        raise RuntimeError("xAI: empty choices")
    msg = choices[0].get("message") if isinstance(choices[0], dict) else None
    content = msg.get("content") if isinstance(msg, dict) else None
    if not isinstance(content, str) or not content.strip():
        raise RuntimeError("xAI: no content")
    return content.strip(), model


def _pick_ai_commentary_backend():
    """Return (label, fn) where fn(prompt) -> (text, model), or None if no key is set."""
    prov = (env("AI_COMMENTARY_PROVIDER") or "auto").strip().lower()
    if prov == "openai":
        order = ["openai"]
    elif prov == "anthropic":
        order = ["anthropic"]
    elif prov == "google":
        order = ["google"]
    elif prov == "gemini":
        order = ["google"]
    elif prov == "xai":
        order = ["xai"]
    else:
        order = ["openai", "anthropic", "google", "xai"]
    for p in order:
        if p == "openai" and env("OPENAI_API_KEY"):
            return ("openai", _openai_commentary)
        if p == "anthropic" and env("ANTHROPIC_API_KEY"):
            return ("anthropic", _anthropic_commentary)
        if p == "google" and (env("GOOGLE_AI_API_KEY") or env("GEMINI_API_KEY")):
            return ("google", _google_gemini_commentary)
        if p == "xai" and env("XAI_API_KEY"):
            return ("xai", _xai_commentary)
    return None


def handle_ai_commentary_post(handler: SimpleHTTPRequestHandler, body: dict) -> None:
    load_dotenv(os.path.join(APP_ROOT, ".env"))
    sym = str(body.get("symbol") or "").strip()
    ex_raw = body.get("exchange")
    ex = str(ex_raw).strip() if ex_raw not in (None, "") else None
    if not sym:
        return json_response(handler, 400, {"error": "missing symbol"})
    quote_in = _sanitize_client_quote(body.get("quote"))
    technical_summary = str(body.get("technical_summary") or "")
    if not quote_in:
        try:
            quote_in = best_quote(sym, ex)
            if isinstance(quote_in, list):
                quote_in = quote_in[0] if quote_in and isinstance(quote_in[0], dict) else {}
            if not isinstance(quote_in, dict):
                quote_in = {}
            quote_in = _sanitize_client_quote(quote_in)
        except Exception as e:  # noqa: BLE001
            return json_response(
                handler,
                400,
                {
                    "error": "missing quote",
                    "detail": redact_for_json_detail(str(e)),
                    "hint": "Open Overview first, or include a `quote` object in the JSON body.",
                },
            )
    picked = _pick_ai_commentary_backend()
    if picked is None:
        return json_response(
            handler,
            503,
            {
                "error": "no llm configured",
                "detail": "Set one of OPENAI_API_KEY, ANTHROPIC_API_KEY, GOOGLE_AI_API_KEY (or GEMINI_API_KEY), or XAI_API_KEY in .env, then restart server.py.",
                "hint": "Optional: AI_COMMENTARY_PROVIDER=openai|anthropic|google|xai|auto (default auto picks first available).",
            },
        )
    label, fn = picked
    prompt = _build_ai_commentary_prompt(sym, ex, quote_in, technical_summary)
    try:
        text, model = fn(prompt)
    except Exception as e:  # noqa: BLE001
        raw = str(e)
        return json_response(
            handler,
            502,
            {
                "error": "ai commentary failed",
                "detail": redact_for_json_detail(raw),
                "hint": "Check model name env vars, billing, and upstream status. Keys are never returned to the browser.",
            },
        )
    return json_response(
        handler,
        200,
        {"ok": True, "provider": label, "model": model, "text": text},
    )


def _build_ai_ask_prompt(question: str) -> str:
    q = (question or "").strip()
    if len(q) > 3200:
        q = q[:3200] + "…"
    core = (
        "Task: answer a general investing / markets literacy question.\n\n"
        f"Question:\n{q}\n\n"
        "Use ## markdown headings when it helps. If the user needs live prices, filings, or news you do not have, "
        "say what is missing and how they could look it up (e.g. exchange filings, broker statement) without naming paywalled scrapers. "
        "Do not invent numbers. End with: **Educational only — not investment advice.**"
    )
    extra = (env("AI_ASK_INSTRUCTIONS") or "").strip()
    if extra:
        if len(extra) > 3500:
            extra = extra[:3500] + "…"
        core += "\n\nAdditional owner instructions (from .env AI_ASK_INSTRUCTIONS):\n" + extra
    return core


def handle_ai_ask(handler: SimpleHTTPRequestHandler, question: str) -> None:
    load_dotenv(os.path.join(APP_ROOT, ".env"))
    q = (question or "").strip()
    if not q:
        return json_response(
            handler,
            400,
            {"error": "missing question", "hint": "Use GET /api/ai-ask?q=your+question (max ~3k chars)."},
        )
    picked = _pick_ai_commentary_backend()
    if picked is None:
        return json_response(
            handler,
            503,
            {
                "error": "no llm configured",
                "detail": "Set OPENAI_API_KEY, ANTHROPIC_API_KEY, GOOGLE_AI_API_KEY (or GEMINI_API_KEY), or XAI_API_KEY in .env, then restart server.py.",
            },
        )
    label, fn = picked
    prompt = _build_ai_ask_prompt(q)
    try:
        text, model = fn(prompt)
    except Exception as e:  # noqa: BLE001
        raw = str(e)
        return json_response(
            handler,
            502,
            {
                "error": "ai ask failed",
                "detail": redact_for_json_detail(raw),
                "hint": "Check API key, model env vars, and upstream status.",
            },
        )
    return json_response(handler, 200, {"ok": True, "provider": label, "model": model, "text": text})


class Handler(SimpleHTTPRequestHandler):
    def end_headers(self) -> None:
        """Prevent Safari / PWA from holding stale HTML, JS, or CSS when the project updates."""
        try:
            path_only = urlparse(self.path).path.lower()
        except Exception:
            path_only = ""
        rel = path_only.lstrip("/")
        if rel.endswith((".html", ".htm", ".js", ".css", ".webmanifest")) or rel == "service-worker.js":
            self.send_header("Cache-Control", "no-store, must-revalidate, max-age=0")
        super().end_headers()

    def do_GET(self) -> None:  # noqa: N802
        p = urlparse(self.path)
        np = _norm_api_path(p.path).lower()
        if np.startswith("/api/") or np == "/api":
            self._api(p)
            return
        super().do_GET()

    def do_PUT(self) -> None:  # noqa: N802
        p = urlparse(self.path)
        path = _norm_api_path(p.path).lower()
        load_dotenv(os.path.join(APP_ROOT, ".env"))
        if path == "/api/shared/portfolio":
            return handle_put_shared_portfolio(self)
        json_response(self, 404, {"error": "unknown put route", "detail": path})

    def do_POST(self) -> None:  # noqa: N802
        p = urlparse(self.path)
        path = _norm_api_path(p.path).lower()
        if path != "/api/ai-commentary":
            json_response(self, 404, {"error": "unknown post route", "detail": path})
            return
        load_dotenv(os.path.join(APP_ROOT, ".env"))
        try:
            n = int(self.headers.get("Content-Length") or "0")
        except ValueError:
            n = 0
        raw = self.rfile.read(n) if n > 0 else b"{}"
        try:
            body = json.loads(raw.decode("utf-8", errors="replace"))
        except json.JSONDecodeError:
            json_response(self, 400, {"error": "invalid json body"})
            return
        if not isinstance(body, dict):
            json_response(self, 400, {"error": "json body must be an object"})
            return
        handle_ai_commentary_post(self, body)

    def _api(self, p) -> None:
        load_dotenv(os.path.join(APP_ROOT, ".env"))
        path = _norm_api_path(p.path).lower()
        if path == "/api/fmp/search":
            path = "/api/search"
        elif path == "/api/fmp/quote":
            path = "/api/quote"
        qs = parse_qs(p.query)
        fmp_key = env("FMP_API_KEY")

        if path == "/api/ai-commentary":
            sym = (qs.get("symbol", [""])[0] or "").strip()
            ex_raw = (qs.get("exchange", [""])[0] or "").strip()
            tech = (qs.get("technical_summary", [""])[0] or "")
            if len(tech) > 4500:
                tech = tech[:4500] + "…"
            body = {"symbol": sym, "exchange": ex_raw, "technical_summary": tech, "quote": {}}
            return handle_ai_commentary_post(self, body)

        if path == "/api/ai-ask":
            qn = (qs.get("q", [""])[0] or "").strip() or (qs.get("question", [""])[0] or "").strip()
            return handle_ai_ask(self, qn)

        if path == "/api/shared/portfolio":
            return handle_get_shared_portfolio(self, parse_qs(p.query or ""))

        if path == "/api/health":
            td = bool(env("TWELVE_DATA_API_KEY"))
            eod = bool(env("EODHD_API_TOKEN") or env("EODHD_API_KEY"))
            ms = bool(env("MARKETSTACK_ACCESS_KEY") or env("MARKETSTACK_API_KEY"))
            av = bool(env("ALPHAVANTAGE_API_KEY"))
            t212 = bool(env("TRADING212_API_KEY") and env("TRADING212_API_SECRET"))
            read_tok = bool((env("SHARED_PORTFOLIO_READ_TOKEN", "") or "").strip())
            write_tok = bool((env("SHARED_PORTFOLIO_WRITE_TOKEN", "") or "").strip())
            snap_path = _shared_portfolio_file_path()
            has_snap = os.path.isfile(snap_path)
            t212_cache: dict = {}
            if t212:
                with T212I_LOCK:
                    t212_cache = {
                        "n_rows": len(T212I_ITEMS),
                        "pages": T212I_STATUS.get("pages", 0),
                        "cache_complete": bool(T212I_STATUS.get("complete")),
                        "cache_loading": bool(T212I_STATUS.get("loading")),
                    }
                    if T212I_STATUS.get("err"):
                        t212_cache["cache_error"] = "yes"
            return json_response(
                self,
                200,
                {
                    "ok": True,
                    "search": "fmp" if fmp_key else "yahoo",
                    "yahoo_quotes": _use_yahoo(),
                    "configured": {
                        "fmp": bool(fmp_key),
                        "twelve_data": td,
                        "eodhd": eod,
                        "marketstack": ms,
                        "alphavantage": av,
                    },
                    "quote_order_default": ["yahoo", "twelve", "eodhd", "marketstack", "alphavantage", "fmp"],
                    "search_yahoo_429_fallback_twelve": bool(td),
                    "search_fallback_eodhd": eod,
                    "search_fallback_alphavantage": bool(env("ALPHAVANTAGE_API_KEY")),
                    "history": {
                        "path": "/api/history",
                        "ranges": sorted(_HISTORY_RANGES),
                        "providers_try": ["yahoo", "eodhd", "twelve_data", "alphavantage"],
                        "providers_try_indian": ["yahoo", "eodhd", "alphavantage", "twelve_data"],
                    },
                    "response_cache": {
                        "quote_ttl_seconds": _quote_cache_ttl_s(),
                        "history_ttl_seconds": _history_cache_ttl_s(),
                        "history_benchmark_ttl_seconds": _history_index_cache_ttl_s(),
                        "news_ttl_seconds": _news_cache_ttl_s(),
                        "corporate_ttl_seconds": _corp_cache_ttl_s(),
                        "fx_eur_ttl_seconds": _fx_cache_ttl_s(),
                    },
                    "instrument_extras": {
                        "news": "/api/news?symbol=…&exchange=…&limit=18",
                        "corporate": "/api/corporate?symbol=…&exchange=…",
                    },
                    "fx": {"eur_reference": "/api/fx-eur"},
                    "llm_commentary": {
                        "commentary": "GET /api/ai-commentary?symbol=…&exchange=…&technical_summary=… (preferred) or POST JSON same fields",
                        "ask": "GET /api/ai-ask?q=… (general investing / literacy questions)",
                        "configured": {
                            "openai": bool(env("OPENAI_API_KEY")),
                            "anthropic": bool(env("ANTHROPIC_API_KEY")),
                            "google": bool(env("GOOGLE_AI_API_KEY") or env("GEMINI_API_KEY")),
                            "xai": bool(env("XAI_API_KEY")),
                        },
                    },
                    "trading212": {
                        "configured": t212,
                        "read_only_positions": "GET /api/t212/rows (splits stock-style vs crypto by instrument type)",
                        "instruments": "GET /api/t212/instruments?q=…&region=eu|all&type=STOCK|ETF (cached; warms in background, ~1 call / 50s)",
                    },
                    "t212_instruments_cache": t212_cache,
                    "shared_family_portfolio": {
                        "get": "GET /api/shared/portfolio?token=READ_TOKEN (v2 bundle for family read-only view)",
                        "put": "PUT /api/shared/portfolio + header X-Portfolio-Write-Key: WRITE_TOKEN",
                        "read_token_configured": read_tok,
                        "write_token_configured": write_tok,
                        "snapshot_on_disk": has_snap,
                    },
                    "api_revision": 16,
                },
            )

        if path == "/api/lan-urls":
            port = int(env("PORT", "8844") or "8844")
            ips = collect_lan_ipv4s()
            urls = [f"http://{ip}:{port}" for ip in ips]
            return json_response(
                self,
                200,
                {
                    "ok": True,
                    "port": port,
                    "ips": ips,
                    "urls": urls,
                    "hint": "Same Wi-Fi as this Mac. Copy a url into iPhone Safari. If urls is empty, use System Settings → Network → Wi-Fi → IP address.",
                },
            )

        if path in ("/api/t212/rows", "/api/t212/positions"):
            return handle_t212_rows(self)
        if path == "/api/t212/instruments":
            return handle_t212_instruments(self, qs)

        if path == "/api/search":
            q = (qs.get("q", [""])[0] or "").strip()
            lim = int(qs.get("limit", ["20"])[0] or "20") if (qs.get("limit", ["20"])[0] or "20").isdigit() else 20
            if not q:
                return json_response(self, 200, [])
            if fmp_key:
                params = {"query": q, "limit": str(lim), "apikey": fmp_key}
                merged: list[dict] = []
                seen: set[str] = set()
                for fmp_ep in ("search-name", "search-symbol"):
                    url = f"https://financialmodelingprep.com/stable/{fmp_ep}?" + urlencode(params)
                    try:
                        data = fetch_json(url)
                    except Exception:  # noqa: BLE001
                        data = []
                    for row in _fmp_search_rows(data):
                        if not isinstance(row, dict):
                            continue
                        row = _normalize_fmp_search_row(row)
                        sym = str(row.get("symbol", "") or "").strip()
                        ex = str(row.get("exchangeShortName", row.get("exchange", "")) or "").strip()
                        k = f"{sym}__{ex}"
                        if sym and k not in seen:
                            seen.add(k)
                            merged.append(row)
                # Blend Yahoo when few hits, or always when FMP returned nothing (search-only rescue).
                yahoo_blend = len(merged) < 10 and (_use_yahoo() or len(merged) == 0)
                if yahoo_blend:
                    y_rows: list[dict] = []
                    try:
                        y_rows = list(yahoo_search(q, min(lim, 20)))
                    except Exception:  # noqa: BLE001
                        y_rows = []
                    if not y_rows:
                        y_rows = list(_search_fallback_twelve_av(q, min(lim, 20)))
                    for row in y_rows:
                        if not isinstance(row, dict):
                            continue
                        sym = str(row.get("symbol", "") or "").strip()
                        ex = str(row.get("exchangeShortName", row.get("exchange", "")) or "").strip()
                        k = f"{sym}__{ex}"
                        if sym and k not in seen:
                            seen.add(k)
                            merged.append(row)
                if not merged:
                    for row in _search_fallback_twelve_av(q, lim):
                        if not isinstance(row, dict):
                            continue
                        sym = str(row.get("symbol", "") or "").strip()
                        ex = str(row.get("exchangeShortName", row.get("exchange", "")) or "").strip()
                        k = f"{sym}__{ex}"
                        if sym and k not in seen:
                            seen.add(k)
                            merged.append(row)
                augment_fmp_search_with_exchange_variants(merged, seen, fmp_key)
                finalize_search_results(merged)
                return json_response(self, 200, merged)
            try:
                return json_response(self, 200, search_yahoo_then_fallbacks(q, lim))
            except Exception as e:  # noqa: BLE001
                msg = str(e)
                rate_like = "429" in msg or "503" in msg or "Too Many Requests" in msg
                hint = (
                    "Search failed on Yahoo, Twelve Data, EODHD, and Alpha Vantage (where keys are set). Wait a few minutes, "
                    "try a ticker (e.g. NVDA), and check /api/health. Free tiers have strict daily limits."
                    if rate_like
                    else "Search failed. Check /api/health for keys, wait a moment, or try a ticker symbol."
                )
                return json_response(
                    self, 502, {"error": "search failed", "detail": redact_for_json_detail(msg), "hint": hint}
                )

        if path == "/api/quote":
            sym = (qs.get("symbol", [""])[0] or "").strip()
            ex = (qs.get("exchange", [""])[0] or "").strip() or None
            if not sym:
                return json_response(self, 400, {"error": "missing symbol"})
            now = time.monotonic()
            q_ttl = _quote_cache_ttl_s()
            if q_ttl > 0:
                hit = _QUOTE_OK_CACHE.get(_quote_cache_key(sym, ex), now)
                if hit is not None:
                    return json_response(self, 200, hit)
            try:
                out = best_quote(sym, ex)
            except Exception as e:  # noqa: BLE001
                raw = str(e)
                detail = redact_for_json_detail(raw)
                payload: dict[str, object] = {"error": "quote failed", "detail": detail}
                rh = _rate_limit_hint_from_message(raw)
                if rh:
                    payload["hint"] = rh
                return json_response(self, 502, payload)
            if q_ttl > 0:
                _QUOTE_OK_CACHE.set(_quote_cache_key(sym, ex), out, q_ttl, now)
            return json_response(self, 200, out)

        if path == "/api/news":
            sym = (qs.get("symbol", [""])[0] or "").strip()
            ex = (qs.get("exchange", [""])[0] or "").strip() or None
            if not sym:
                return json_response(self, 400, {"error": "missing symbol"})
            lim_raw = (qs.get("limit", ["18"])[0] or "18").strip()
            lim = int(lim_raw) if lim_raw.isdigit() else 18
            lim = min(max(lim, 1), 40)
            now = time.monotonic()
            n_ttl = _news_cache_ttl_s()
            n_key = _extras_cache_key(sym, ex, f"news|{lim}")
            if n_ttl > 0:
                hit = _NEWS_OK_CACHE.get(n_key, now)
                if hit is not None:
                    return json_response(self, 200, hit)
            try:
                out = build_news_payload(sym, ex, fmp_key, lim)
            except Exception as e:  # noqa: BLE001
                raw = str(e)
                pl: dict[str, object] = {"error": "news failed", "detail": redact_for_json_detail(raw)}
                rh = _rate_limit_hint_from_message(raw)
                if rh:
                    pl["hint"] = rh
                return json_response(self, 502, pl)
            if n_ttl > 0:
                _NEWS_OK_CACHE.set(n_key, out, n_ttl, now)
            return json_response(self, 200, out)

        if path == "/api/corporate":
            sym = (qs.get("symbol", [""])[0] or "").strip()
            ex = (qs.get("exchange", [""])[0] or "").strip() or None
            if not sym:
                return json_response(self, 400, {"error": "missing symbol"})
            now = time.monotonic()
            c_ttl = _corp_cache_ttl_s()
            c_key = _extras_cache_key(sym, ex, "corp")
            if c_ttl > 0:
                hit = _CORP_OK_CACHE.get(c_key, now)
                if hit is not None:
                    return json_response(self, 200, hit)
            try:
                out = build_corporate_payload(sym, ex, fmp_key)
            except Exception as e:  # noqa: BLE001
                raw = str(e)
                pl2: dict[str, object] = {"error": "corporate failed", "detail": redact_for_json_detail(raw)}
                rh2 = _rate_limit_hint_from_message(raw)
                if rh2:
                    pl2["hint"] = rh2
                return json_response(self, 502, pl2)
            if c_ttl > 0:
                _CORP_OK_CACHE.set(c_key, out, c_ttl, now)
            return json_response(self, 200, out)

        if path == "/api/fx-eur":
            now = time.monotonic()
            f_ttl = _fx_cache_ttl_s()
            f_key = "fx|eur|latest"
            if f_ttl > 0:
                hit = _FX_OK_CACHE.get(f_key, now)
                if hit is not None:
                    return json_response(self, 200, hit)
            try:
                out = build_fx_eur_payload()
            except Exception as e:  # noqa: BLE001
                raw = str(e)
                plf: dict[str, object] = {"error": "fx failed", "detail": redact_for_json_detail(raw)}
                rhf = _rate_limit_hint_from_message(raw)
                if rhf:
                    plf["hint"] = rhf
                return json_response(self, 502, plf)
            if f_ttl > 0:
                _FX_OK_CACHE.set(f_key, out, f_ttl, now)
            return json_response(self, 200, out)

        if path == "/api/history":
            sym = (qs.get("symbol", [""])[0] or "").strip()
            ex = (qs.get("exchange", [""])[0] or "").strip() or None
            rng_raw = (qs.get("range", ["1y"])[0] or "1y").strip().lower()
            rng = rng_raw if rng_raw in _HISTORY_RANGES else "1y"
            if not sym:
                return json_response(self, 400, {"error": "missing symbol"})
            now = time.monotonic()
            h_ttl = _history_effective_cache_ttl_s(sym)
            if h_ttl > 0:
                hit = _HISTORY_OK_CACHE.get(_history_cache_key(sym, ex, rng), now)
                if hit is not None:
                    return json_response(self, 200, hit)
            try:
                out = best_history(sym, ex, rng)
            except Exception as e:  # noqa: BLE001
                raw = str(e)
                detail = redact_for_json_detail(raw)
                hist_hint = (
                    "Try another range (5d,1mo,3mo,6mo,1y,2y,5y,max). India: Yahoo → EODHD → Alpha Vantage → Twelve. "
                    "Else: Yahoo → EODHD → Twelve → AV. If you see 429/402 from every provider, wait several minutes "
                    "or raise API_HISTORY_CACHE_SECONDS in .env so successful responses are reused longer. "
                    "Restart if api_revision < 13."
                )
                rh = _rate_limit_hint_from_message(raw)
                payload: dict[str, object] = {
                    "error": "history failed",
                    "detail": detail,
                    "hint": f"{hist_hint} — {rh}" if rh else hist_hint,
                }
                return json_response(self, 502, payload)
            if h_ttl > 0:
                _HISTORY_OK_CACHE.set(_history_cache_key(sym, ex, rng), out, h_ttl, now)
            return json_response(self, 200, out)

        return json_response(
            self,
            404,
            {
                "error": "unknown route",
                "detail": path,
                "hint": "Use the URL printed when you start python3 server.py. Port 5174 is often taken by Vite — set PORT=8844 in .env. Routes: /api/search /api/quote /api/history /api/news /api/corporate /api/fx-eur /api/health GET /api/ai-commentary GET /api/ai-ask (POST /api/ai-commentary also supported)",
            },
        )


# --- Shared family portfolio (read-only link) — server snapshot; not the same as per-browser localStorage. ---
_SHARED_PF_LOCK = threading.Lock()


def _shared_portfolio_file_path() -> str:
    p = (env("SHARED_PORTFOLIO_PATH", "") or "").strip()
    if p:
        return p
    return os.path.join(APP_ROOT, "data", "shared_portfolio.json")


def handle_get_shared_portfolio(handler: SimpleHTTPRequestHandler, qs: dict[str, list[str]]) -> None:
    want = (env("SHARED_PORTFOLIO_READ_TOKEN", "") or "").strip()
    if not want:
        return json_response(
            handler,
            503,
            {
                "ok": False,
                "error": "shared_view_not_configured",
                "detail": "Set SHARED_PORTFOLIO_READ_TOKEN in .env on the server, then redeploy.",
            },
        )
    got = (qs.get("token", [""])[0] or qs.get("read", [""])[0] or "").strip()
    if not got or got != want:
        return json_response(handler, 403, {"ok": False, "error": "forbidden"})
    path = _shared_portfolio_file_path()
    with _SHARED_PF_LOCK:
        if not os.path.isfile(path):
            return json_response(
                handler,
                404,
                {
                    "ok": False,
                    "error": "not_published",
                    "detail": "Owner has not published a snapshot yet (PUT /api/shared/portfolio with write key).",
                },
            )
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:  # noqa: BLE001
            return json_response(
                handler,
                500,
                {"ok": False, "error": "read_failed", "detail": redact_for_json_detail(str(e))[:400]},
            )
    if not isinstance(data, dict) or data.get("v") != 2 or not isinstance(data.get("brokers"), dict):
        return json_response(
            handler,
            500,
            {"ok": False, "error": "invalid_snapshot", "detail": "Stored JSON is not a v2 portfolio bundle."},
        )
    meta = data.get("_meta") if isinstance(data.get("_meta"), dict) else {}
    out = {k: v for k, v in data.items() if k != "_meta"}
    updated = str(meta.get("updated_at") or "")[:32]
    return json_response(
        handler,
        200,
        {"ok": True, "bundle": out, "updated_at": updated, "_note": "read-only; owner publishes via PUT /api/shared/portfolio"},
    )


def handle_put_shared_portfolio(handler: SimpleHTTPRequestHandler) -> None:
    wkey = (env("SHARED_PORTFOLIO_WRITE_TOKEN", "") or "").strip()
    if not wkey:
        return json_response(
            handler,
            503,
            {
                "ok": False,
                "error": "publish_not_configured",
                "detail": "Set SHARED_PORTFOLIO_WRITE_TOKEN in .env; keep it secret. Redeploy.",
            },
        )
    got = (handler.headers.get("X-Portfolio-Write-Key") or handler.headers.get("X-Write-Key") or "").strip()
    if not got or got != wkey:
        return json_response(handler, 403, {"ok": False, "error": "forbidden"})
    try:
        n = int(handler.headers.get("Content-Length") or "0")
    except ValueError:
        n = 0
    raw = handler.rfile.read(n) if n > 0 else b""
    try:
        body = json.loads(raw.decode("utf-8", errors="replace"))
    except json.JSONDecodeError:
        return json_response(handler, 400, {"ok": False, "error": "invalid json"})
    if not isinstance(body, dict) or body.get("v") != 2 or not isinstance(body.get("brokers"), dict):
        return json_response(
            handler,
            400,
            {
                "ok": False,
                "error": "expected_v2_portfolio",
                "detail": "Body must be { v: 2, brokers: { ... } } matching the app backup shape.",
            },
        )
    body = dict(body)
    body.pop("_meta", None)
    body["_meta"] = {
        "updated_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "source": "PUT /api/shared/portfolio",
    }
    path = _shared_portfolio_file_path()
    ddir = os.path.dirname(path)
    try:
        os.makedirs(ddir, exist_ok=True)
    except OSError as e:
        return json_response(
            handler,
            500,
            {"ok": False, "error": "mkdir_failed", "detail": str(e)[:200]},
        )
    tmp = path + ".tmp"
    payload = json.dumps(body, ensure_ascii=False, indent=0).encode("utf-8")
    with _SHARED_PF_LOCK:
        with open(tmp, "wb") as f:
            f.write(payload)
        os.replace(tmp, path)
    return json_response(
        handler,
        200,
        {"ok": True, "bytes": len(payload), "path_hint": ddir, "updated_at": body["_meta"]["updated_at"]},
    )


def main() -> int:
    # Default 8844 avoids collision with Vite (often 5173/5174) when both are used locally.
    port = int(env("PORT", "8844") or "8844")
    os.chdir(APP_ROOT)
    try:
        httpd = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    except OSError as e:
        if e.errno == errno.EADDRINUSE or "address already in use" in str(e).lower():
            print(
                f"\nPort {port} is already in use. Another app (e.g. Vite) may be using it.\n"
                f"Set a free PORT in .env (e.g. PORT=8844) or stop the other program, then run again.\n",
                flush=True,
            )
        raise
    print(f"John’sStockApp → http://localhost:{port}", flush=True)
    ips = collect_lan_ipv4s()
    mob_path = write_mobile_url_file(port, ips)
    if ips:
        print("", flush=True)
        print("iPhone / iPad (same Wi-Fi) — copy ONE line into Safari:", flush=True)
        for ip in ips:
            print(f"   http://{ip}:{port}", flush=True)
        if mob_path:
            print("", flush=True)
            print(f"Same URLs saved to: {mob_path}", flush=True)
    else:
        print("", flush=True)
        print(
            f"No LAN IP auto-detected. In System Settings → Network → Wi-Fi → Details, copy IP, then Safari:",
            flush=True,
        )
        print(f"   http://<that-ip>:{port}", flush=True)
        if mob_path:
            print(f"(Template saved to: {mob_path})", flush=True)
    print(
        f"API: /api/search  /api/quote  /api/history  /api/news  /api/corporate  /api/fx-eur  /api/health  GET /api/ai-commentary  GET /api/ai-ask  |  Yahoo={_use_yahoo()}  |  FMP={'on' if env('FMP_API_KEY') else 'off'}",
        flush=True,
    )
    print(
        "If the browser shows 501 on POST, an old server may still be bound to this port — run: "
        f"lsof -iTCP:{port} -sTCP:LISTEN  then kill that PID, and start this server again. GET /api/ai-commentary avoids POST.",
        flush=True,
    )
    print("Health check: GET /api/health  →  expect api_revision: 16, llm_commentary, shared_family_portfolio.", flush=True)
    t212_instruments_warmer_start(from_boot=True)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
