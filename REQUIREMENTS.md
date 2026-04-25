# John’sStockApp — personal requirements

**Audience:** Single user — personal study, analysis, and reference only. **Not** for public distribution.

**Goal:** One web app that gives a single view of investments and supports basic research from within the app. The design must allow **future features and changes** without a full rewrite.

---

## Background (user context)

- Investor in the **Indian** share market since **2008**.
- Since **2023**, also investing in **European**, **British**, and **American** stocks and ETFs via **Trading 212** (resident and tax resident in **Germany**; German citizen; originally from India with **OCI**).

---

## 1. Search

- Web-based **search** for **Indian**, **European**, **British**, and **American** stocks and ETFs listed on their respective exchanges.
- Search results must indicate the **local currency** of the respective exchange (where no currency symbol exists, use the **short form** of the currency, e.g. EUR, INR, USD).
- User must be able to **filter** results by **country**, **exchange**, and **currency**, then select an instrument.

---

## 2. Selected stock / ETF — primary view (overview)

After selecting a stock or ETF from filtered results, show **first**:

| Data | Priority |
|------|----------|
| **Current price** (near the symbol / name) | High |
| **Previous closing price** | High |
| **Today’s opening price** | High |
| **Today’s high and low** | High |
| **52-week high and low** | High |

**Immediately below the quote block:**

- **News** pertaining to the selected stock/ETF.
- **Corporate action** calendar (or equivalent) for that instrument.

Information should **not** all appear in one overwhelming screen: use **progressive disclosure** — summary first, deeper detail behind tabs or clear navigation.

---

## 3. Selected stock / ETF — secondary areas (tabs)

### Tab A — Technical analysis & chart

- **Technical analysis** with commonly used parameters (RSI, moving averages, trends, Bollinger Bands, etc.).
- **RS55** (as presented by **Vivek Bajaj**) — marked as **very important** to the user; exact definition/rules to be confirmed at implementation time.
- **Buy / Hold / Sell** indication based on **AI-assisted analysis** (must be framed as **educational / informational**, not financial advice).
- **Market sentiment** label, e.g. **Bullish**, **Neutral**, **Bearish**.
- **Chart** — the **main OHLC/area** block is in a **collapsible “Price history”** section (default closed) so the overview is not dominated by the canvas; the user can expand for chart tools or load indicators from the **Technicals** tab.
- **Chart types** with at least **three** options:
  - **Candle**
  - **Heikin Ashi**
  - **Area**
- Chart should support common **studies/indicators**, including at least: **RSI**, **moving averages**, **trend markers** (buy/sell style signals), **Bollinger Bands**, **RS55**.

### Tab B — Fundamental analysis

- **Fundamental analysis** view for the same stock or ETF (depth and data sources to be defined when implementing).
- **AI narrative** must use **official provider APIs** (keys in server `.env`), not scraping of Grok / Gemini / Copilot / ChatGPT web chat.

---

## 4. Portfolio (main tab)

- Import portfolio as **CSV** (Trading 212 history export, Zerodha holdings saved as CSV, or the app template). **PDF** import was considered for T212 / Zerodha statements; it remains **out of scope for now** — CSV covers the practical workflow.

**Two separate portfolios under this section:**

1. **Trading 212** — one tracked portfolio.  
2. **Zerodha** — second tracked portfolio.

For each holding (under both portfolios), show:

- **Full name** and **symbol**
- **Exchange**
- **Average buy price**
- **Current price**
- **Profit** (green) / **Loss** (red) in **local currency** and **percentage**

**Analytics** (sector, country, and other breakdowns) using **pie charts** or other appropriate charts.

**Priority (owner 2026): Trading 212 portfolio intelligence**

- **Sectoral lens** with explicit weight buckets such as **Information technology**, **Manufacturing & mobility**, **Medical & pharmaceuticals**, **Consumer & gastronomy**, **Metals & materials**, **Energy**, **Financial services**, **Crypto & digital assets**, and **Other** — value-weighted, clearly labeled as **heuristic** until a paid GICS / data vendor is wired.
- **AI-driven narrative** on top of those weights (concentration, overlap, EUR context for Germany-based T212 use) — framed as **educational**, not advice; may use a server-side LLM with keys in `.env` only.

**Interaction:** Clicking any stock or ETF in a portfolio should open the **same** rich instrument view as from search (overview, news, corporate actions, technical tab, chart types, fundamental tab).

---

## 5. Notes (main tab)

- Dedicated area for **personal notes** (study, thesis, reminders).

---

## 6. Theme

- **White**, **Dark**, and **Auto** (follow system preference) theme options.
- Overall look: **neat and attractive** colour scheme; **not cluttered** on first load.

---

## 7. Currency & future “Euro total” view

- Display values using **local currency symbols** where available; otherwise **currency short codes** (EUR, INR, USD, GBP, etc.).

**Future (reserved):**

- A tab or section ready to show **growth of total portfolio holdings in Euro** over time.

**Implemented (multi-ledger, ECB reference):**

- **All ledgers in €** — a **total + per-ledger** block at the top of Portfolio using the same rows as each tab, converted with ECB reference rates (`/api/fx-eur` / Frankfurter). The combined total includes **T212, Crypto, Zerodha, eToro, MFs, insurance, fixed deposits** whenever those rows have a valid ISO currency the ECB set returns.
- **Non-ISO `EURO` / `€` in saved rows (e.g. T212 / manual)** must map to **`EUR`** for the FX table or those rows are skipped in € sums. The app **migrates** `EURO` / `€` to **`EUR` in local storage** on load, normalizes on import/sync, and uses a shared **`eurPerUnitToEur`** helper so **`EUR` always has rate 1** even if an API object omits `eur_per_unit.EUR`.
- **Crypto live quotes (Yahoo `*-USD`)** with a row in **EUR**: the app multiplies the USD spot by the **USD→€** ECB rate so `last` and the **€** roll-up match a Euro-denominated position. The same **USD→€** adjustment applies to **coin rows on the Trading 212 tab** when the row is detected as crypto (symbol / exchange / quote kind) and the row currency is **EUR**. **T212 tab** uses the crypto-style Yahoo pair for those symbols when requesting quotes.
- **By ledger (€) — split** — If coins still appear only under the **Trading 212** import (not under **Crypto (T212)**), the combined **€** view **attributes** their value to **Crypto (T212)** and **excludes** them from the **Trading 212** € leg so the **grand total** is not double-counted. The full **T212** table on the T212 tab is unchanged and still lists every T212 row.
- **Currency in tables** — the euro area uses the **€** label (and ISO `EUR` under the hood) instead of the literal string `EURO` where the app controls formatting.

**Done for single tabs:** combined **market value and P/L in EUR** for Trading 212 / crypto-style tables using the same ECB reference (not bank executable rates).

**Instrument page — price chart (2026-04):**

- **Price history** is a **collapsible** block (closed by default) to save main-column space. **History loads** when the user **expands** that block **or** opens the **Technicals** tab (indicators need the same series). The **Fundamentals** tab does not auto-fetch history; expand **Price history** when a chart is wanted there.

---

## 8. UX principles (non-functional)

- **Progressive disclosure:** do not present all information in a single window; point users to further detail inside the app.
- **Extensibility:** structure and APIs should make it straightforward to **add more features later** (the user has further plans).
- **Personal use only** — not a product for public distribution; still follow good practices for **API keys** (server-side, not exposed in the browser) and **disclaimers** on AI/sentiment outputs.

---

## 9. Implementation status (high level)

This section is a living checklist; update as the app evolves.

| Area | Status (as of project rebuild) |
|------|----------------------------------|
| Search + filters + quote basics | Partially implemented |
| Instrument page `#/symbol?…` (Overview / Technicals & chart / Fundamentals) | Implemented — Fundamentals tab shows quote snapshot + volume + optional ratio fields; LLM block explains API-only AI path; news + corporate wired |
| Add to portfolio from instrument page | Implemented |
| Watchlist tab + add from instrument | Implemented (local storage) |
| Portfolio row → same instrument page as Search | Implemented (symbol link) |
| News + corporate actions on symbol | Implemented (`/api/news` Yahoo RSS + optional FMP; `/api/corporate` EODHD and/or FMP) |
| Tabbed symbol workspace — full TA, AI labels | Not implemented (Technicals: chart types, overlays, RS55 proxy, SMA/RSI/BB text; no AI labels yet) |
| Multi-style charts + studies (Bollinger, MAs, RS55) | Partial — **Area / Candles / Heikin** + optional **SMA20**, **SMA50**, **Bollinger (20,2σ)** overlays; **RS55-style** vs **SPY** (non-India) or **^NSEI** (NSE/BSE); trend markers / Supertrend not built |
| AI Buy/Hold/Sell + sentiment | Not implemented |
| Dual portfolio (T212 vs Zerodha) — separate rows per broker | Implemented (CSV / manual / refresh per tab; migration from old single list → T212) |
| PDF import (T212 / Zerodha) | Deferred — CSV import used instead |
| Portfolio analytics (sector/country charts) | Partial — **four** compact **allocation** pies (no long briefing card); **LLM AI narrative** not wired yet |
| Notes | Implemented (local storage) |
| Themes (incl. Auto) | Implemented (Light / Dark / Auto + system listener) |
| Total portfolio in EUR / growth tab | Partial — **all ledgers** top card + per-ledger grid (ECB ref) when FX loads; `EURO`→`EUR` normalization; growth-over-time tab not built |

---

## 10. Layout & portfolio UX (2026-04)

- **Full-width main column** — the app shell uses the **full browser width** (no 720px cap on `<main>`); content uses **responsive padding** so tables and summaries use more space on desktop.
- **Less on-screen copy** — the Portfolio page **omits** long instructional paragraphs (import tips, local-storage essays, “portfolio briefing” narrative, yellow **AI keys** info banner, sector **details** block, and most disclaimers on EUR cards). Error banners when the API is **down** or the server is **stale** still appear.
- **Allocation pies** — four **smaller** donuts in **one row** on wide viewports; **two columns** on medium, **one** on small phones. The rule-based “briefing” card is **off** to reduce clutter.
- **Tables** — holdings tables use **tighter padding**, **slightly smaller type**, and **`table-layout: fixed`** to reduce **horizontal scroll**; wrap still possible on very narrow phones.
- **Sync from Trading 212 (T212 tab)** — after a successful sync, the app **fetches live quotes** for both the **T212** and **Crypto (T212)** ledgers so `last` prices and € roll-ups can populate without a separate refresh (user can still use **Refresh prices** for other cases).
- **Next (explicitly future)** — richer **colour / background** theme and optional **side navigation** vs top tabs; not part of this iteration.

---

*Document generated from the product owner’s written specification; amend in place as requirements change. The assistant updates this file when the product owner asks to align the written spec with shipped behaviour, or when recording agreed scope (not on every code edit by default).*

---
