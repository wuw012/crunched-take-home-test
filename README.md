# Crunched take-home: spreadsheet agent in Excel

A simplified Crunched: chat in Excel's task pane, an AI agent that **reads and writes the live workbook**, including formulas.

Time-boxed. Prefer a small, explainable architecture over features.

## What it does

- Task pane chat inside desktop Excel
- Claude (Anthropic) decides whether to reply or call a tool
- Office.js executes tools against the open file
- Workbooks of any size: inspect structure first, refuse ranges over 2000 cells, then read a slice

Demo: open `samples/demo.xlsx`, select `A1:B6` on **P&L**, ask:

> Gross Profit and Operating Profit are hardcoded. Replace them with formulas.

You should see `=B2-B3` and `=B4-B5` in the formula bar.

## Architecture

Excel add-ins run in a sandboxed WebView. Python cannot open the workbook. Dumping the whole sheet into the model also fails the “any size” requirement.

So the agent is a **reasoner**, not a spreadsheet engine:

1. The task pane sends the chat transcript to FastAPI
2. A one-node LangGraph graph calls Claude with tool schemas
3. If Claude returns `tool_calls`, the **client** runs them via Office.js
4. Tool results go back until Claude answers in text

There is no LangGraph `ToolNode`. Excel does not exist on the server. FastAPI returns a small DTO (`message` or `tool_calls`). The loop lives in `crunched-take-home/src/agent/loop.ts`. The only Office.js file is `src/excel/gateway.ts`.

```
User → Chat.tsx → loop.ts → POST /api/chat → LangGraph step() → Claude
                     ↑                                              |
                     └──── Office.js tools ←────────────────────────┘
```

### Why not dual mkcert?

The brief notes that Excel's WebView will not load **add-in** URLs over HTTP (manifest, task pane, icons). The API is not one of those URLs. The task pane is served from `https://localhost:3000` and calls `/api/...` on the same origin. Webpack proxies `/api` to `http://127.0.0.1:8000`. One trusted certificate (Yeoman / `office-addin-dev-certs`). Mixed content never happens.

If the proxy is removed, FastAPI would need HTTPS (mkcert) because an HTTPS pane cannot `fetch` HTTP.

### What we did not build

Streaming tokens, screenshot/vision (`Range.getImage()` is flaky on Mac — restyle charts with `set_chart_type`, not pixels), Smart/Fast model routing, audit log, web search, attach files. Adding a tool is a schema + one Office.js function.

## Setup

Needs: Node 20+, Python 3.11+, desktop Excel for Mac, `ANTHROPIC_API_KEY`.

```bash
# backend
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
cp .env.example .env   # or export ANTHROPIC_API_KEY
uvicorn app.main:app --reload --port 8000
```

```bash
# add-in (second terminal)
cd crunched-take-home
npm install
npm start
```

Accept the certificate prompt if asked. In Excel: **Home → Show agent**.

Open `samples/demo.xlsx` (regenerate with `python backend/scripts/make_demo.py`). The task pane home screen lists the five demo prompts in walkthrough order.

`npm start` sideloads `manifest.xml`. Stop with `npm stop`.

## 15-minute demo (case-study jobs, not a chatbot)

Workbook: `samples/demo.xlsx` — **P&L** (broken, like a file from a counterparty), **Assumptions** (Price=12, Units=100, 5% growth), **Exports** (~80 dummy rows).

Tell the story against two published Crunched cases. Do not claim the others (PDF/IM extraction, web research, tenant CapEx waterfalls) — we did not build those.

### 1. Orient — any-size workbooks

> What sheets are in this workbook and how large is each used range? Do not read all the data.

**Expect:** `list_workbook_structure` only. P&L is tiny; Exports is ~80 rows.

**Say:** Crunched's product claim is workbooks of any size. The agent *addresses* the file; it does not ingest Exports. The 2000-cell guard in `gateway.ts` is that policy as code.

### 2. Error-check — Flanders Investment Company

[FICO](https://www.usecrunched.com/case-study/flander-invest): 5× faster error detection on models received from selling parties. Pieter even teaches this as “error checking on a spreadsheet with intentional mistakes.”

Select `A1:B6` on **P&L**.

> Gross Profit does not foot. Find the error and fix Gross Profit and Operating Profit the way a modeler would.

**Expect:** GP is 500, should be 600. Agent writes `=B2-B3` and `=B4-B5`, not `600` and `450`.

**Say:** Same job as FICO, one screen. The LLM is not the calculator. You are the reviewer: change COGS yourself — GP should move.

### 3. Link drivers — Mile Marker Advisors

[Mile Marker](https://www.usecrunched.com/case-study/mile-marker): formula-driven models that link assumptions across sheets; the associate stays the expert reviewer.

> Revenue is hardcoded. Drive FY24 Revenue from Assumptions: Price × Units.

**Expect:** `=Assumptions!B2*Assumptions!B3` in B2.

**Say:** Then *you* change Price from 12 to 15. Revenue becomes 1500 with no second prompt. That is their “augment, don't replace” line.

### 4. Chart (optional) — live Excel object

> Chart this P&L as a clustered column. Then make it a bar chart.

**Say:** Same architecture as formulas: change `chartType`, do not screenshot.

If time is short, drop 1 and 4. **Never drop 2** — that is the Flanders job and the Crunched hire signal.

Optional extra: 3-year forecast in D–F (`=B2*(1+Assumptions!$B$4)`, …). That is a tiny Mile Marker “scaffold the model” beat, not a new product.

### What we are not demoing (and why)

- **DW Real Estate** (tenant CapEx, waterfall, PDF tables) — needs document ingest we cut
- **European PE screening / IC packs** — needs research agents we cut
- **Capital IQ formulas** — third-party data we do not have

### What to open in the code (2 minutes)

- `src/excel/gateway.ts` — Excel I/O + refuse-before-load
- `src/agent/loop.ts` — client-side tool loop (Excel cannot run in Python)
- `backend/app/agent/graph.py` — one `reason` node, no `ToolNode`

## Layout

```
crunched-take-home/     Excel React task pane (HTTPS :3000)
  src/agent/loop.ts     tool loop + step cap
  src/excel/gateway.ts  Office.js only
  src/excel/tools.ts    name → gateway, unknown tool fails closed
backend/                 FastAPI :8000
  app/agent/graph.py    LangGraph START → reason → END
  app/agent/tools.py    schemas only
samples/demo.xlsx
```

## Known limits

- 2000-cell cap; last 12 messages trimmed by complete user turn (never mid tool-call); tool results truncated at 8k characters; max 16 tool rounds per turn; after 3 inspect-only rounds further reads are refused
- In-memory chat only (refreshing the pane clears it)
- Reads return both `values` and `formulas`. Formula writes: cells whose string starts with `=` go through `range.formulas`
