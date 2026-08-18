# Spreadsheet agent in Excel

Take-home for Crunched: a task-pane chat that **reads and writes the live workbook**, including formulas, without dumping a sheet of any size into the model.

Repo: https://github.com/wuw012/crunched-take-home-test (public).

## Setup

Needs Node 20+, Python 3.11+, desktop Excel for Mac, and `ANTHROPIC_API_KEY`.

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"
cp .env.example .env   # paste ANTHROPIC_API_KEY
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

```bash
cd crunched-take-home
npm install
npm start
```

Run **both** processes. Optional: `pytest` from `backend/` after `pip install -e ".[dev]"`.

Accept the certificate prompt if asked. In Excel: **Home → Show agent**. Open `samples/demo.xlsx` (regenerate with `python backend/scripts/make_demo.py`). Sideload is `manifest.xml`; stop with `npm stop`.

The home screen lists the walkthrough prompts in order.

## General thoughts

Excel add-ins run in a sandboxed WebView. Python cannot open the file. So the agent is a **reasoner**, not a spreadsheet engine:

1. The pane sends the transcript to FastAPI.
2. A one-node LangGraph graph (`START → reason → END`) calls Claude with tool schemas only.
3. If Claude returns `tool_calls`, the **client** runs them in Office.js.
4. Results go back until Claude answers in text.

There is no LangGraph `ToolNode`. Excel does not exist on the server. FastAPI returns a small DTO (`message` or `tool_calls`). The loop is `crunched-take-home/src/agent/loop.ts`.

```
User → Chat.tsx → loop.ts → POST /api/chat → LangGraph step() → Claude
                     ↑                                              |
                     └──── Office.js tools ←────────────────────────┘
```

The shape is three ports: a **reasoner**, a **tool contract**, and an **Excel adapter**. Office.js is only in `gateway.ts`. Adding a tool is a Pydantic schema plus one gateway function.

**Same split, other drawings**

COM, Graph, or a formula engine would only replace the adapter. They are not a different architecture. The choices at this layer were:

- **Reasoner in the add-in** (Anthropic from TypeScript, loop and Office.js in one WebView). One process, no FastAPI. The API key would live in the pane, and the brief asked for Python + LangGraph. Crunched also needs a server if they ever route models or log turns.
- **Server-owned loop** (one request holds the conversation; the pane is a dumb executor over SSE/WebSocket). Closer to a production agent host. I kept **stateless `POST /api/chat`** so each step is a DTO you can show in the network tab: `message` or `tool_calls`.
- **LangGraph `interrupt` as the loop**, client resumes with tool results. That is the idiomatic graph when tools cannot run in Python. A one-node `reason` + TypeScript loop is the same ports with less graph machinery to defend in 15 minutes.
- **One generated tool contract** (TS types → Pydantic, or the reverse). Cleaner than two hand-written schemas. I did not add a codegen step for a 4-hour take-home; the drift risk is real and I would do that next.

**Any size** means address the workbook, do not ingest it. `list_workbook_structure` returns sheet names, used-range dimensions, and chart metadata — never cell values. Reads load `rowCount` / `columnCount` before `values`. Over 2000 cells, the tool refuses and asks for a smaller A1 slice. `Exports` in the demo is there to prove we never send ~80 dummy rows.

**Formulas, not pasted numbers.** Reads return both computed `values` and `formulas`, so hardcoded totals are visible. Writes whose strings start with `=` go through `range.formulas`. After Gross Profit is `=B2-B3`, changing COGS updates GP without another prompt. That is the product: Excel stays the engine; the associate stays the reviewer.

**HTTPS.** Office will not load the add-in itself over HTTP (manifest, task pane, icons). FastAPI is not one of those URLs. The WebView is a browser: an HTTPS page cannot `fetch` HTTP (mixed content). This repo uses the usual production shape, locally. The pane is `https://localhost:3000` and calls same-origin `/api/chat`; webpack proxies that to `http://127.0.0.1:8000`. The WebView never sees HTTP. `office-addin-dev-certs` is a local CA. In production you put a real certificate on a reverse proxy in front of the static add-in and FastAPI the same way (`/` → task pane, `/api` → uvicorn).

If the proxy were removed and the pane called `http://127.0.0.1:8000` directly, Excel would block it. TLS on FastAPI at `https://localhost:8000` would still be a different origin (different port), so you would also need CORS. Same-origin proxy is the smaller local setup. Dual mkcert is valid; it is not required.

**What I cut.** Token streaming, screenshot/vision (`Range.getImage()` is unreliable on Mac — charts change `chartType` instead), Smart/Fast routing, a persistent audit trail (the pane only shows workbook actions for the current session), web search, file attach, IM/PDF extraction. Each extra tool is a schema plus one Office.js function; the walkthrough is easier to explain if that list stays short.

## This is not production

Sideload, `office-addin-dev-certs`, webpack’s `/api` proxy, `uvicorn --reload --host 127.0.0.1`, and `ANTHROPIC_API_KEY` in `backend/.env` are a laptop demo. A firm would not run that.

What is already true in this repo, and what I would change without moving the ports:

- **The add-in is not hosted.** `webpack.config.js` still has `urlProd = "https://www.contoso.com/"`. Production is: HTTPS static files at a real origin, that origin in `manifest.xml`, TLS on a reverse proxy with `/` → the pane and `/api` → FastAPI. `office-addin-dev-certs` does not ship.
- **The API is a private key with no lock.** `POST /api/chat` has no auth. CORS is allowlisted to `https://localhost:3000` in `main.py`; the webpack proxy means the pane does not even need CORS today. Bind `0.0.0.0` and anyone who can reach the port spends the Anthropic key. Production: the pane gets a user session; the server holds the model key; the API rejects unauthenticated steps.
- **The loop is a debugger.** Stateless `POST /api/chat` is easy to show in the network tab. It also means N model round-trips per user turn, no server-side cancellation if the pane dies, and two hand-written tool schemas (`app/agent/tools.py` vs `executeTool` in TypeScript) that will drift. Production: one generated contract; then, if the extra machinery pays, a server-owned loop with retries and a kill switch.
- **“Any size” is a refuse, not a scanner.** 2000 cells is a hard stop (`policy.ts` / `MAX_CELLS`). That is the right *policy* (do not ingest `Exports`). It is not a reader for a million-row sheet. Production still starts with structure, then pages slices with a progress line — same gateway, not a dump.
- **There is no audit trail.** The pane lists workbook actions from the current React state. Refresh clears them. Nobody is tied to a range. Crunched’s trail is: this user wrote `=B2-B3` at P&L!B4 at this time. That is storage plus identity, not a nicer chat bubble.
- **I only ran this on Excel for Mac.** Office.js is the port; Windows and Excel for the web still have to be checked. I would not claim them.

I would not start production by adding Smart/Fast, PDF ingest, or web search. Those are more tools. The first production work is hosting, auth, and a contract that cannot drift.

## Limits

- 2000 cells per read/write; over that the tool refuses and asks for a smaller A1 slice
- 16 tool rounds per turn; after 3 inspect-only rounds further reads are blocked
- Last 12 messages trimmed by complete user turn; tool results truncated at 8k characters
- In-memory chat only (refreshing the pane clears it)

## 15-minute walkthrough

Workbook: `samples/demo.xlsx` — broken **P&L**, **Assumptions** (Price=12, Units=100, 5% growth), **Exports** (~80 rows).

This is a counterparty error-check in the spirit of [Flanders Investment Company](https://www.usecrunched.com/case-study/flander-invest), then a live driver link in the spirit of [Mile Marker Advisors](https://www.usecrunched.com/case-study/mile-marker). I am not demoing [DW Real Estate](https://www.usecrunched.com/case-study/dw-real-estate) CapEx/PDF ingest or PE research agents; those tools are not in the repo.

**1. Orient.** *What sheets are in this workbook and how large is each used range? Do not read all the data.*  
Expect `list_workbook_structure` only.

**2. Error-check (do not skip).** Select P&L `A1:B6`. *Gross Profit does not foot. Find the error and fix Gross Profit and Operating Profit the way a modeler would.*  
GP is 500, should be 600. Expect `=B2-B3` and `=B4-B5`, not `600`. Then change COGS yourself — GP should move.

**3. Link drivers.** *Revenue is hardcoded. Drive FY24 Revenue from Assumptions: Price × Units.*  
Expect `=Assumptions!B2*Assumptions!B3`. Change Price 12→15; Revenue becomes 1500 with no second prompt.

**4–5. Optional.** Chart P&L then switch to bar (`create_chart` / `set_chart_type`). Stub FY25–FY27 in D–F with `=B2*(1+Assumptions!$B$4)` and fill-right.

If time is short: 2, then 3, then `gateway.ts` / `loop.ts` / `graph.py`.

## Layout

```
crunched-take-home/     Excel React task pane (HTTPS :3000)
  src/agent/loop.ts     client tool loop
  src/excel/gateway.ts  Office.js + refuse-before-load
backend/                 FastAPI :8000
  app/agent/graph.py    one reason node, no ToolNode
  app/agent/tools.py    schemas only
samples/demo.xlsx
```
