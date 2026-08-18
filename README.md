# Spreadsheet agent in Excel

Crunched take-home. Chat in Excel's task pane reads and writes the open workbook, including formulas. It never dumps a whole sheet into the model.

Repo: https://github.com/wuw012/crunched-take-home-test (public).

## Setup

Node 20+, Python 3.11+, desktop Excel for Mac, `ANTHROPIC_API_KEY`. Both processes have to be running.

```bash
cd backend
python3 -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
cp .env.example .env   # paste ANTHROPIC_API_KEY
uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

```bash
cd crunched-take-home
npm install && npm start
```

`pytest` from `backend/` after the install. Accept the certificate prompt if Excel shows one. Home, Show agent. Open `samples/demo.xlsx` (rebuild with `python backend/scripts/make_demo.py`). `npm start` sideloads `manifest.xml`; `npm stop` unloads it. The home screen lists the walkthrough prompts.

## General thoughts

Excel add-ins run in a sandboxed WebView. Python cannot open the file. The model decides; Excel executes.

```
User -> Chat.tsx -> loop.ts -> POST /api/chat -> LangGraph step() -> Claude
                     ^                                              |
                     +---- Office.js tools <------------------------+
```

One-node graph (`START -> reason -> END`), tool schemas only, no `ToolNode`. FastAPI returns `message` or `tool_calls`. The loop is `src/agent/loop.ts`. Office.js is only in `gateway.ts`. A new tool is a Pydantic schema plus one gateway function.

I kept a stateless `POST /api/chat` so each step is a DTO in the network tab. Putting Claude in the pane would drop FastAPI, but the brief asked for Python and LangGraph, and the API key would sit in the WebView. A server-owned loop over SSE is closer to a real agent host; LangGraph `interrupt` is the idiomatic graph when tools cannot run in Python. Same ports, more to explain in 15 minutes. Two hand-written schemas (`tools.py` vs `executeTool`) will drift. I would generate one contract next.

"Any size" means look, don't swallow. `list_workbook_structure` returns sheet names, used-range sizes, and chart metadata, no cell values. Reads check `rowCount` / `columnCount` before `values` and refuse over 2000 cells. `Exports` is in the demo so you can see those ~80 dummy rows never leave Excel.

Reads return `values` and `formulas`, so hardcoded totals show up. Writes starting with `=` go through `range.formulas`. After GP is `=B2-B3`, change COGS yourself and GP moves. Excel is the calculator.

Office will not load the add-in over HTTP. An HTTPS page also cannot `fetch` HTTP. The pane is `https://localhost:3000` and calls same-origin `/api/chat`; webpack proxies that to `http://127.0.0.1:8000`. Dual mkcert on FastAPI still needs CORS (different port). Production is a real cert on a reverse proxy, `/` to the pane and `/api` to uvicorn.

I skipped streaming, vision (`Range.getImage()` is flaky on Mac; charts change `chartType`), Smart/Fast, a stored audit trail, web search, attach, and IM/PDF. Adding a tool is cheap. I wanted a walkthrough I could finish.

This is a laptop demo: sideload, `office-addin-dev-certs`, webpack proxy, `ANTHROPIC_API_KEY` in `.env`. `urlProd` is still `https://www.contoso.com/`. `POST /api/chat` has no auth; bind `0.0.0.0` and anyone who can reach the port spends the key. The pane's action list is React state, not an audit trail. Refresh clears it. I only ran this on Excel for Mac.

I would not start production by adding PDF ingest or web search. Hosting, auth, and a contract that cannot drift come first.

## Limits

- 2000 cells per read/write; over that, refuse and ask for a smaller A1 slice
- 16 tool rounds per turn; after 3 inspect-only rounds, further reads are blocked
- Last 12 messages trimmed by complete user turn; tool results truncated at 8k characters
- In-memory chat only

## 15-minute walkthrough

`samples/demo.xlsx`: broken P&L, Assumptions (Price=12, Units=100, 5% growth), Exports (~80 dummy rows).

Mapped to a counterparty error-check like [Flanders](https://www.usecrunched.com/case-study/flander-invest), then a live driver like [Mile Marker](https://www.usecrunched.com/case-study/mile-marker). Not [DW Real Estate](https://www.usecrunched.com/case-study/dw-real-estate) CapEx/PDF or PE research. Those tools are not in the repo.

1. Orient. "What sheets are in this workbook and how large is each used range? Do not read all the data." Expect `list_workbook_structure` only.

2. Error-check. Select P&L `A1:B6`. "Gross Profit does not foot. Find the error and fix Gross Profit and Operating Profit the way a modeler would." GP is 500, should be 600. Expect `=B2-B3` and `=B4-B5`, not `600`. Change COGS yourself. GP should move.

3. Link drivers. "Revenue is hardcoded. Drive FY24 Revenue from Assumptions: Price x Units." Expect `=Assumptions!B2*Assumptions!B3`. Change Price 12 to 15. Revenue becomes 1500.

4-5. Optional. Chart the P&L, then switch it to a bar. Stub FY25-FY27 in D-F with `=B2*(1+Assumptions!$B$4)` and fill right.

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
