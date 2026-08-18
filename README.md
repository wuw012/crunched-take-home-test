# Spreadsheet agent in Excel

Crunched take-home. Chat in Excel's task pane reads and writes the open workbook, including formulas. It never dumps a whole sheet into the model.

Repo: https://github.com/wuw012/crunched-take-home-test (public).

## Setup

Node 20+, Python 3.11+, desktop Excel for Mac, `ANTHROPIC_API_KEY`. Both processes have to be running. Default model is Sonnet 4.5 (`ANTHROPIC_MODEL` overrides). Restart uvicorn after changing `.env`.

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

`pytest` from `backend/` after the install. Accept the certificate prompt if Excel shows one. In Excel: Home, Show agent. Open `samples/demo.xlsx` (rebuild with `python backend/scripts/make_demo.py`). `npm start` sideloads `manifest.xml`; `npm stop` unloads it. The home screen lists the walkthrough prompts.

## General thoughts

Excel add-ins run in a sandboxed WebView. Python cannot open the file. The model decides; Excel executes.

```
User -> Chat.tsx -> loop.ts -> POST /api/chat -> LangGraph step() -> Claude
                     ^                                              |
                     +---- Office.js tools <------------------------+
```

One-node graph (`START -> reason -> END`), tool schemas only, no `ToolNode`. FastAPI returns `message` or `tool_calls`. The loop is `src/agent/loop.ts`. Office.js is only in `gateway.ts`. A new tool is a Pydantic schema plus one gateway function.

I kept a stateless `POST /api/chat` so each step is a DTO in the network tab. Putting Claude in the pane would drop FastAPI, but the brief asked for Python and LangGraph, and the API key would sit in the WebView. A server-owned loop over SSE is closer to a real agent host; LangGraph `interrupt` is the idiomatic graph when tools cannot run in Python. Same ports, more to explain in 15 minutes. Two hand-written schemas (`tools.py` vs `executeTool`) will drift. I would generate one contract next.

"Any size" means look, don't swallow. `list_workbook_structure` returns sheet names, used-range sizes, and chart metadata, no cell values. Reads check `rowCount` / `columnCount` before `values` and refuse over 2000 cells. Demo `Exports` is 81×26 (>2000 cells) so a full used-range read is refused; those dummy rows stay in Excel.

Reads return `values` and `formulas`, so hardcoded totals show up. Writes go cell by cell: strings that start with `=` are formulas, the rest are values, omitted cells are left alone. `start_cell` is one cell (B4), not a range. After GP is `=B2-B3`, change COGS yourself and GP moves. Excel is the calculator.

Office will not load the add-in over HTTP. An HTTPS page also cannot `fetch` HTTP. The pane is `https://localhost:3000` and calls same-origin `/api/chat`; webpack proxies that to `http://127.0.0.1:8000`. Dual mkcert on FastAPI still needs CORS (different port). Production is a real cert on a reverse proxy, `/` to the pane and `/api` to uvicorn.

## Known holes

Stop pairs cancelled `tool_result`s, so the next send is a legal Anthropic transcript. It does not undo Excel writes that already ran. `Excel.run` ignores `AbortSignal`.

Invalid `write_range` args do not 400 FastAPI. Python swallows `ValidationError` so the pane can refuse (`range_too_large`, `start_cell_not_a1`) and the model sees a tool result. The schemas are a contract for Claude, not a gate.

Write tests mock `Excel.run`. They check that labels go to `values`, formulas to `formulas`, and omitted cells are never assigned. There is no live-workbook test.

`create_chart` always pins to D8:L22. A second chart stacks on the same rectangle. `set_chart_type` without `name` is the last chart on the sheet. Skip the chart beat unless they ask.

Cell JSON in tool results is untrusted workbook data. The prompt says so. There is no delimiter around it.

Per-cell writes are for a P&L stub. A 2000-cell `write_range` would create 2000 Range objects in one `Excel.run`. I would batch rectangles before using this as a dump writer.

This is a laptop demo: sideload, `office-addin-dev-certs`, webpack `/api` proxy, key in `.env`. `urlProd` is still Contoso. `POST /api/chat` has no auth. Bind `0.0.0.0` and anyone who can reach the port spends the key. Webpack's `Access-Control-Allow-Origin: *` on that same origin can hit `/api/chat` while `npm start` is running. The action list is React state; refresh clears it. I only ran this on Excel for Mac.

If this shipped, I would start with hosting, auth, and one generated tool contract. I would not start with PDF ingest or web search.

## Limits

- 2000 cells per read/write (and chart source); over that, refuse and ask for a smaller A1 slice
- 16 tool rounds per turn; after 5 cell-inspect rounds (`get_selection` / `read_range`), further reads are blocked. Structure listing does not count.
- Last 48 messages trimmed by complete user turn; tool results truncated at 8k characters (shown as a failed action)
- In-memory chat only
- Current Excel selection is appended to the model user message (address and size only)

## 15-minute walkthrough

`samples/demo.xlsx`: broken P&L, Assumptions (Price=12, Units=100, 5% growth), Exports (81×26 dummy cells, over the read cap).

Mapped to a counterparty error-check like [Flanders](https://www.usecrunched.com/case-study/flander-invest), then a live driver like [Mile Marker](https://www.usecrunched.com/case-study/mile-marker). Not [DW Real Estate](https://www.usecrunched.com/case-study/dw-real-estate) CapEx/PDF or PE research. Those tools are not in the repo.

1. Orient. "What sheets are in this workbook and how large is each used range? Do not read all the data." Expect `list_workbook_structure` only. Exports used range should be over the cap. Do not expect a `read_range` of Exports.

2. Error-check. Select P&L `A1:B6`. "Gross Profit does not foot. Find the error and fix Gross Profit and Operating Profit the way a modeler would." GP is 500, should be 600. Expect `=B2-B3` and `=B4-B5`, not `600`. Change COGS yourself. GP should move.

3. Link drivers. "Revenue is hardcoded. Drive FY24 Revenue from Assumptions: Price x Units." Expect `=Assumptions!B2*Assumptions!B3` (every cell qualified). FY24 Revenue becomes 1200 (`12×100`). Change Price 12 to 15; Revenue becomes 1500.

Skip 4–5 unless they ask. Chart pins to D8:L22. Forecast formulas are `D2=B2*(1+Assumptions!$B$4)`, then `E2=D2*(1+…)`, `F2=E2*(1+…)` — not fill-right from B2 (that hits empty C).

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
