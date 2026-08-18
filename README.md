# Spreadsheet agent in Excel

Crunched take-home. Chat in Excel's task pane reads and writes the open workbook, including formulas. It never dumps a whole sheet into the model.

Repo: https://github.com/wuw012/crunched-take-home-test (public).

## Setup

Node 20+, Python 3.11+, desktop Excel for Mac, `ANTHROPIC_API_KEY`. Both processes have to be running. Default model is Sonnet 4.5. Restart uvicorn after changing `.env`.

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

`pytest` from `backend/` after the install. Accept the certificate prompt if Excel shows one. Home, Show agent. Open `samples/demo.xlsx` (rebuild with `python backend/scripts/make_demo.py`). `npm start` sideloads `manifest.xml`; `npm stop` unloads it.

## General thoughts

Excel add-ins run in a sandboxed WebView. Python cannot open the file. The model decides; Excel executes.

```
User -> Chat.tsx -> loop.ts -> POST /api/chat -> LangGraph step() -> Claude
                     ^                                              |
                     +---- Office.js tools <------------------------+
```

One-node graph, tool schemas only, no `ToolNode`. FastAPI returns `message` or `tool_calls`. Loop is `loop.ts`. Office.js is only in `gateway.ts`. A new tool is a Pydantic schema plus one gateway function.

I mirrored your stack (React, Office.js, FastAPI, LangGraph). Office.js stays in the pane because Python cannot `Excel.run`. I used one `reason` node so the server never pretends to open the file. The loop is still linear; LangGraph is the wrapper. I kept stateless `POST /api/chat` so each step is a DTO in the network tab. Two hand-written schemas will drift. I would generate one contract next.

"Any size" means look, don't swallow. `list_workbook_structure` returns names and used-range sizes, no cell values. Reads check `rowCount` / `columnCount` before `values` and refuse over 2000 cells. Demo Exports is 81×26, so a used-range read refuses on camera.

Reads return `values` and `formulas`. Writes go cell by cell: `=` is a formula, the rest are values, omitted cells are left alone. `start_cell` is one cell (B4). After GP is `=B2-B3`, change COGS yourself. Excel is the calculator.

The pane is `https://localhost:3000` and calls same-origin `/api/chat`; webpack proxies to `http://127.0.0.1:8000`. Office will not load the add-in over HTTP, and an HTTPS page cannot `fetch` HTTP. Production is a real cert on a reverse proxy.

Stop pairs cancelled tool results so the next send does not 500; it does not undo Excel writes already in the grid. Python does not 400 invalid `write_range` args; the pane refuses and the model sees the tool error. Write tests mock `Excel.run`. Charts pin to D8:L22. Cell JSON is untrusted; no delimiter.

Laptop demo: no auth, Contoso `urlProd`, key in `.env`, Mac only, action list is React state. I would ship hosting, auth, and one contract before PDF ingest.

Caps: 2000 cells; 16 tool rounds; 5 cell-inspect rounds then block further reads (structure listing does not count); last 48 messages trimmed by complete user turn.

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

## 15-minute walkthrough

`samples/demo.xlsx`: broken P&L, Assumptions (Price=12, Units=100, 5% growth), Exports over the read cap.

Flanders-shaped error-check, then Mile Marker-shaped driver link. Not DW CapEx/PDF. Those tools are not in the repo.

1. Orient. "What sheets are in this workbook and how large is each used range? Do not read all the data." Expect `list_workbook_structure` only. Exports should be over the cap.

2. Error-check. Select P&L `A1:B6`. "Gross Profit does not foot. Find the error and fix Gross Profit and Operating Profit the way a modeler would." GP is 500, should be 600. Expect `=B2-B3` and `=B4-B5`, not `600`. Change COGS; GP should move.

3. Link drivers. "Revenue is hardcoded. Drive FY24 Revenue from Assumptions: Price x Units." Expect `=Assumptions!B2*Assumptions!B3`. Change Price 12 to 15; Revenue becomes 1500.

Skip chart/forecast unless they ask. If time is short: 2, then 3, then `gateway.ts` / `loop.ts` / `graph.py`.
