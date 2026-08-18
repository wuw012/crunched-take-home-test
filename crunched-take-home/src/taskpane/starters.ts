export type Starter = { label: string; hint: string; prompt: string };

export const STARTERS: Starter[] = [
  {
    label: "Orient the workbook",
    hint: "List sheets and sizes — do not ingest Exports",
    prompt: "What sheets are in this workbook and how large is each used range? Do not read all the data.",
  },
  {
    label: "Error-check the P&L",
    hint: "Gross Profit does not foot — write formulas",
    prompt:
      "Gross Profit does not foot. Find the error and fix Gross Profit and Operating Profit the way a modeler would.",
  },
  {
    label: "Link revenue to drivers",
    hint: "FY24 Revenue = Assumptions Price × Units",
    prompt: "Revenue is hardcoded. Drive FY24 Revenue from Assumptions: Price × Units.",
  },
  {
    label: "Chart the P&L",
    hint: "Clustered column, then switch it to a bar chart",
    prompt: "Chart this P&L as a clustered column. Then make it a bar chart.",
  },
  {
    label: "Stub a 3-year forecast",
    hint: "FY25–FY27 in D–F, YoY from Assumptions, formulas only",
    prompt:
      "Add a 3-year forecast in D–F: FY25–FY27 revenue growing at the Assumptions YoY growth rate. Formulas only, linked to FY24 Revenue.",
  },
];
