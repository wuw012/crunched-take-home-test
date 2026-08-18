import assert from "node:assert/strict";
import { test } from "node:test";
import { ChatMessage } from "../api/types";
import { MAX_INSPECT_ROUNDS } from "../shared/limits";
import { runTurn, userMessageContent } from "./loop";

test("abort after the first model step skips later steps and tools", async () => {
  const controller = new AbortController();
  let steps = 0;
  const executed: string[] = [];
  const statuses: string[] = [];

  const messages = await runTurn(
    [],
    "fix gross profit",
    (status) => {
      statuses.push(status);
    },
    () => undefined,
    {
      signal: controller.signal,
      stepChat: async () => {
        steps += 1;
        if (steps > 1) {
          throw new Error("stepChat should not run after abort");
        }
        controller.abort();
        return {
          type: "tool_calls",
          tool_calls: [{ id: "1", name: "read_range", args: { sheet: "P&L", a1: "A1:B6" } }],
        };
      },
      executeTool: async (name) => {
        executed.push(name);
        return "{}";
      },
    }
  );

  assert.equal(steps, 1);
  assert.deepEqual(executed, []);
  assert.ok(statuses.includes("Stopping…"));
  assert.ok(messages.some((message: ChatMessage) => message.role === "user"));
  assert.equal(
    messages.some((message) => message.role === "tool"),
    false
  );
});

test("abort after tool_calls are published cancels remaining tools", async () => {
  const controller = new AbortController();
  const executed: string[] = [];

  const messages = await runTurn(
    [],
    "fix gross profit",
    () => undefined,
    (next) => {
      const last = next[next.length - 1];
      if (last?.role === "assistant" && last.tool_calls?.length) {
        controller.abort();
      }
    },
    {
      signal: controller.signal,
      stepChat: async () => ({
        type: "tool_calls",
        tool_calls: [
          { id: "1", name: "read_range", args: { sheet: "P&L", a1: "A1:B6" } },
          { id: "2", name: "write_range", args: { sheet: "P&L", start_cell: "B4" } },
        ],
      }),
      executeTool: async (name) => {
        executed.push(name);
        return "{}";
      },
    }
  );

  assert.deepEqual(executed, []);
  const tools = messages.filter((message) => message.role === "tool");
  assert.equal(tools.length, 2);
  assert.equal(JSON.parse(tools[0].content).error, "cancelled");
  assert.equal(JSON.parse(tools[1].content).error, "cancelled");
});

test("abort after the first tool still pairs remaining tool_calls", async () => {
  const controller = new AbortController();
  const executed: string[] = [];

  const messages = await runTurn(
    [],
    "fix gross profit",
    () => undefined,
    () => undefined,
    {
      signal: controller.signal,
      stepChat: async () => ({
        type: "tool_calls",
        tool_calls: [
          { id: "1", name: "read_range", args: { sheet: "P&L", a1: "A1:B6" } },
          { id: "2", name: "write_range", args: { sheet: "P&L", start_cell: "B4" } },
        ],
      }),
      executeTool: async (name) => {
        executed.push(name);
        controller.abort();
        return "{}";
      },
    }
  );

  assert.deepEqual(executed, ["read_range"]);
  const tools = messages.filter((message) => message.role === "tool");
  assert.equal(tools.length, 2);
  assert.equal(JSON.parse(tools[0].content).error, undefined);
  assert.equal(JSON.parse(tools[1].content).error, "cancelled");
});

test("blocked inspect retries return too_many_reads instead of stopping the turn", async () => {
  const ranges = ["A1:B3", "A1:B6", "A1:B4", "A1:B5", "A1:B2", "A1:B2"];
  let steps = 0;

  const messages = await runTurn(
    [],
    "fix gross profit",
    () => undefined,
    () => undefined,
    {
      stepChat: async (history) => {
        const lastTool = [...history].reverse().find((message) => message.role === "tool");
        if (lastTool && JSON.parse(lastTool.content).error === "too_many_reads") {
          return { type: "message", content: "I will write formulas now." };
        }
        const a1 = ranges[Math.min(steps, ranges.length - 1)];
        steps += 1;
        return {
          type: "tool_calls",
          tool_calls: [{ id: String(steps), name: "read_range", args: { sheet: "P&L", a1 } }],
        };
      },
      executeTool: async () => JSON.stringify({ values: [[1]] }),
    }
  );

  const toolMessages = messages.filter((message) => message.role === "tool");
  assert.equal(toolMessages.length, MAX_INSPECT_ROUNDS + 1);
  assert.equal(JSON.parse(toolMessages[MAX_INSPECT_ROUNDS].content).error, "too_many_reads");
  const last = messages[messages.length - 1];
  assert.equal(last?.role, "assistant");
  assert.equal(last && last.role === "assistant" ? last.content : null, "I will write formulas now.");
});

test("userMessageContent appends the Excel selection", () => {
  assert.equal(userMessageContent("fix GP"), "fix GP");
  assert.equal(
    userMessageContent("fix GP", "P&L!A1:B6 · 6×2"),
    "fix GP\n\nExcel selection: P&L!A1:B6 · 6×2"
  );
});
