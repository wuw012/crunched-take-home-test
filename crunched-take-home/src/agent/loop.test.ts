import assert from "node:assert/strict";
import { test } from "node:test";
import { ChatMessage } from "../api/types";
import { runTurn } from "./loop";

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
