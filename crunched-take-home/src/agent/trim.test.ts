import assert from "node:assert/strict";
import { test } from "node:test";
import { ChatMessage } from "../api/types";
import { MAX_MESSAGES } from "../shared/limits";
import { trimHistory } from "./trim";

function user(content: string): ChatMessage {
  return { role: "user", content };
}

function assistant(content: string): ChatMessage {
  return { role: "assistant", content };
}

function toolCall(id: string): ChatMessage {
  return {
    role: "assistant",
    content: null,
    tool_calls: [{ id, name: "get_selection", args: {} }],
  };
}

function toolResult(id: string): ChatMessage {
  return { role: "tool", tool_call_id: id, content: "{}" };
}

function assertPairing(messages: ChatMessage[]): void {
  const pending = new Set<string>();
  for (const message of messages) {
    if (message.role === "assistant" && message.tool_calls?.length) {
      for (const call of message.tool_calls) {
        pending.add(call.id);
      }
    } else if (message.role === "tool") {
      assert.ok(pending.has(message.tool_call_id), `orphan tool result ${message.tool_call_id}`);
      pending.delete(message.tool_call_id);
    }
  }
}

test("under-cap history is unchanged", () => {
  const messages = [user("hi"), assistant("hello")];
  assert.deepEqual(trimHistory(messages), messages);
});

test("drops oldest complete turn to fit the cap", () => {
  const messages: ChatMessage[] = [];
  for (let i = 0; i < 7; i += 1) {
    messages.push(user(`q${i}`), assistant(`a${i}`));
  }
  assert.equal(messages.length, 14);

  const trimmed = trimHistory(messages);
  assert.equal(trimmed.length, MAX_MESSAGES);
  assert.deepEqual(trimmed[0], user("q1"));
});

test("never splits an assistant tool_calls message from its tool results", () => {
  const naiveSplit: ChatMessage[] = [
    toolCall("x"),
    toolResult("x"),
    ...Array.from({ length: 11 }, (_, i) => user(`u${i}`)),
  ];
  assert.equal(naiveSplit.length, 13);
  assert.equal(naiveSplit.slice(-MAX_MESSAGES)[0]?.role, "tool");

  const trimmedOldest = trimHistory(naiveSplit);
  assertPairing(trimmedOldest);
  assert.equal(
    trimmedOldest.some((message) => message.role === "tool"),
    false
  );

  const withCurrentPair: ChatMessage[] = [];
  for (let i = 0; i < 10; i += 1) {
    withCurrentPair.push(user(`old-${i}`));
  }
  withCurrentPair.push(user("now"), toolCall("1"), toolResult("1"));
  const trimmedCurrent = trimHistory(withCurrentPair);
  assertPairing(trimmedCurrent);
  assert.ok(
    trimmedCurrent.some((message) => message.role === "assistant" && message.tool_calls?.[0]?.id === "1")
  );
  assert.ok(trimmedCurrent.some((message) => message.role === "tool" && message.tool_call_id === "1"));
});

test("keeps an oversized current turn intact", () => {
  const current: ChatMessage[] = [user("big")];
  for (let i = 0; i < MAX_MESSAGES; i += 1) {
    current.push(assistant(`step-${i}`));
  }
  assert.ok(current.length > MAX_MESSAGES);
  assert.deepEqual(trimHistory(current), current);
});
