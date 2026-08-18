import assert from "node:assert/strict";
import { test } from "node:test";
import { chatApiError } from "./client";

test("chatApiError does not dump response bodies", () => {
  assert.equal(chatApiError(503), "Backend is missing ANTHROPIC_API_KEY, or the key is invalid.");
  assert.equal(chatApiError(400), "The chat request was rejected.");
  assert.equal(chatApiError(502), "The model request failed. Check the backend log.");
  assert.equal(chatApiError(500), "Chat API 500");
});
