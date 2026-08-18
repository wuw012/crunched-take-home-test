import { ChatMessage, StepResponse } from "./types";

export function chatApiError(status: number): string {
  if (status === 503) {
    return "Backend is missing ANTHROPIC_API_KEY, or the key is invalid.";
  }
  if (status === 429) {
    return "The model is rate limited. Try again in a moment.";
  }
  if (status === 502) {
    return "The model request failed. Check the backend log.";
  }
  if (status === 400) {
    return "The chat request was rejected.";
  }
  return `Chat API ${status}`;
}

export async function stepChat(messages: ChatMessage[], signal?: AbortSignal): Promise<StepResponse> {
  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ messages }),
    signal,
  });
  if (!response.ok) {
    throw new Error(chatApiError(response.status));
  }
  return (await response.json()) as StepResponse;
}
