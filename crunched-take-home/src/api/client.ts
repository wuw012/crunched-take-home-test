import { ChatMessage, StepResponse } from "./types";

export async function stepChat(messages: ChatMessage[]): Promise<StepResponse> {
  const response = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ messages }),
  });
  if (!response.ok) {
    const body = await response.text();
    throw new Error(`Chat API ${response.status}: ${body}`);
  }
  return (await response.json()) as StepResponse;
}
