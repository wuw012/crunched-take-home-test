import { ChatMessage } from "../api/types";
import { MAX_MESSAGES } from "../shared/limits";

function turnsOf(messages: ChatMessage[]): ChatMessage[][] {
  const turns: ChatMessage[][] = [];
  for (const message of messages) {
    if (message.role === "user" || turns.length === 0) {
      turns.push([message]);
    } else {
      turns[turns.length - 1].push(message);
    }
  }
  return turns;
}

/** Drop oldest complete user turns until at or under the cap. Never split a tool pair or the current turn. */
export function trimHistory(
  messages: ChatMessage[],
  maxMessages: number = MAX_MESSAGES
): ChatMessage[] {
  if (messages.length <= maxMessages) {
    return messages;
  }
  const turns = turnsOf(messages);
  while (turns.length > 1 && turns.reduce((count, turn) => count + turn.length, 0) > maxMessages) {
    turns.shift();
  }
  return turns.flat();
}
