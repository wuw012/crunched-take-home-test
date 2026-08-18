import { ChatMessage } from "../api/types";
import { describeTool, describeToolFailure, parseToolError } from "../excel/tools";

export type AuditLine = { text: string; failed: boolean };

export type VisibleItem =
  | { key: string; kind: "user" | "assistant"; text: string }
  | { key: string; kind: "steps"; lines: AuditLine[] };

function collectToolResults(messages: ChatMessage[], start: number): Map<string, string> {
  const results = new Map<string, string>();
  for (let index = start; index < messages.length; index += 1) {
    const message = messages[index];
    if (message.role !== "tool") {
      break;
    }
    results.set(message.tool_call_id, message.content);
  }
  return results;
}

export function visibleMessages(messages: ChatMessage[]): VisibleItem[] {
  const visible: VisibleItem[] = [];
  messages.forEach((message, index) => {
    if (message.role === "user") {
      visible.push({ key: `${index}-user`, kind: "user", text: message.content });
    }
    if (message.role === "assistant" && message.tool_calls?.length) {
      const results = collectToolResults(messages, index + 1);
      const lines = message.tool_calls.map((call) => {
        const content = results.get(call.id);
        const failed = content ? parseToolError(content) !== null : false;
        return {
          text: failed
            ? describeToolFailure(call.name, call.args, content ?? "")
            : describeTool(call.name, call.args),
          failed,
        };
      });
      const last = visible[visible.length - 1];
      if (last?.kind === "steps") {
        last.lines.push(...lines);
      } else {
        visible.push({ key: `${index}-steps`, kind: "steps", lines });
      }
    }
    if (message.role === "assistant" && message.content) {
      visible.push({ key: `${index}-assistant`, kind: "assistant", text: message.content });
    }
  });
  return visible;
}

export function lastUserText(messages: ChatMessage[]): string {
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (message.role === "user" && message.content.trim()) {
      const text = message.content.trim().replace(/\s+/g, " ");
      return text.length > 72 ? `${text.slice(0, 71)}…` : text;
    }
  }
  return "Open the last thread";
}

export function hasAssistantAfter(list: VisibleItem[], index: number): boolean {
  return list.slice(index + 1).some((item) => item.kind === "assistant");
}

export function hasSuccessfulWrite(messages: ChatMessage[]): boolean {
  for (let index = 0; index < messages.length; index += 1) {
    const message = messages[index];
    if (message.role !== "assistant" || !message.tool_calls?.length) {
      continue;
    }
    const results = collectToolResults(messages, index + 1);
    for (const call of message.tool_calls) {
      if (call.name !== "write_range") {
        continue;
      }
      const content = results.get(call.id);
      if (content && parseToolError(content) === null) {
        return true;
      }
    }
  }
  return false;
}
