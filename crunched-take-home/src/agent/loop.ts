import { stepChat } from "../api/client";
import { ChatMessage, ToolCall } from "../api/types";
import { describeTool, executeTool } from "../excel/tools";
import { MAX_STEPS } from "../shared/limits";
import { trimHistory } from "./trim";

const MAX_INSPECT_ROUNDS = 3;
const INSPECT_TOOLS = new Set(["list_workbook_structure", "get_selection", "read_range"]);

function toolRoundKey(calls: ToolCall[]): string {
  return JSON.stringify(calls.map((call) => ({ name: call.name, args: call.args })));
}

function isInspectOnly(calls: ToolCall[]): boolean {
  return calls.length > 0 && calls.every((call) => INSPECT_TOOLS.has(call.name));
}

export async function runTurn(
  history: ChatMessage[],
  userText: string,
  onStatus: (status: string) => void,
  onMessages: (messages: ChatMessage[]) => void
): Promise<ChatMessage[]> {
  const messages: ChatMessage[] = [...history, { role: "user", content: userText }];
  let previousRound: string | null = null;
  let inspectRounds = 0;

  const publish = () => onMessages([...messages]);

  onStatus("Working…");

  for (let step = 0; step < MAX_STEPS; step += 1) {
    const response = await stepChat(trimHistory(messages));

    if (response.type === "message") {
      messages.push({ role: "assistant", content: response.content });
      publish();
      return messages;
    }

    const roundKey = toolRoundKey(response.tool_calls);
    if (roundKey === previousRound) {
      messages.push({
        role: "assistant",
        content: "Stopped: the same tool call repeated. Send another message to continue.",
      });
      publish();
      return messages;
    }
    previousRound = roundKey;

    messages.push({ role: "assistant", content: null, tool_calls: response.tool_calls });
    publish();

    const blockInspect = isInspectOnly(response.tool_calls) && inspectRounds >= MAX_INSPECT_ROUNDS;
    for (const call of response.tool_calls) {
      onStatus(`${describeTool(call.name, call.args)}…`);
      const content = blockInspect
        ? JSON.stringify({
            error: "too_many_reads",
            hint: "You already inspected the workbook. write_range the formulas now. Do not read again.",
          })
        : await executeTool(call.name, call.args);
      messages.push({ role: "tool", tool_call_id: call.id, content });
      publish();
    }
    if (!blockInspect) {
      inspectRounds = isInspectOnly(response.tool_calls) ? inspectRounds + 1 : 0;
    }
    onStatus("Working…");
  }

  messages.push({
    role: "assistant",
    content: `Stopped after ${MAX_STEPS} tool rounds. Send another message to continue.`,
  });
  publish();
  return messages;
}
