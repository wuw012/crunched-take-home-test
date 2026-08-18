import { stepChat as postChat } from "../api/client";
import { ChatMessage, StepResponse, ToolCall } from "../api/types";
import { describeTool as describeToolLabel, executeTool as runExcelTool } from "../excel/tools";
import { MAX_STEPS } from "../shared/limits";
import { isCellInspect, isInspectOnly, shouldBlockInspect, toolRoundKey } from "./guards";
import { trimHistory } from "./trim";

export type RunTurnOptions = {
  signal?: AbortSignal;
  selection?: string | null;
  stepChat?: (messages: ChatMessage[], signal?: AbortSignal) => Promise<StepResponse>;
  executeTool?: (name: string, args: Record<string, unknown>) => Promise<string>;
  describeTool?: (name: string, args: Record<string, unknown>) => string;
};

function isAbortError(error: unknown): boolean {
  return error instanceof Error && error.name === "AbortError";
}

function stopped(signal?: AbortSignal): boolean {
  return Boolean(signal?.aborted);
}

function cancelledResult(): string {
  return JSON.stringify({
    error: "cancelled",
    hint: "The user stopped this turn before the tool finished.",
  });
}

export function userMessageContent(userText: string, selection?: string | null): string {
  if (!selection) {
    return userText;
  }
  return `${userText}\n\nExcel selection: ${selection}`;
}

export async function runTurn(
  history: ChatMessage[],
  userText: string,
  onStatus: (status: string) => void,
  onMessages: (messages: ChatMessage[]) => void,
  options: RunTurnOptions = {}
): Promise<ChatMessage[]> {
  const step = options.stepChat ?? postChat;
  const execute = options.executeTool ?? runExcelTool;
  const describe = options.describeTool ?? describeToolLabel;
  const { signal } = options;

  const messages: ChatMessage[] = [
    ...history,
    { role: "user", content: userMessageContent(userText, options.selection) },
  ];
  let previousRound: string | null = null;
  let inspectRounds = 0;

  const publish = () => onMessages([...messages]);

  const halt = (pending: ToolCall[] = []) => {
    for (const call of pending) {
      messages.push({ role: "tool", tool_call_id: call.id, content: cancelledResult() });
    }
    onStatus("Stopping…");
    publish();
    return messages;
  };

  onStatus("Working…");

  for (let round = 0; round < MAX_STEPS; round += 1) {
    if (stopped(signal)) {
      return halt();
    }

    let response: StepResponse;
    try {
      response = await step(trimHistory(messages), signal);
    } catch (error) {
      if (isAbortError(error) || stopped(signal)) {
        return halt();
      }
      throw error;
    }

    if (stopped(signal)) {
      return halt();
    }

    if (response.type === "message") {
      messages.push({ role: "assistant", content: response.content });
      publish();
      return messages;
    }

    const roundKey = toolRoundKey(response.tool_calls);
    const blockInspect = shouldBlockInspect(response.tool_calls, inspectRounds);
    if (roundKey === previousRound && !blockInspect) {
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

    if (stopped(signal)) {
      return halt(response.tool_calls);
    }

    for (let index = 0; index < response.tool_calls.length; index += 1) {
      const call = response.tool_calls[index];
      if (stopped(signal)) {
        return halt(response.tool_calls.slice(index));
      }
      onStatus(`${describe(call.name, call.args)}…`);
      const content = blockInspect
        ? JSON.stringify({
            error: "too_many_reads",
            hint: "You already inspected the workbook. write_range the formulas now. Do not read again.",
          })
        : await execute(call.name, call.args);
      messages.push({ role: "tool", tool_call_id: call.id, content });
      publish();
    }
    if (stopped(signal)) {
      return halt();
    }
    if (!blockInspect) {
      if (isCellInspect(response.tool_calls)) {
        inspectRounds += 1;
      } else if (!isInspectOnly(response.tool_calls)) {
        inspectRounds = 0;
      }
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
