import { stepChat as postChat } from "../api/client";
import { ChatMessage, StepResponse } from "../api/types";
import { describeTool as describeToolLabel, executeTool as runExcelTool } from "../excel/tools";
import { MAX_STEPS } from "../shared/limits";
import { isInspectOnly, shouldBlockInspect, toolRoundKey } from "./guards";
import { trimHistory } from "./trim";

export type RunTurnOptions = {
  signal?: AbortSignal;
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

  const messages: ChatMessage[] = [...history, { role: "user", content: userText }];
  let previousRound: string | null = null;
  let inspectRounds = 0;

  const publish = () => onMessages([...messages]);

  const halt = () => {
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

    const blockInspect = shouldBlockInspect(response.tool_calls, inspectRounds);
    for (const call of response.tool_calls) {
      if (stopped(signal)) {
        return halt();
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
