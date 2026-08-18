export type ToolCall = {
  id: string;
  name: string;
  args: Record<string, unknown>;
};

export type ChatMessage =
  | { role: "user"; content: string }
  | { role: "assistant"; content: string | null; tool_calls?: ToolCall[] }
  | { role: "tool"; tool_call_id: string; content: string };

export type StepResponse =
  | { type: "message"; content: string }
  | { type: "tool_calls"; tool_calls: ToolCall[] };
