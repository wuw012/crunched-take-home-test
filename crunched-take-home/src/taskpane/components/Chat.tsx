import * as React from "react";
import { runTurn } from "../../agent/loop";
import { ChatMessage } from "../../api/types";
import { describeTool } from "../../excel/tools";
import Markdown from "./Markdown";

type Starter = { label: string; hint: string; prompt: string };

const STARTERS: Starter[] = [
  {
    label: "Orient the workbook",
    hint: "List sheets and sizes — do not ingest Exports",
    prompt: "What sheets are in this workbook and how large is each used range? Do not read all the data.",
  },
  {
    label: "Error-check the P&L",
    hint: "Gross Profit does not foot — write formulas",
    prompt:
      "Gross Profit does not foot. Find the error and fix Gross Profit and Operating Profit the way a modeler would.",
  },
  {
    label: "Link revenue to drivers",
    hint: "FY24 Revenue = Assumptions Price × Units",
    prompt: "Revenue is hardcoded. Drive FY24 Revenue from Assumptions: Price × Units.",
  },
  {
    label: "Chart the P&L",
    hint: "Clustered column, then switch it to a bar chart",
    prompt: "Chart this P&L as a clustered column. Then make it a bar chart.",
  },
  {
    label: "Stub a 3-year forecast",
    hint: "FY25–FY27 in D–F, YoY from Assumptions, formulas only",
    prompt:
      "Add a 3-year forecast in D–F: FY25–FY27 revenue growing at the Assumptions YoY growth rate. Formulas only, linked to FY24 Revenue.",
  },
];

const TEXTAREA_MAX_PX = 132;

type Screen = "home" | "thread";

type VisibleItem =
  | { key: string; kind: "user" | "assistant"; text: string }
  | { key: string; kind: "steps"; lines: string[] };

function visibleMessages(messages: ChatMessage[]): VisibleItem[] {
  const visible: VisibleItem[] = [];
  messages.forEach((message, index) => {
    if (message.role === "user") {
      visible.push({ key: `${index}-user`, kind: "user", text: message.content });
    }
    if (message.role === "assistant" && message.tool_calls?.length) {
      const lines = message.tool_calls.map((call) => describeTool(call.name, call.args));
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

function lastUserText(messages: ChatMessage[]): string {
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (message.role === "user" && message.content.trim()) {
      const text = message.content.trim().replace(/\s+/g, " ");
      return text.length > 72 ? `${text.slice(0, 71)}…` : text;
    }
  }
  return "Open the last thread";
}

function hasAssistantAfter(list: VisibleItem[], index: number): boolean {
  return list.slice(index + 1).some((item) => item.kind === "assistant");
}

function StarterButton(props: { label: string; hint: string; onClick: () => void }) {
  return (
    <button type="button" className="starter" onClick={props.onClick}>
      <span>
        <span className="starter__label">{props.label}</span>
        <span className="starter__hint">{props.hint}</span>
      </span>
      <span className="starter__arrow" aria-hidden="true">
        →
      </span>
    </button>
  );
}

const Chat: React.FC = () => {
  const [screen, setScreen] = React.useState<Screen>("home");
  const [messages, setMessages] = React.useState<ChatMessage[]>([]);
  const [draft, setDraft] = React.useState("");
  const [status, setStatus] = React.useState<string | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [composerFocused, setComposerFocused] = React.useState(false);
  const [expandedAudits, setExpandedAudits] = React.useState<Record<string, boolean>>({});
  const scroller = React.useRef<HTMLDivElement>(null);
  const textarea = React.useRef<HTMLTextAreaElement>(null);
  const turnId = React.useRef(0);
  const busy = status !== null;
  const list = visibleMessages(messages);

  React.useEffect(() => {
    if (screen !== "thread") {
      return;
    }
    scroller.current?.scrollTo({ top: scroller.current.scrollHeight, behavior: "smooth" });
  }, [list, status, error, screen]);

  React.useEffect(() => {
    const el = textarea.current;
    if (!el) {
      return;
    }
    el.style.height = "auto";
    el.style.height = `${Math.min(el.scrollHeight, TEXTAREA_MAX_PX)}px`;
  }, [draft]);

  async function send(text: string, mode: "continue" | "new" = "continue") {
    if (!text) {
      return;
    }
    if (busy && mode === "continue") {
      return;
    }
    const history = mode === "new" ? [] : messages;
    const id = ++turnId.current;
    setDraft("");
    setError(null);
    setScreen("thread");
    setMessages([...history, { role: "user", content: text }]);
    setStatus("Working…");
    try {
      const next = await runTurn(
        history,
        text,
        (nextStatus) => {
          if (id === turnId.current) {
            setStatus(nextStatus);
          }
        },
        (nextMessages) => {
          if (id === turnId.current) {
            setMessages(nextMessages);
          }
        }
      );
      if (id !== turnId.current) {
        return;
      }
      setMessages(next);
    } catch (err) {
      if (id !== turnId.current) {
        return;
      }
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      if (id === turnId.current) {
        setStatus(null);
      }
    }
  }

  function goHome() {
    setScreen("home");
  }

  function stopTurn() {
    turnId.current += 1;
    setStatus(null);
  }

  function onSubmit(event: React.FormEvent) {
    event.preventDefault();
    void send(draft.trim(), "continue");
  }

  function toggleAudit(key: string, collapsedDefault: boolean) {
    setExpandedAudits((current) => {
      const open = current[key] ?? !collapsedDefault;
      return { ...current, [key]: !open };
    });
  }

  return (
    <div className="chat" aria-busy={busy}>
      <header className="chat__header">
        <button
          type="button"
          className="chat__back"
          onClick={goHome}
          tabIndex={screen === "thread" ? 0 : -1}
          aria-hidden={screen !== "thread"}
          disabled={screen !== "thread"}
        >
          <span aria-hidden="true">&lt;</span> Back
        </button>
        <div className="brand">
          <img className="brand__mark" src="assets/logo-oxblood.png" alt="" width={16} height={16} />
          <h1 className="brand__name">Crunched</h1>
        </div>
      </header>

      <div className="chat__messages" ref={scroller}>
        {screen === "home" ? (
          <div className="chat__empty">
            <div className="starters">
              {messages.length > 0 ? (
                <StarterButton
                  label="Continue last chat"
                  hint={lastUserText(messages)}
                  onClick={() => setScreen("thread")}
                />
              ) : null}
              {STARTERS.map((starter) => (
                <StarterButton
                  key={starter.label}
                  label={starter.label}
                  hint={starter.hint}
                  onClick={() => void send(starter.prompt, "new")}
                />
              ))}
            </div>
          </div>
        ) : (
          <>
            {list.map((item, index) => {
              if (item.kind === "steps") {
                const collapsedDefault = hasAssistantAfter(list, index);
                const open = expandedAudits[item.key] ?? !collapsedDefault;
                if (!open) {
                  return (
                    <button
                      key={item.key}
                      type="button"
                      className="chat__audit-summary"
                      onClick={() => toggleAudit(item.key, collapsedDefault)}
                    >
                      {item.lines.length} workbook {item.lines.length === 1 ? "action" : "actions"}
                    </button>
                  );
                }
                return (
                  <div key={item.key} className="chat__audit">
                    <button
                      type="button"
                      className="chat__audit-label"
                      onClick={() => toggleAudit(item.key, collapsedDefault)}
                    >
                      Workbook actions
                    </button>
                    {item.lines.map((line, lineIndex) => (
                      <div key={`${item.key}-${lineIndex}`} className="chat__step">
                        <span className="chat__step-index">{String(lineIndex + 1).padStart(2, "0")}</span>
                        <span>{line}</span>
                      </div>
                    ))}
                  </div>
                );
              }
              return (
                <div key={item.key} className={`bubble bubble--${item.kind}`}>
                  {item.kind === "assistant" ? <Markdown text={item.text} /> : item.text}
                </div>
              );
            })}
            {error ? <div className="chat__error">{error}</div> : null}
          </>
        )}
      </div>

      <form className="chat__composer" onSubmit={onSubmit}>
        <div className="composer">
          <textarea
            ref={textarea}
            value={draft}
            disabled={busy && screen === "thread"}
            placeholder="Ask to inspect, error-check, or write formulas…"
            rows={1}
            aria-label="Message"
            onChange={(event) => setDraft(event.target.value)}
            onFocus={() => setComposerFocused(true)}
            onBlur={() => setComposerFocused(false)}
            onKeyDown={(event) => {
              if (event.key === "Enter" && !event.shiftKey) {
                event.preventDefault();
                event.currentTarget.form?.requestSubmit();
              }
            }}
          />
          <div className="composer__bar">
            {screen === "thread" && busy ? (
              <>
                <span className="composer__status" aria-live="polite">
                  {status}
                </span>
                <button type="button" className="composer__stop" onClick={stopTurn}>
                  Stop
                </button>
              </>
            ) : (
              <>
                <span className="composer__hint">{composerFocused ? "Enter to send" : ""}</span>
                <button type="submit" disabled={busy || !draft.trim()}>
                  Send
                </button>
              </>
            )}
          </div>
        </div>
      </form>
    </div>
  );
};

export default Chat;
