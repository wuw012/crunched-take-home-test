import * as React from "react";
import { runTurn } from "../../agent/loop";
import { ChatMessage } from "../../api/types";
import { refreshSelectionChip } from "../selection";
import { STARTERS } from "../starters";
import {
  hasAssistantAfter,
  hasSuccessfulWrite,
  lastUserText,
  visibleMessages,
} from "../viewModel";
import Markdown from "./Markdown";

const TEXTAREA_MAX_PX = 132;

type Screen = "home" | "thread";

function StarterButton(props: {
  label: string;
  hint: string;
  disabled?: boolean;
  onClick: () => void;
}) {
  return (
    <button type="button" className="starter" disabled={props.disabled} onClick={props.onClick}>
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
  const [selection, setSelection] = React.useState<string | null>(null);
  const scroller = React.useRef<HTMLDivElement>(null);
  const textarea = React.useRef<HTMLTextAreaElement>(null);
  const turnId = React.useRef(0);
  const abortRef = React.useRef<AbortController | null>(null);
  const busy = status !== null;
  const list = visibleMessages(messages);
  const wrote = hasSuccessfulWrite(messages);

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

  React.useEffect(() => {
    let cancelled = false;
    const refresh = async () => {
      const chip = await refreshSelectionChip();
      if (!cancelled) {
        setSelection(chip);
      }
    };
    const onFocus = () => {
      void refresh();
    };
    window.addEventListener("focus", onFocus);
    void refresh();
    return () => {
      cancelled = true;
      window.removeEventListener("focus", onFocus);
    };
  }, []);

  async function send(text: string, mode: "continue" | "new" = "continue") {
    if (!text || busy) {
      return;
    }
    const history = mode === "new" ? [] : messages;
    const controller = new AbortController();
    abortRef.current = controller;
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
        },
        { signal: controller.signal }
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
        if (abortRef.current === controller) {
          abortRef.current = null;
        }
      }
    }
  }

  function goHome() {
    setScreen("home");
  }

  function stopTurn() {
    abortRef.current?.abort();
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
        {selection ? (
          <div className="chat__selection" title="Current Excel selection">
            {selection}
          </div>
        ) : null}
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
                  disabled={busy}
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
                      <div
                        key={`${item.key}-${lineIndex}`}
                        className={line.failed ? "chat__step chat__step--failed" : "chat__step"}
                      >
                        <span className="chat__step-index">{String(lineIndex + 1).padStart(2, "0")}</span>
                        <span>{line.text}</span>
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
            {wrote ? (
              <div className="chat__notice">
                Changes are in the open file. Excel Undo (⌘Z or Ctrl+Z) reverts.
              </div>
            ) : null}
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
