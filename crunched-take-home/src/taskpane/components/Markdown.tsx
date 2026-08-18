import * as React from "react";

function inline(text: string, keyPrefix: string): React.ReactNode[] {
  const chunks = text.split(/(\*\*[^*]+\*\*|`[^`]+`)/g);
  return chunks.map((chunk, index) => {
    const key = `${keyPrefix}-${index}`;
    if (chunk.startsWith("**") && chunk.endsWith("**") && chunk.length >= 4) {
      return <strong key={key}>{chunk.slice(2, -2)}</strong>;
    }
    if (chunk.startsWith("`") && chunk.endsWith("`") && chunk.length >= 2) {
      return <code key={key}>{chunk.slice(1, -1)}</code>;
    }
    return <React.Fragment key={key}>{chunk}</React.Fragment>;
  });
}

export default function Markdown({ text }: { text: string }) {
  const blocks: React.ReactNode[] = [];
  let bullets: string[] = [];

  const flushList = (key: string) => {
    if (bullets.length === 0) {
      return;
    }
    blocks.push(
      <ul key={key} className="md-list">
        {bullets.map((item, index) => (
          <li key={`${key}-${index}`}>{inline(item, `${key}-${index}`)}</li>
        ))}
      </ul>
    );
    bullets = [];
  };

  text.split("\n").forEach((line, index) => {
    const match = line.match(/^\s*[-*]\s+(.*)$/);
    if (match) {
      bullets.push(match[1]);
      return;
    }
    flushList(`list-${index}`);
    if (line.trim() === "") {
      return;
    }
    blocks.push(
      <p key={`p-${index}`} className="md-p">
        {inline(line, `p-${index}`)}
      </p>
    );
  });
  flushList("list-end");

  return <div className="md">{blocks}</div>;
}
