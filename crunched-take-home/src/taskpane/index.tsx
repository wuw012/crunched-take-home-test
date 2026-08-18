import * as React from "react";
import { createRoot } from "react-dom/client";
import Chat from "./components/Chat";

/* global document, Office, module, require, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  root?.render(<Chat />);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/Chat", () => {
    const NextChat = require("./components/Chat").default;
    root?.render(<NextChat />);
  });
}
