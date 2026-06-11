import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider } from "@fluentui/react-components";
import { silksTheme } from "./theme";

/* global document, Office, module, require, HTMLElement */

const title = "Silks AI";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes.
   Browser-preview fallback: outside a Word host, Office.onReady may never fire,
   so race it against a short timeout and render anyway. editor.ts mocks the
   document layer when not in Word, so the UI is fully explorable in a browser. */
let rendered = false;
const renderApp = () => {
  if (rendered) return;
  rendered = true;
  root?.render(
    <FluentProvider theme={silksTheme}>
      <App title={title} />
    </FluentProvider>
  );
};

Office.onReady(renderApp);
setTimeout(renderApp, 1500); // browser-preview fallback

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
