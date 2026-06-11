import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider } from "@fluentui/react-components";
import { silksTheme } from "./theme";

/* global document, Office, module, require, HTMLElement */

const title = "Silks AI";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={silksTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
