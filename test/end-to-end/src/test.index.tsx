import App from "./test.app";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import React from "react";
import { createRoot } from "react-dom/client";
/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const container = document.getElementById("container");
const root = createRoot(container);
const render = (Component) => {
  root.render(<Component title={title} isOfficeInitialized={isOfficeInitialized} />);
};

/* Render application after Office initializes */
Office.initialize = () => {
  isOfficeInitialized = true;
  render(App);
};

/* Initial render showing a progress bar */
render(App);

if ((module as any).hot) {
  (module as any).hot.accept("./test.app", () => {
    const NextApp = require("./test.app").default;
    render(NextApp);
  });
}
