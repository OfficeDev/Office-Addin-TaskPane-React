import App from "./test.app";
import { initializeIcons } from "@fluentui/font-icons-mdl2";
import * as React from "react";
import * as ReactDOM from "react-dom";
/* global document, Office, module, require */

initializeIcons();

let isOfficeInitialized = false;

const title = "Contoso Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <Component title={title} isOfficeInitialized={isOfficeInitialized} />,
    document.getElementById("container")
  );
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
