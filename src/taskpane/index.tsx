import * as React from "react";
import * as ReactDOM from "react-dom";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, module, require */

const title = "Contoso Task Pane Add-in";

const render = (Component) => {
  ReactDOM.render(
    <FluentProvider theme={webLightTheme}>
      <Component title={title} />
    </FluentProvider>,

    document.getElementById("container")
  );
};

/* Render application after Office initializes */
Office.onReady(() => {
  render(App);
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}
