import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "../../../src/taskpane/components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { pingTestServer } from "office-addin-test-helpers";
import { testExcelEnd2End, testPowerPointEnd2End, testWordEnd2End} from "./host-tests"

/* global document, Office, module, require */

const port: number = 4201;

const title = "Contoso Task Pane Add-in";

const rootElement: HTMLElement = document.getElementById("container");
const root = createRoot(rootElement);

/* Render application after Office initializes */
Office.onReady(async (info) => {
  const testServerResponse: object = await pingTestServer(port);
  if (testServerResponse["status"] == 200) {
    //render(App);
    root.render(
      <FluentProvider theme={webLightTheme}>
        <App title={title} />
      </FluentProvider>
    );

    switch (info.host) {
      case Office.HostType.Excel: {
        return testExcelEnd2End(port);
      }
      case Office.HostType.PowerPoint: {
        return testPowerPointEnd2End(port);
      }
      case Office.HostType.Word: {
        return testWordEnd2End(port);
      }
    }
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("../../../src/taskpane/components/App", () => {
    const NextApp = require("../../../src/taskpane/components/App").default;
    root.render(NextApp);
  });
}


