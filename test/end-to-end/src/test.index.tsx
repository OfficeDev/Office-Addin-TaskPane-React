import * as React from "react";
import * as ReactDOM from "react-dom";
//import App from "./test.app";
import App from "../../../src/taskpane/components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { pingTestServer } from "office-addin-test-helpers";
import { testExcelEnd2End, testPowerPointEnd2End, testWordEnd2End} from "./end2end-tests"

/* global document, Office, module, require */

const port: number = 4201;

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
Office.onReady(async (info) => {
  const testServerResponse: object = await pingTestServer(port);
  if (testServerResponse["status"] == 200) {
    render(App);

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

// if ((module as any).hot) {
//   (module as any).hot.accept("./test.app", () => {
//     const NextApp = require("./test.app").default;
//     render(NextApp);
//   });
// }
