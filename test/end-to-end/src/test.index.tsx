import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "../../../src/taskpane/components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import { testExcelEnd2End, testPowerPointEnd2End, testWordEnd2End } from "./host-tests";
import * as testHelpers from "./test-helpers";

/* global document, Office, module, require */

const port: number = 4201;

const title = "Contoso Task Pane Add-in";

const rootElement: HTMLElement = document.getElementById("container");
const root = createRoot(rootElement);

/* Render application after Office initializes */
Office.onReady(async (info) => {
  let testValues: any = [];
  try {
    const testServerResponse: { status?: number } = await pingTestServer(port);
    if (testServerResponse?.status === 200) {
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
    } else {
      testHelpers.addErrorResult(testValues, `Ping failed: ${JSON.stringify(testServerResponse)}`);
      await sendTestResults(testValues, port).catch(() => {});
    }
  } catch (err) {
    testHelpers.addErrorResult(testValues, `Initialization failed: ${testHelpers.formatError(err)}`);
    await sendTestResults(testValues, port).catch(() => {});
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("../../../src/taskpane/components/App", () => {
    const NextApp = require("../../../src/taskpane/components/App").default;
    root.render(NextApp);
  });
}
