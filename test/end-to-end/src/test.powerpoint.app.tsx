import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "../../../src/taskpane/components/Header";
import HeroList, { HeroListItem } from "../../../src/taskpane/components/HeroList";
import Progress from "../../../src/taskpane/components/Progress";
import * as powerpoint from "../../../src/taskpane/components/PowerPoint.App";
import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import * as testHelpers from "./test-helpers";

/* global Office, PowerPoint, require */
const port: number = 4201;
let testValues: any = [];

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
    Office.onReady(async () => {
      const testServerResponse: object = await pingTestServer(port);
      if (testServerResponse["status"] == 200) {
        this.runTest();
      }
    });
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration",
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality",
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro",
        },
      ],
    });
  }

  async runTest(): Promise<void> {
    try {
      // Execute taskpane code
      const powerpointApp = new powerpoint.default(this.props, this.context);
      await powerpointApp.click();
      await testHelpers.sleep(2000);

      // Get output of executed taskpane code
      PowerPoint.run(async () => {
        // get selected text
        const selectedText = await this.getSelectedText();
        // send test results
        testHelpers.addTestResult(testValues, "output-message", selectedText, " Hello World!");
        await sendTestResults(testValues, port);
        testValues.pop();
        Promise.resolve();
      });
    } catch {
      Promise.reject();
    }
  }

  async getSelectedText(): Promise<string> {
    return new Promise((resolve, reject) => {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result: Office.AsyncResult<string>) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(result.error);
        } else {
          resolve(result.value);
        }
      });
    });
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }}>
            Run
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}
