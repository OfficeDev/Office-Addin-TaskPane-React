import * as React from "react";
import { Button, ButtonType } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
/* global Button, console, document, Header, HeroList, HeroListItem, Office, Progress */

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

  click = async () => {
    try {
      // Get the GUID of the selected task
      Office.context.document.getSelectedTaskAsync((result) => {
        let taskGuid;
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          taskGuid = result.value;

          // Set the specified fields for the selected task.
          const targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
          const fieldValues = ["New task name", "Notes for the task."];

          // Set the field value. If the call is successful, set the next field.
          for (let index = 0; index < targetFields.length; index++) {
            Office.context.document.setTaskFieldAsync(taskGuid, targetFields[index], fieldValues[index], (result) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                index++;
              } else {
                console.log(result.error);
              }
            });
          }
        } else {
          console.log(result.error);
        }
      });
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source files, then click <b>Run</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Run
          </Button>
        </HeroList>
      </div>
    );
  }
}
