import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from './Header';
import HeroList, { HeroListItem } from './HeroList';
import Progress from './Progress';

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
            listItems: []
        };
    }

    componentDidMount() {
        this.setState({
            listItems: [
                {
                    icon: 'Ribbon',
                    primaryText: 'Achieve more with Office integration'
                },
                {
                    icon: 'Unlock',
                    primaryText: 'Unlock features and functionality'
                },
                {
                    icon: 'Design',
                    primaryText: 'Create and visualize like a pro'
                }
            ]
        });
    }

    click = async () => {
        switch (Office.context.host) {
            case Office.HostType.Excel:
              return this.runExcel();
            case Office.HostType.OneNote:
              return this.runOneNote();
            case Office.HostType.Outlook:
              return this.runOutlook();
            case Office.HostType.PowerPoint:
              return this.runPowerPoint();
            case Office.HostType.Project:
              return this.runProject();
            case Office.HostType.Word:
              return this.runWord();
          }
    }

    render() {
        const {
            title,
            isOfficeInitialized,
        } = this.props;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/logo-filled.png'
                    message='Please sideload your addin to see app body.'
                />
            );
        }

        return (
            <div className='ms-welcome'>
                <Header logo='assets/logo-filled.png' title={this.props.title} message='Welcome' />
                <HeroList message='Discover what Office Add-ins can do for you today!' items={this.state.listItems}>
                    <p className='ms-font-l'>Modify the source files, then click <b>Run</b>.</p>
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} onClick={this.click}>Run</Button>
                </HeroList>
            </div>
        );
    }

     runExcel = async () => {
        try {
          await Excel.run(async context => {
            /**
             * Insert your Excel code here
             */
            const range = context.workbook.getSelectedRange();
      
            // Read the range address
            range.load("address");
      
            // Update the fill color
            range.format.fill.color = "yellow";
      
            await context.sync();
            console.log(`The range address was ${range.address}.`);
          });
        } catch (error) {
          console.log(error);
        }
      }
      
      runOneNote = async () => {
        /**
         * Insert your OneNote code here
         */
      }
      
      
      runOutlook = async () => {
        /**
         * Insert your Outlook code here
         */
      }
      
      runPowerPoint = async () => {
        /**
         * Insert your PowerPoint code here
         */
        Office.context.document.setSelectedDataAsync("Hello World!",
          {
            coercionType: Office.CoercionType.Text
          },
          result => {
            if (result.status === Office.AsyncResultStatus.Failed) {
              console.error(result.error.message);
            }
          }
        );
      }
      
      runProject = async () =>{
        /**
         * Insert your Outlook code here
         */
      }
      
      runWord = async () => {
        return Word.run(async context => {
          /**
           * Insert your Word code here
           */
          const range = context.document.getSelection();
      
          // Read the range text
          range.load("text");
      
          // Update font color
          range.font.color = "red";
      
          await context.sync();
          console.log(`The selected text was ${range.text}.`);
        });
      }
}
