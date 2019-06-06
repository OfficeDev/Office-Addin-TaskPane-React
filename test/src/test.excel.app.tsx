import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from '../../src/taskpane/components/Header';
import HeroList, { HeroListItem } from '../../src/taskpane/components/HeroList';
import Progress from '../../src/taskpane/components/Progress';
import * as excel from "../../src/taskpane/components/excel.App";
import { pingTestServer, sendTestResults } from "office-addin-test-helpers";
import * as testHelpers from "./test-helpers";
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
            listItems: []
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

    async runTest(): Promise<void> {
        return new Promise<void>(async (resolve, reject) => {
            try {
                // Execute taskpane code
                const excelApp = new excel.default(this.props, this.context);
                await excelApp.click();
                await testHelpers.sleep(2000);

                // Get output of executed taskpane code
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    const cellFill = range.format.fill;
                    cellFill.load('color');
                    await context.sync();
                    await testHelpers.sleep(2000);

                    testHelpers.addTestResult(testValues, "fill-color", cellFill.color, "#FFFF00");
                    await sendTestResults(testValues, port);
                    testValues.pop();
                    await testHelpers.closeWorkbook();
                    resolve();
                });
            } catch {
                reject();
            }
        });
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
                    <Button className='ms-welcome__action' buttonType={ButtonType.hero} iconProps={{ iconName: 'ChevronRight' }} >Run</Button>
                </HeroList>
            </div>
        );
    }
}
