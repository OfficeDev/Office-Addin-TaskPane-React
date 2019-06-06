import * as React from 'react';
import { Button, ButtonType } from 'office-ui-fabric-react';
import Header from '../../src/taskpane/components/Header';
import HeroList, { HeroListItem } from '../../src/taskpane/components/HeroList';
import Progress from '../../src/taskpane/components/Progress';
import * as excel from "./test.excel.app";
import * as word from "./test.word.app";
import { pingTestServer} from "office-addin-test-helpers";
const port: number = 4201;

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
        Office.onReady(async (info) => {
            if (info.host === Office.HostType.Excel || info.host === Office.HostType.Word) {
                const testServerResponse: object = await pingTestServer(port);
                if (testServerResponse["status"] == 200) {
                    if (info.host === Office.HostType.Excel) {
                        const excelApp = new excel.default(this.props, this.context);
                        return excelApp.runTest();
                    } else {
                        const wordApp = new word.default(this.props, this.context);
                        return wordApp.runTest();
                    }
                }
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
    click = async () => { 
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
}
