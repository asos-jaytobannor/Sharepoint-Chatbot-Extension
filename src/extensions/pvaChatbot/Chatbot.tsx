import * as React from "react";
import { useBoolean, useId } from '@uifabric/react-hooks';
import * as ReactWebChat from 'botframework-webchat';
import { Dialog, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Spinner } from 'office-ui-fabric-react/lib/Spinner';

export interface IChatbotProps { }

const dialogContentProps = {
    type: DialogType.normal,
    title: 'Cloudio',
    closeButtonAriaLabel: 'Close'
};
export const PVAChatbotDialog: React.FunctionComponent = () => {
    const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
    const labelId: string = useId('dialogLabel');
    const subTextId: string = useId('subTextLabel');
    const modalProps = React.useMemo(
        () => ({
            isBlocking: false,
        }),
        [labelId, subTextId],
    );
    const BOT_ID = "1c1645d6-a78a-452b-a451-34f13845fd8f";
    const theURL = "https://powerva.microsoft.com/api/botmanagement/v1/directline/directlinetoken?botId=" + BOT_ID;
    const store = ReactWebChat.createStore(
        {},
        ({ dispatch }) => next => action => {
            return next(action);
        }
    );
    fetch(theURL)
        .then(response => response.json())
        .then(conversationInfo => {
            document.getElementById("loading-spinner").style.display = 'none';
            document.getElementById("webchat").style.minHeight = '50vh';
            ReactWebChat.renderWebChat(
                {
                    directLine: ReactWebChat.createDirectLine({
                        token: conversationInfo.token,
                    }),
                    store: store,
                },
                document.getElementById('webchat')
            );
        })
        .catch(err => console.error("An error occurred: " + err));

    return (
        <>
            <DefaultButton secondaryText="Opens the Chatbot Dialog" onClick={toggleHideDialog} text="Open Cloudio Chatbot" />
            <Dialog styles={{
                main: { selectors: { ['@media (min-width: 480px)']: { width: 450, minWidth: 450, maxWidth: '1000px' } } }
            }} hidden={hideDialog} onDismiss={toggleHideDialog} dialogContentProps={dialogContentProps} modalProps={modalProps}>
                <div id="chatContainer" style={{ display: "flex", flexDirection: "column", alignItems: "center" }}>
                    <div id="webchat" role="main" style={{ width: "100%", height: "0rem" }}></div>
                    <Spinner id="loading-spinner" label="Loading..." style={{ paddingTop: "1rem", paddingBottom: "1rem" }} />
                </div>
            </Dialog>
        </>
    );
};

export default class Chatbot extends React.Component<IChatbotProps> {
    constructor(props: IChatbotProps) {
        super(props);
    }
    public render(): JSX.Element {
        return (
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", paddingBottom: "1rem" }}>
                <PVAChatbotDialog />
            </div>
        );
    }
}  