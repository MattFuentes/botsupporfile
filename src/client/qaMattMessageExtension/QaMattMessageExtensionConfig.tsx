import * as React from "react";
import { Provider, Flex, Header, Checkbox, Button } from "@fluentui/react-northstar";
import { app } from "@microsoft/teams-js";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";

/**
 * Implementation of the qa-matt Message Extension configuration page
 */
export const QaMattMessageExtensionConfig = () => {

    const [{ inTeams, theme }] = useTeams();
    const [onOrOff, setOnOrOff] = useState<boolean>();

    useEffect(() => {
        if (inTeams === true) {
            app.notifySuccess();
            setOnOrOff(true);
        }
    }, [inTeams]);

    return (
        <Provider theme={theme} styles={{ height: "100vh", width: "100vw", padding: "1em" }}>
            <Flex fill={true}>
                <Flex.Item>
                    <div>
                        <Header content="qa-matt Message Extension configuration" />
                        <Checkbox
                            label="On or off?"
                            toggle
                            checked={onOrOff}
                            onChange={() => { setOnOrOff(!onOrOff); }} />
                        <Button onClick={() => app.notifySuccess()} primary>OK</Button>
                    </div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
