import * as React from "react";
import { Provider, Flex, Text, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app } from "@microsoft/teams-js";

/**
 * Implementation of the MS600ConfigTab content page
 */
export const MS600ConfigTab = () => {
  const [{ inTeams, theme, context }] = useTeams();
  const [entityId, setEntityId] = useState<string | undefined>();

  useEffect(() => {
    if (inTeams === true) {
      app.notifySuccess();
    } else {
      setEntityId("Not in Microsoft Teams");
    }
  }, [inTeams]);

  useEffect(() => {
    if (context) {
      setEntityId(context.page.id);
    }
  }, [context]);

  /**
   * The render() method to create the UI of the tab
   */
  return (
    <Provider theme={theme}>
      <Flex
        fill={true}
        column
        styles={{
          padding: ".8rem 0 .8rem .5rem",
        }}
      >
        <Flex.Item>
          <Header content="This is your tab" />
        </Flex.Item>
        <Flex.Item>
          <div>
            {inTeams ? (
              <div>
                <Text content="Siamo dentro Teams." />
              </div>
            ) : (
              <div>
                <Text content="Non siamo dentro Teams." />
              </div>
            )}
          </div>
        </Flex.Item>
        <Flex.Item
          styles={{
            padding: ".8rem 0 .8rem .5rem",
          }}
        >
          <Text size="smaller" content="(C) Copyright companyname" />
        </Flex.Item>
      </Flex>
    </Provider>
  );
};
