import * as React from "react";

import {
  Provider,
  Flex,
  Text,
  Button,
  Header,
  List,
  Alert,
  WindowMaximizeIcon,
  ExclamationTriangleIcon,
  Label,
  Input,
  ToDoListIcon,
  Checkbox,
} from "@fluentui/react-northstar";

import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, dialog, UrlDialogInfo } from "@microsoft/teams-js";
import { UserProfile } from "../components/UserProfile";
import { UserEmails } from "../components/UserEmails";

const CenteredWithPadding: React.CSSProperties = {
  padding: "1rem",
  maxWidth: "80vw",
  margin: "0 auto",
};

/**
 * Implementation of the Tab content page
 */
export const MS600TAB_PERSONAL = () => {
  //#region YouTube Player (Task module example)
  const [youTubeVideoId, setYouTubeVideoId] = useState<string | undefined>(
    "VlEH4vtaxp4"
  );
  const [useCard, setUseCard] = useState<boolean>(false);

  const appRoot = (): string => {
    if (typeof window === "undefined") {
      return "https://{{HOSTNAME}}";
    } else {
      return window.location.protocol + "//" + window.location.host;
    }
  };

  const onShowVideo = (): void => {
    const dialogInfo: UrlDialogInfo = {
      title: "YouTube Player",
      url: appRoot() + `/MS600TAB_PERSONAL/player.html?vid=${youTubeVideoId}`,
      fallbackUrl:
        appRoot() + `/MS600TAB_PERSONAL/player.html?vid=${youTubeVideoId}`,
      size: {
        width: 1000,
        height: 700,
      },
    };

    dialog.url.open(dialogInfo);
  };

  const onChangeVideo = (): void => {
    const dialogInfo = {
      title: "YouTube Video Selector",
      url:
        appRoot() +
        `/MS600TAB_PERSONAL/selector.html?theme={theme}&vid=${youTubeVideoId}`,
      size: {
        width: 350,
        height: 150,
      },
    };

    const submitHandler: dialog.DialogSubmitHandler = (response) => {
      console.log(`Submit handler - err: ${response.err}`);
      setYouTubeVideoId(response.result?.toString());
    };

    dialog.url.open(dialogInfo, submitHandler);
  };

  //#endregion

  //#region PersonalTab example

  const [{ inTeams, theme, context }] = useTeams();
  const [entityId, setEntityId] = useState<string | undefined>();

  const [todoItems, setTodoItems] = useState<string[]>([
    "Submit time sheet",
    "Submit expense report",
  ]);
  const [newTodoValue, setNewTodoValue] = useState<string>("");

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

  const handleOnChanged = (event): void => {
    setNewTodoValue(event.target.value);
  };

  const handleOnClick = (event: React.MouseEvent<HTMLButtonElement>): void => {
    const newTodoItems = todoItems;
    newTodoItems.push(newTodoValue);
    setTodoItems(newTodoItems);
    setNewTodoValue("");
  };
  //#endregion

  //#region Authentication example
  // const getAccessToken = async (promptConsent: boolean = false): Promise<string> => {
  //   try {
  //     const accessToken = await authentication.authenticate({
  //       url: window.location.origin + "/auth-start.html",
  //       width: 600,
  //       height: 535
  //     });
  //     return Promise.resolve(accessToken);
  //   } catch (error) {
  //     return Promise.reject(error);
  //   }
  // };
  //#endregion

  return (
    <Provider theme={theme}>
      <Flex column gap="gap.smaller" style={CenteredWithPadding}>
        <Header content="Personal TAB" />
        <Alert
          icon={<ExclamationTriangleIcon />}
          content={`EntityID: ${entityId}`}
          dismissible
        />
      </Flex>

      {inTeams && (
        <>
          <Flex column gap="gap.smaller" style={CenteredWithPadding}>
            <fieldset>
              <legend>TASK MODULES</legend>

              <Flex.Item>
                <div>
                  <div>
                    <Flex column gap="gap.smaller">
                      <Flex.Item>
                        <div style={{ display: "flex", alignItems: "center" }}>
                          <Checkbox
                            label="Use Card"
                            checked={useCard}
                            onChange={() => setUseCard((u) => !u)}
                          />
                          <strong>
                            {useCard
                              ? `I'll use Card`
                              : `I'll use an html page`}
                          </strong>
                        </div>
                      </Flex.Item>

                      <Flex.Item>
                        <div>
                          <Text style={{marginRight: "1rem"}}>YouTube Video ID:</Text>
                          <Input value={youTubeVideoId} disabled></Input>
                        </div>
                      </Flex.Item>
                    </Flex>
                  </div>
                  <Flex gap="gap.smaller">
                    <Button
                      content="Change Video ID"
                      style={{marginRight: "1rem"}}
                      onClick={() => onChangeVideo()}
                    />
                    <Button
                      primary
                      content="Show Video"
                      onClick={() => onShowVideo()}
                    />
                  </Flex>
                </div>
              </Flex.Item>

              <Flex.Item
                styles={{
                  padding: ".8rem 0 .8rem .5rem",
                }}
              >
                <Text content="(C) Copyright Contoso" size="smaller"></Text>
              </Flex.Item>
            </fieldset>
          </Flex>

          <Flex column gap="gap.smaller" style={CenteredWithPadding}>
            <UserProfile />
          </Flex>

          <Flex column gap="gap.smaller" style={CenteredWithPadding}>
            <UserEmails />
          </Flex>
        </>
      )}
    </Provider>
  );
};
