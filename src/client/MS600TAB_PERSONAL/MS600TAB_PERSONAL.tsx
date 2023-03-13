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
} from "@fluentui/react-northstar";

import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, dialog, UrlDialogInfo } from "@microsoft/teams-js";
import { UserProfile } from "../components/UserProfile";
import { UserEmails } from "../components/UserEmails";

const CenteredWithPadding: React.CSSProperties = {
  padding : "1rem",
  maxWidth: "80vw",
  margin  : "0 auto"
};

/**
 * Implementation of the Tab content page
 */
export const MS600TAB_PERSONAL = () => {
  //#region YouTube Player (Task module example)
  const [youTubeVideoId, setYouTubeVideoId] = useState<string | undefined>("VlEH4vtaxp4");
  
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
      fallbackUrl: appRoot() + `/MS600TAB_PERSONAL/player.html?vid=${youTubeVideoId}`,
      size: {
        width: 1000,
        height: 700
      }
    };

    dialog.url.open(dialogInfo);
  };
  
  const onChangeVideo = (): void => {

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

      {
        inTeams && <>
          <Flex column gap="gap.smaller" style={CenteredWithPadding}>
            <Flex.Item>
              <div>
                <div>
                  <Text>YouTube Video ID:</Text>
                  <Input value={youTubeVideoId} disabled></Input>
                </div>
                <div>
                  <Button content="Change Video ID" onClick={() => onChangeVideo()}></Button>
                  <Button content="Show Video" primary onClick={() => onShowVideo()}></Button>
                </div>
              </div>
            </Flex.Item>

            <Flex.Item styles={{
              padding: ".8rem 0 .8rem .5rem"
            }}>
              <Text content="(C) Copyright Contoso" size="smaller"></Text>
            </Flex.Item>
          </Flex>

          <UserProfile />
          <UserEmails />
        </>
      }
      
      <Flex column gap="gap.smaller" style={CenteredWithPadding}>
        <Header content="This is your tab" />
        <Alert
          icon={<ExclamationTriangleIcon />}
          content={entityId}
          dismissible
        ></Alert>
        <Text content="These are your to-do items:" size="medium"></Text>
        <List selectable>
          {todoItems.map((todoItem, i) => (
            <List.Item
              key={i}
              media={<WindowMaximizeIcon outline />}
              content={todoItem}
              index={i}
            ></List.Item>
          ))}
        </List>

        <Flex gap="gap.medium">
          <Flex.Item grow>
            <Flex>
              <Label
                icon={<ToDoListIcon />}
                styles={{
                  background: "darkgray",
                  height: "auto",
                  padding: "0 15px",
                }}
              ></Label>
              <Flex.Item grow>
                <Input
                  placeholder="New todo item"
                  fluid
                  value={newTodoValue}
                  onChange={handleOnChanged}
                ></Input>
              </Flex.Item>
            </Flex>
          </Flex.Item>
          <Button content="Add Todo" primary onClick={handleOnClick}></Button>
        </Flex>
        <Text content="(C) Copyright Contoso" size="smallest"></Text>
      </Flex>
    </Provider>
  );
};
