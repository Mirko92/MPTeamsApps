import * as React from "react";

import {
  Provider,
  Flex,
  Text,
  Button,
  Header,
  Alert,
  ExclamationTriangleIcon,
  Input,
  Checkbox,
} from "@fluentui/react-northstar";

import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import { app, dialog, UrlDialogInfo, tasks, DialogInfo } from "@microsoft/teams-js";
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
    "mmw57bp8AGI"
  );
  const [useCard, setUseCard] = useState<boolean>(false);
  const [formDisabled, setFormDisabled] = useState<boolean>(true);

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


  const openUrlDialog = (): void => {
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
      
      if (response.result) {
        setYouTubeVideoId(response.result?.toString());
      }
    };

    dialog.url.open(dialogInfo, submitHandler);
  }

  
  const openCardDialog = (): void => {
    const card = require("./YouTubeSelectorCard.json");

    card.body
      .find(x => x.id === "FormContainer")
      .items
      .find(x => x.id === "youTubeVideoId")
      .value = youTubeVideoId;

    const taskModuleInfo: DialogInfo = {
      title: "YouTube Video Selector",
      width: 350,
      height: 250,
      card
    };

    tasks.startTask(taskModuleInfo, (err, result: { youTubeVideoId?: string }) => { 
      if (err) {
        console.error(err);
        return;
      }

      console.log(`openCardDialog YouTubeSelectorCard Result: `, result);

      setYouTubeVideoId(result.youTubeVideoId);
    });

    // LA NUOVA VERSIONE NON SEMBRA FUNZIONARE 
    // const submitHandler: dialog.DialogSubmitHandler = (response) => {
    //   console.log(`Submit handler response`, response);
    //   console.log(`Submit handler - err: ${response.err}`);
    //   setYouTubeVideoId(response.result?.toString());
    // };

    // dialog.adaptiveCard.open(
    //   {
    //     title: "Card Dialog",
    //     card: require("./YouTubeSelectorCard.json"),
    //     size: {
    //       width: 350,
    //       height: 150,
    //     }
    //   }, 
    //   submitHandler
    // );
  }

  const onChangeVideo = (): void => {
    console.log(`onChangeVideo`);

    if (useCard) {
      openCardDialog();
    } else {
      openUrlDialog();
    }
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
      <Flex column gap="gap.smaller" style={CenteredWithPadding} >
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
            <fieldset className="task_modules">
              <legend>Task Modules Section</legend>

              <section className="yt_section">
                <b>Youtube Video Selection</b>

                <div className="yt_section__video_info">
                  <Text>YouTube Video ID:</Text>
                  <Input value={youTubeVideoId} disabled={formDisabled} />
                </div>

                <div className="yt_section__actions">
                  <Button
                    content="Change Video ID"
                    style={{ marginRight: "1rem" }}
                    onClick={() => onChangeVideo()}
                  />
                  <Button
                    primary
                    content="Show Video"
                    onClick={() => onShowVideo()}
                  />
                </div>
              </section>

              <section className="task_modules__config">
                <b>Configuration</b>

                <div>
                  <Checkbox
                    label="Use Card ( or html page)"
                    checked={useCard}
                    onChange={() => setUseCard((u) => !u)}
                  />
                </div>
                <div>
                  <Checkbox
                    label="Enable form"
                    checked={!formDisabled}
                    onChange={() => setFormDisabled((u) => !u)}
                  />
                </div>
              </section>
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
