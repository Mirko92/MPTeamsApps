import { BotDeclaration } from "express-msteams-host";
import * as debug from "debug";
import { CardFactory, ConversationState, MemoryStorage, UserState, TurnContext, AdaptiveCardInvokeValue, AdaptiveCardInvokeResponse, StatusCodes, MessageFactory, Activity, BotFrameworkAdapter, ConversationParameters, teamsGetChannelId, TaskModuleRequest, TaskModuleResponse, TaskModuleTaskInfo } from "botbuilder";
import { DialogBot } from "./dialogBot";
import { MainDialog } from "./dialogs/mainDialog";
import WelcomeCard from "./cards/welcomeCard";

import ResponseCard from "./cards/responseCard";
import * as ACData from "adaptivecards-templating";

// Initialize debug logging module
const log = debug("msteams");
const myLog = (message: string) => log(`###################### ConversationalBot ${message} ######################`)

/**
 * Implementation for Conversational Bot
 */
@BotDeclaration(
  "/api/messages",
  new MemoryStorage(),
  process.env.MICROSOFT_APP_ID,
  process.env.MICROSOFT_APP_PASSWORD
)
export class ConversationalBot extends DialogBot {
  constructor(conversationState: ConversationState, userState: UserState) {
    super(conversationState, userState, new MainDialog());

    myLog("Constructor");

    this.onMembersAdded(async (context, next) => {
      myLog("onMembersAdded");
      const membersAdded = context.activity.membersAdded;
      if (membersAdded && membersAdded.length > 0) {
        for (let cnt = 0; cnt < membersAdded.length; cnt++) {
          if (membersAdded[cnt].id !== context.activity.recipient.id) {
            await this.sendWelcomeCard(context);
          }
        }
      }
      await next();
    });

    this.onMessageReaction(this._onMessageReaction);
  }

  private _onMessageReaction = async (context: TurnContext, next: () => Promise<void>) => {
    try {
      if (context.activity.reactionsAdded) {
        context.activity.reactionsAdded.forEach(async (reaction) => {
          if (reaction.type === "like") {
            await context.sendActivity("Thank you!");
          }
        });
      }
      await next();
    } catch (error) {
      log("onMessageReaction: error\n", error);
    }
  }

  public async sendWelcomeCard(context: TurnContext): Promise<void> {
    const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
    await context.sendActivity({ attachments: [welcomeCard] });
  }

  protected async onAdaptiveCardInvoke(context: TurnContext, invokeValue: AdaptiveCardInvokeValue): Promise<any> {
    let cardResponse: AdaptiveCardInvokeResponse;

    try {
      const verb = invokeValue.action.verb;
      log(`onAdaptiveCardInvoke verb: '${verb}'`);
      switch (verb) {
        case "update":
          {
            let clickCount: number = invokeValue.action.data.count as number;
            const cardData = {
              message: `Updated count: ${++clickCount}`,
              count: clickCount,
              showDelete: true
            };
            const template = new ACData.Template(ResponseCard);
            const context: ACData.IEvaluationContext = {
              $root: cardData
            };
            const acCard = template.expand(context);

            cardResponse = {
              statusCode: StatusCodes.OK,
              type: "application/vnd.microsoft.card.adaptive",
              value: acCard
            } as unknown as AdaptiveCardInvokeResponse;

          }
          break;

        case "newconversation":
          {
            const message = MessageFactory.text("This will be the first message in a new thread");
            await this.teamsCreateConversation(context, message);
            return Promise.resolve({
              statusCode: 200,
              type: "application/vnd.microsoft.activity.message",
              value: "Thread created"
            });
          }

        case "delete":
          await context.deleteActivity(context!.activity!.replyToId!);
          return Promise.resolve({
            statusCode: 200,
            type: "application/vnd.microsoft.activity.message",
            value: "Deleting activity..."
          });

        default:
          return Promise.resolve({
            statusCode: 200,
            type: "application/vnd.microsoft.activity.message",
            value: "I don't know how to process that verb"
          });
      }
      return Promise.resolve(cardResponse);
    } catch (error) {
      return Promise.reject(error);
    }
  }

  private async teamsCreateConversation(context: TurnContext, message: Partial<Activity>): Promise<void> {
    // get a reference to the bot adapter & create a connection to the Teams API
    const adapter = <BotFrameworkAdapter>context.adapter;
    const connectorClient = adapter.createConnectorClient(context.activity.serviceUrl);

    // set current teams channel in new conversation parameters
    const teamsChannelId = teamsGetChannelId(context.activity);
    const conversationParameters: ConversationParameters = {
      isGroup: true,
      channelData: {
        channel: {
          id: teamsChannelId
        }
      },
      activity: message as Activity,
      bot: context.activity.recipient
    };

    // create conversation and send message
    await connectorClient.conversations.createConversation(conversationParameters);
  }

  protected handleTeamsTaskModuleFetch(context: TurnContext, request: TaskModuleRequest): Promise<TaskModuleResponse> {
    myLog(`handleTeamsTaskModuleFetch`);
    let response: TaskModuleResponse;

    switch (request.data.taskModule) {
      case "selector":
        myLog(`handleTeamsTaskModuleFetch selector`);
        response = ({
          task: {
            type: "continue",
            value: {
              title: "YouTube Video Selector",
              card: this.getSelectorAdaptiveCard(request.data.videoId),
              width: 350,
              height: 250
            } as TaskModuleTaskInfo
          }
        } as TaskModuleResponse);
        break;
      case "player":
        myLog(`handleTeamsTaskModuleFetch player`);
        response = ({
          task: {
            type: "continue",
            value: {
              title: "YouTube Player",
              url: `https://${process.env.PUBLIC_HOSTNAME}/MS600TAB_PERSONAL/player.html?vid=${request.data.videoId}`,
              width: 1000,
              height: 700
            } as TaskModuleTaskInfo
          }
        } as TaskModuleResponse);
        break;
      default:
        myLog(`handleTeamsTaskModuleFetch default`);
        response = ({
          task: {
            type: "continue",
            value: {
              title: "YouTube Player",
              url: `https://${process.env.PUBLIC_HOSTNAME}/MS600TAB_PERSONAL/player.html?vid=NRY-9Eel2n0&default=1`,
              width: 1000,
              height: 700
            } as TaskModuleTaskInfo
          }
        } as TaskModuleResponse);
        break;
    };

    console.log("handleTeamsTaskModuleFetch() response", response);
    return Promise.resolve(response);
  }

  protected handleTeamsTaskModuleSubmit(context: TurnContext, request: TaskModuleRequest): Promise<TaskModuleResponse> {
    const response: TaskModuleResponse = {
      task: {
        type: "continue",
        value: {
          title: "YouTube Player",
          url: `https://${process.env.PUBLIC_HOSTNAME}/MS600TAB_PERSONAL/player.html?vid=${request.data.youTubeVideoId}`,
          width: 1000,
          height: 700
        } as TaskModuleTaskInfo
      }
    } as TaskModuleResponse;
    return Promise.resolve(response);
  }

  private getSelectorAdaptiveCard(defaultVideoId: string = "") {
    return CardFactory.adaptiveCard({
      type: "AdaptiveCard",
      version: "1.0",
      body: [
        {
          type: "Container",
          items: [
            {
              type: "TextBlock",
              text: "YouTube Video Selector",
              weight: "bolder",
              size: "extraLarge"
            }
          ]
        },
        {
          type: "Container",
          items: [
            {
              type: "TextBlock",
              text: "Enter the ID of a YouTube video to show in the task module player.",
              wrap: true
            },
            {
              type: "Input.Text",
              id: "youTubeVideoId",
              value: defaultVideoId
            }
          ]
        }
      ],
      actions: [
        {
          type: "Action.Submit",
          title: "Update"
        }
      ]
    });
  }
}
