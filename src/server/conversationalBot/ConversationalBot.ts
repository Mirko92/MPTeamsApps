import { BotDeclaration } from "express-msteams-host";
import * as debug from "debug";
import { CardFactory, ConversationState, MemoryStorage, UserState, TurnContext } from "botbuilder";
import { DialogBot } from "./dialogBot";
import { MainDialog } from "./dialogs/mainDialog";
import WelcomeCard from "./cards/welcomeCard";

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
  }

  public async sendWelcomeCard(context: TurnContext): Promise<void> {
    const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
    await context.sendActivity({ attachments: [welcomeCard] });
  }

}
