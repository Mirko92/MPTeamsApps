import {
  ConversationState,
  UserState,
  TeamsActivityHandler,
  TurnContext,
  CardFactory,
  ChannelInfo,
  TeamInfo,
  MessageFactory
} from "botbuilder";
import { MainDialog } from "./dialogs/mainDialog";

const log = (msg: string) => console.log(`###################### DialogBot ${msg} ######################`)
export class DialogBot extends TeamsActivityHandler {
  public dialogState: any;

  constructor(public conversationState: ConversationState, public userState: UserState, public dialog: MainDialog) {
    super();
    log(`constructor`);

    this.conversationState = conversationState;
    this.userState = userState;
    this.dialog = dialog;
    this.dialogState = this.conversationState.createProperty("DialogState");

    this.onMessage(this._onMessage);

    this.onTeamsChannelCreatedEvent(this._onTeamsChannelCreatedEvent);

    this.onReactionsAdded( this._onReactionAdded );

    this.onReactionsRemoved( this._onReactionRemoved )
  }

  public async run(context: TurnContext) {
    await super.run(context);
    // Save any state changes. The load happened during the execution of the Dialog.
    await this.conversationState.saveChanges(context, false);
    await this.userState.saveChanges(context, false);
  }

  private _onReactionAdded = async () => {
    log(`_onReactionAdded`);

  }

  private _onReactionRemoved = async () => {
    log(`_onReactionRemoved`);

  }

  private _onMessage = async (context: TurnContext, next: () => Promise<void>) => {
    log(`_onMessage`);

    // Run the MainDialog with the new message Activity.
    await this.dialog.run(context, this.dialogState);
    await next();
  }

  private _onTeamsChannelCreatedEvent = async (channelInfo: ChannelInfo, teamInfo: TeamInfo, turnContext: TurnContext, next: () => Promise<void>): Promise<void> => {
    log(`_onTeamsChannelCreatedEvent`);
    
    const card = CardFactory.adaptiveCard({});
    const message = MessageFactory.attachment(card);
    await turnContext.sendActivity(message);
    await next()
  }
}
