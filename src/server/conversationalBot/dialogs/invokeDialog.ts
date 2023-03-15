import {
  ComponentDialog,
  DialogSet,
  DialogState,
  DialogTurnResult,
  DialogTurnStatus,
  TextPrompt,
  WaterfallDialog,
  WaterfallStepContext
} from "botbuilder-dialogs";
import {
  CardFactory,
  MessageFactory,
  StatePropertyAccessor,
  TurnContext
} from "botbuilder";

const INVOKE_DIALOG_ID = "invokeDialog";
const INVOKE_WATERFALL_DIALOG_ID = "invokeWaterfallDialog";

export class InvokeDialog extends ComponentDialog {
  constructor() {
    super(INVOKE_DIALOG_ID);
    this.addDialog(new TextPrompt("TextPrompt"))
      .addDialog(new WaterfallDialog(INVOKE_WATERFALL_DIALOG_ID, [
        this.introStep.bind(this)
      ]));
    this.initialDialogId = INVOKE_WATERFALL_DIALOG_ID;
  }

  public async run(context: TurnContext, accessor: StatePropertyAccessor<DialogState>) {
    const dialogSet = new DialogSet(accessor);
    dialogSet.add(this);
    const dialogContext = await dialogSet.createContext(context);
    const results = await dialogContext.continueDialog();
    if (results.status === DialogTurnStatus.empty) {
      await dialogContext.beginDialog(this.id);
    }
  }

  private async introStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
    const card = CardFactory.heroCard("Learn Microsoft Teams", undefined, [
      ...[
        { videoId: "mmw57bp8AGI", taskModule: "player",     title: "Festa degli uomini" },
        { videoId: "AQcdZYkFPCY", taskModule: "player",     title: "Watch 'Microsoft Teams embedded web experiences'" },
        { videoId: "aHoRK8cr6Og", taskModule: "player",     title: "Watch 'Task-oriented interactions in Microsoft Teams with messaging extensions'" },
        { videoId: "aHoRK8cr6Og", taskModule: "something",  title: "Watch a invalid action..." },
        { videoId: "QHPBw7F4OL4", taskModule: "selector",   title: "Watch Specific Video" },
      ].map((x) => ({
        type: "invoke",
        title: x.title,
        value: { type: "task/fetch", taskModule: x.taskModule, videoId: x.videoId }
      }))
    ]);

    const message = MessageFactory.attachment(card);
    await stepContext.context.sendActivity(message);
    return await stepContext.endDialog();
  }

  
}
