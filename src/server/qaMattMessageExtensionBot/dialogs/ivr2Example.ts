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
    MessageFactory,
    StatePropertyAccessor,
    TurnContext,
    CardFactory
} from "botbuilder";
import WelcomeCard from "../cards/welcomeCard";

const IVR_DIALOG_ID = "ivr2Dialog";
const IVR_WATERFALL_DIALOG_ID = "ivrWaterfallDialog";

export class Ivr2Example extends ComponentDialog {
    constructor() {
        super(IVR_DIALOG_ID);
        this.addDialog(new TextPrompt("TextPrompt"))
            .addDialog(new WaterfallDialog(IVR_WATERFALL_DIALOG_ID, [
                this.introStep.bind(this)
            ]));
        this.initialDialogId = IVR_WATERFALL_DIALOG_ID;
    }

    public async run(context: TurnContext, accessor: StatePropertyAccessor<DialogState>) {
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);
        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        await this.sendWelcomeCard(context);
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }
    public async sendWelcomeCard(context: TurnContext): Promise<void> {
        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
        await context.sendActivity({ attachments: [welcomeCard] });
    }

    private async introStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const message = MessageFactory.text("Opcion 2.\nGrupo Desarrollo parte de Newtech");
        await stepContext.context.sendActivity(message);
        return await stepContext.endDialog();
    }
}
