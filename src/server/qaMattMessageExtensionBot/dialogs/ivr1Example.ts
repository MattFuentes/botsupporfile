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
    TurnContext
} from "botbuilder";

const IVR_DIALOG_ID = "ivr1Dialog";
const IVR_WATERFALL_DIALOG_ID = "ivrWaterfallDialog";

export class Ivr1Example extends ComponentDialog {
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
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    private async introStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        const message = MessageFactory.text("Opcion 1.\nMicrosoft Teams es una plataforma unificada de comunicación y colaboración que combina chat persistente en el lugar de trabajo, reuniones de video, almacenamiento de archivos e integración de aplicaciones.");
        await stepContext.context.sendActivity(message);
        return await stepContext.endDialog();
    }
}
