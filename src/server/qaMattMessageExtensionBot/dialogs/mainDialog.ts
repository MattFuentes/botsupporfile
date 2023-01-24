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
    InputHints,
    TurnContext
} from "botbuilder";
import { TeamsInfoDialog } from "./teamsInfoDialog";
import { HelpDialog } from "./helpDialog";
import { MentionUserDialog } from "./mentionUserDialog";
import { IvrExample } from "./ivrExample";
import { Ivr1Example } from "./ivr1Example";
import { Ivr2Example } from "./ivr2Example";

const MAIN_DIALOG_ID = "mainDialog";
const MAIN_WATERFALL_DIALOG_ID = "mainWaterfallDialog";

export class MainDialog extends ComponentDialog {
    public onboarding: boolean;
    constructor() {
        super(MAIN_DIALOG_ID);
        this.addDialog(new TextPrompt("TextPrompt"))
            .addDialog(new TeamsInfoDialog())
            .addDialog(new HelpDialog())
            .addDialog(new IvrExample())
            .addDialog(new Ivr1Example())
            .addDialog(new Ivr2Example())
            .addDialog(new MentionUserDialog())
            .addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG_ID, [
                this.introStep.bind(this),
                this.actStep.bind(this),
                this.finalStep.bind(this)
            ]));
        this.initialDialogId = MAIN_WATERFALL_DIALOG_ID;
        this.onboarding = false;
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
        if ((stepContext.options as any).restartMsg) {
            const messageText = (stepContext.options as any).restartMsg ? (stepContext.options as any).restartMsg : "En que te puedo ayudar?";
            const promptMessage = MessageFactory.text(messageText, messageText, InputHints.ExpectingInput);
            return await stepContext.prompt("TextPrompt", { prompt: promptMessage });
        } else {
            this.onboarding = true;
            return await stepContext.next();
        }
    }

    private async actStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        if (stepContext.result) {
            /*
            ** This is where you would add LUIS to your bot, see this link for more information:
            ** https://docs.microsoft.com/en-us/azure/bot-service/bot-builder-howto-v4-luis?view=azure-bot-service-4.0&tabs=javascript
            */
            const result = stepContext.result.trim().toLocaleLowerCase();
            switch (result) {
                case "who" :
                case "who am i?": {
                    return await stepContext.beginDialog("teamsInfoDialog");
                }
                case "ayuda": {
                    return await stepContext.beginDialog("helpDialog");
                }
                case "ivr": {
                    return await stepContext.beginDialog("ivrDialog");
                }
                case "1": {
                    return await stepContext.beginDialog("ivr1Dialog");
                }
                case "2": {
                    return await stepContext.beginDialog("ivr2Dialog");
                }
                case "menciona me":
                case "mencion": {
                    return await stepContext.beginDialog("mentionUserDialog");
                }
                default: {
                    await stepContext.context.sendActivity("No entendi tu pregunta, vuelve a escribirlo por favor!");
                    return await stepContext.next();
                }
            }
        } else if (this.onboarding) {
            switch (stepContext.context.activity.text) {
                case "who": {
                    return await stepContext.beginDialog("teamsInfoDialog");
                }
                case "ayuda": {
                    return await stepContext.beginDialog("helpDialog");
                }
                case "mencion": {
                    return await stepContext.beginDialog("mentionUserDialog");
                }
                case "ivr": {
                    return await stepContext.beginDialog("ivrDialog");
                }
                default: {
                    await stepContext.context.sendActivity("No entendi tu pregunta, vuelve a escribirlo por favor!");
                    return await stepContext.next();
                }
            }
        }
        return await stepContext.next();
    }

    private async finalStep(stepContext: WaterfallStepContext): Promise<DialogTurnResult> {
        return await stepContext.replaceDialog(this.initialDialogId, { restartMsg: "En que te puedo ayudar?" });
    }
}
