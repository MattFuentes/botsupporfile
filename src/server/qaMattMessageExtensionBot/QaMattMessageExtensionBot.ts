import { BotDeclaration, PreventIframe, MessageExtensionDeclaration } from "express-msteams-host";
import * as debug from "debug";
import { TeamsActivityHandler, StatePropertyAccessor, ActivityTypes, CardFactory, ConversationState, MemoryStorage, UserState, TurnContext } from "botbuilder";
import { DialogBot } from "./dialogBot";
import { MainDialog } from "./dialogs/mainDialog";
import WelcomeCard from "./cards/welcomeCard";
import QaMattMessageExtension from "../qaMattMessageExtension/QaMattMessageExtension";
import { DialogSet, DialogState } from "botbuilder-dialogs";
// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for qa-matt Message Extension Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_ID,
    // eslint-disable-next-line no-undef
    process.env.MICROSOFT_APP_PASSWORD)
@PreventIframe("/qaMattMessageExtensionBot/aboutQaMatt.html")
export class QaMattMessageExtensionBot extends DialogBot {
    constructor(conversationState: ConversationState, userState: UserState) {
        super(conversationState, userState, new MainDialog());
        // Message extension QaMattMessageExtension
        this._qaMattMessageExtension = new QaMattMessageExtension();

        this.onMembersAdded(async (context, next) => {
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

    /** Local property for QaMattMessageExtension */
    @MessageExtensionDeclaration("qaMattMessageExtension")
    private _qaMattMessageExtension: QaMattMessageExtension;

    public async sendWelcomeCard(context: TurnContext): Promise<void> {
        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
        await context.sendActivity({ attachments: [welcomeCard] });
    }

}
