// with thanks to https://github.com/garrytrinder/msteams-bot-sso
import { Activity, Attachment, AttachmentLayoutTypes, CardFactory, HeroCard, InputHints, MessageFactory, StatePropertyAccessor, TurnContext } from 'botbuilder';

import {
    ComponentDialog,
    ConfirmPrompt,
    DateTimePrompt,
    DialogSet,
    DialogState,
    DialogTurnResult,
    DialogTurnStatus,
    OAuthPrompt,
    PromptOptions,
    TextPrompt,
    WaterfallDialog,
    WaterfallStepContext
} from 'botbuilder-dialogs';
import fetch from 'node-fetch';
import { GraphService } from '../services/graphService';

const MAIN_WATERFALL_DIALOG = 'mainWaterfallDialog';
const CONFIRM_PROMPT = 'ConfirmPrompt';
const MAIN_DIALOG = 'MainDialog';
const OAUTH_PROMPT = 'OAuthPrompt';

export class MainDialog extends ComponentDialog {

    constructor() {
        super('MainDialog');
        console.log('Main Dialog constructor');
        // Define the main dialog and its related components.
        // This is a sample "book a flight" dialog.

        this.addDialog(new WaterfallDialog(MAIN_WATERFALL_DIALOG, [
            this.promptStep.bind(this),
            this.loginStep.bind(this),
            this.ensureOAuth.bind(this),
            this.displayToken.bind(this)
        ]));

        this.initialDialogId = MAIN_WATERFALL_DIALOG;

        this.addDialog(new OAuthPrompt(OAUTH_PROMPT, {
            connectionName: "MicrosoftGraph",//config.connectionName,
            text: 'Please Sign In',
            title: 'Sign In',
            timeout: 300000
        }));
        this.addDialog(new ConfirmPrompt(CONFIRM_PROMPT));
    }

    /**
     * The run method handles the incoming activity (in the form of a DialogContext) and passes it through the dialog system.
     * If no dialog is active, it will start the default dialog.
     * @param {TurnContext} context
     */
    public async run(context: TurnContext, accessor: StatePropertyAccessor<DialogState>) {
        console.log("Running Main Dialog");
        const dialogSet = new DialogSet(accessor);
        dialogSet.add(this);
        const dialogContext = await dialogSet.createContext(context);
        const results = await dialogContext.continueDialog();
        if (results.status === DialogTurnStatus.empty) {
            await dialogContext.beginDialog(this.id);
        }
    }

    async promptStep(stepContext) {
        try {
            return await stepContext.beginDialog(OAUTH_PROMPT);
        } catch (err) {
            console.error(err);
        }
    }

    async loginStep(stepContext) {
        // Get the token from the previous step. Note that we could also have gotten the
        // token directly from the prompt itself. There is an example of this in the next method.
        const tokenResponse = stepContext.result;
        if (!tokenResponse || !tokenResponse.token) {
            await stepContext.context.sendActivity('Login was not successful please try again.');
        } else {
            const client = new GraphService(tokenResponse.token);
            const me = await client.getUsersMail();
            await stepContext.context.sendActivity(`Here's your mail ${me}`);
            return await stepContext.prompt(CONFIRM_PROMPT, 'Would you like to view your token?');
        }
        return await stepContext.endDialog();
    }

    async ensureOAuth(stepContext) {
        await stepContext.context.sendActivity('Thank you.');

        const result = stepContext.result;
        if (result) {
            // Call the prompt again because we need the token. The reasons for this are:
            // 1. If the user is already logged in we do not need to store the token locally in the bot and worry
            // about refreshing it. We can always just call the prompt again to get the token.
            // 2. We never know how long it will take a user to respond. By the time the
            // user responds the token may have expired. The user would then be prompted to login again.
            //
            // There is no reason to store the token locally in the bot because we can always just call
            // the OAuth prompt to get the token or get a new token if needed.
            return await stepContext.beginDialog(OAUTH_PROMPT);
        }
        return await stepContext.endDialog();
    }

    async displayToken(stepContext) {
        const tokenResponse = stepContext.result;
        if (tokenResponse && tokenResponse.token) {
            await stepContext.context.sendActivity(`Here is your token ${tokenResponse.token}`);
        }
        return await stepContext.endDialog();
    }    
}