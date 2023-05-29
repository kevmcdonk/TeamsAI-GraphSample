import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  AdaptiveCardInvokeValue,
  AdaptiveCardInvokeResponse,
  MemoryStorage,
  ConversationState,
  StatePropertyAccessor,
  BotState,
  UserState,
} from "botbuilder";
import { Dialog, DialogState, DialogTurnStatus, DialogSet, WaterfallDialog } from "botbuilder-dialogs";
import { 
  Application, 
  ConversationHistory, 
  DefaultPromptManager, 
  DefaultTurnState, 
  OpenAIModerator, 
  AzureOpenAIPlanner, AI
} from '@microsoft/teams-ai';
import path from "path";
import config from "./config";

import rawWelcomeCard from "./adaptiveCards/welcome.json";
import rawLearnCard from "./adaptiveCards/learn.json";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import {
  TeamsBotSsoPrompt,
  TeamsBotSsoPromptSettings,
  TeamsFx,
} from "@microsoft/teamsfx";
import authConfig from "./authConfig";
import { GraphService } from "./services/graphService";
import { start } from "repl";

export interface DataInterface {
  likeCount: number;
}

type ApplicationTurnState = DefaultTurnState<ConversationState>;

// Create AI components
const planner = new AzureOpenAIPlanner({
  apiKey: config.openAIKey,
  defaultModel: 'GPT35Completions',
  logRequests: true,
  endpoint: 'https://openai-woeb2.openai.azure.com/'
});

const promptManager = new DefaultPromptManager<ApplicationTurnState>(path.join(__dirname, './prompts'));
// Define storage and application
const storage = new MemoryStorage();
const app = new Application<ApplicationTurnState>({
  storage,
  ai: {
      planner,
      // moderator,
      promptManager,
      prompt: 'chat',
      history: {
          assistantHistoryType: 'text'
      }
  }
});
const graphService = new GraphService();

app.message('/history', async (context, state) => {
  const history = ConversationHistory.toString(state, 2000, '\n\n');
  await context.sendActivity(history);
});

app.ai.prompts.addFunction('readMail', async (context, state) => {
  //var id = createNewWorkItem(state, data);
  // Note that the bot doesn't run so all that bot stuff here is pointless...
  // So how do I get dialogs...
  await context.sendActivity(`Here's your email`);
  return graphService.getUsersMail();
});


export class TeamsBot extends TeamsActivityHandler {
  // record the likeCount
  likeCountObj: { likeCount: number };
  conversationDataAccessor: StatePropertyAccessor<any>;
  graphToken: string;
  conversationState: BotState;
  userState: BotState;
  starterDialog: WaterfallDialog;
  dialogState: any;

  constructor(conversationState: BotState, userState: BotState, dialog: Dialog) {
    super();

    this.likeCountObj = { likeCount: 0 };

    const dialogState = conversationState.createProperty("dialogState");
    const dialogs = new DialogSet(dialogState);

    const GraphTokenProperty = "GRAPH_TOKEN";

    dialogs.add(dialog);

    this.starterDialog = new WaterfallDialog("taskNeedingLogin", [
      async (step) => {
        return await step.beginDialog("TEAMS_BOT_SSO_PROMPT");
      },
      async (step) => {
       const token = step.result;
       if (token) {
         this.graphToken = token;
       } else {
        await step.context.sendActivity(`Sorry... We couldn't log you in. Try again later.`);
        return await step.endDialog();
       }
     },
    ]);
    dialogs.add(this.starterDialog);  
    this.conversationDataAccessor = conversationState.createProperty(GraphTokenProperty);  

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      const conversationData = await this.conversationDataAccessor.get(
        context, { promptedForUserName: false });
      if (this.graphToken == null) {
        if (conversationData.graphToken == null) {
          const dc = await dialogs.createContext(context);
          
          const results = await dc.continueDialog();
          if (results.status === DialogTurnStatus.empty) {
              const step = await dc.beginDialog("taskNeedingLogin");
              const token = step.result;
              if (token) {
                this.graphToken = token;
              } else {
                await context.sendActivity(`Sorry... We couldn't log you in. Try again later.`);
                return await dc.endDialog();
              }
          }
        }
      }

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card =
            AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
        case "learn": {
          this.likeCountObj.likeCount = 0;
          const card = AdaptiveCards.declare<DataInterface>(
            rawLearnCard
          ).render(this.likeCountObj);
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    this.onMembersAdded(async (context, next) => {
      const membersAdded = context.activity.membersAdded;
      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id) {
          const card =
            AdaptiveCards.declareWithoutData(rawWelcomeCard).render();
          await context.sendActivity({
            attachments: [CardFactory.adaptiveCard(card)],
          });
          break;
        }
      }
      await next();
    });
  }

  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(
    context: TurnContext,
    invokeValue: AdaptiveCardInvokeValue
  ): Promise<AdaptiveCardInvokeResponse> {
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const card = AdaptiveCards.declare<DataInterface>(rawLearnCard).render(
        this.likeCountObj
      );
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200, type: undefined, value: undefined };
    }
  }
}
