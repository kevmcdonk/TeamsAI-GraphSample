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
import { MainDialog } from "./dialogs/MainDialog";

export interface DataInterface {
  likeCount: number;
}

export class TeamsBot extends TeamsActivityHandler {
  // record the likeCount
  likeCountObj: { likeCount: number };
  conversationDataAccessor: StatePropertyAccessor<any>;
  graphToken: string;
  conversationState: BotState;
  userState: BotState;
  starterDialog: WaterfallDialog;
  mainDialog: MainDialog;
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
    this.mainDialog = new MainDialog(userState);
    dialogs.add(this.mainDialog);
    this.conversationDataAccessor = conversationState.createProperty(GraphTokenProperty);  

    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      /*
      const conversationData = await this.conversationDataAccessor.get(
        context, { promptedForUserName: false });
      if (this.graphToken == null) {
        if (conversationData.graphToken == null) {
          const dc = await dialogs.createContext(context);
          
          const results = await dc.continueDialog();
          if (results.status === DialogTurnStatus.empty) {
              const step = await dc.beginDialog("MainDialog");
              const token = step.result;
              if (token) {
                this.graphToken = token;
              } else {
                await context.sendActivity(`Sorry... We couldn't log you in. Try again later.`);
                return await dc.endDialog();
              }
          }
        }
      }*/

      let txt = context.activity.text;
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
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
