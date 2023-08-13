// Import required packages
import * as restify from "restify";

import path from "path";
// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  ConfigurationBotFrameworkAuthentication,
  TurnContext,
  MemoryStorage,
  UserState,
  CardFactory,
  Attachment,
  AttachmentLayoutTypes,
  MessageFactory,
} from "botbuilder";
import {
  Application,
  ConversationHistory,
  DefaultPromptManager,
  DefaultTurnState,
  OpenAIModerator,
  AzureOpenAIPlanner,
  AI,
  DefaultConversationState,
} from "@microsoft/teams-ai";
import { OAuthPromptSettings } from "botbuilder-dialogs";
import { MailData } from "./models/mailData";
import { SiteData } from "./models/siteData";
import { SearchResultData } from "./models/searchResultData";
import { AdaptiveCards } from "@microsoft/adaptivecards-tools";

// This bot's main dialog.
import { TeamsBot } from "./teamsBot";
import config from "./config";
import {
  MessageBuilder,
  TeamsBotSsoPrompt,
  TeamsBotSsoPromptSettings,
} from "@microsoft/teamsfx";
import authConfig from "./authConfig";
import { GraphService } from "./services/graphService";
import mailCard from "./adaptiveCards/email.json";
import siteCard from "./adaptiveCards/site.json";
import searchResultCard from "./adaptiveCards/searchResult.json";

type ApplicationTurnState = DefaultTurnState<ConversationState>;
type TData = Record<string, any>;

interface ConversationState extends DefaultConversationState {
  searchResults: any;
  searchQuery: string;
  documentText: string;
  events: any;
}


// eslint-disable-next-line @typescript-eslint/no-empty-interface

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: config.botId,
  MicrosoftAppPassword: config.botPassword,
  MicrosoftAppType: "MultiTenant",
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(
  {},
  credentialsFactory
);

const adapter = new CloudAdapter(botFrameworkAuthentication);

// Catch-all for errors.
const onTurnErrorHandler = async (context: TurnContext, error: Error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a trace activity, which will be displayed in Bot Framework Emulator
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );

  // Send a message to the user
  await context.sendActivity(
    `The bot encountered unhandled error:\n ${error.message}`
  );
  await context.sendActivity(
    "To continue to run this bot, please fix the bot source code."
  );
};

// Set the onTurnError for the singleton CloudAdapter.
adapter.onTurnError = onTurnErrorHandler;

const loginUrl = process.env.INITIATE_LOGIN_ENDPOINT;
const TeamsBotSsoPromptId = "TEAMS_BOT_SSO_PROMPT";
const settings: TeamsBotSsoPromptSettings = {
  scopes: ["User.Read", "Mail.Read"],
  timeout: 900000,
  endOnInvalidMessage: true,
};

const dialog = new TeamsBotSsoPrompt(
  authConfig,
  loginUrl,
  TeamsBotSsoPromptId,
  settings
);

const memoryStorage = new MemoryStorage();
const userState = new UserState(memoryStorage);
//this.dialogState = this.conversationState.createProperty("DialogState");

console.log("OpenAIEndpoint: " + config.openAIEndpoint);
console.log("OpenAIKey: " + config.openAIKey);
const planner = new AzureOpenAIPlanner({
  apiKey: config.openAIKey,
  defaultModel: "gpt-graphsample",
  logRequests: true,
  endpoint: config.openAIEndpoint,
});

const promptManager = new DefaultPromptManager<ApplicationTurnState>(
  path.join(__dirname, "./prompts")
);
// Define storage and application
const storage = new MemoryStorage();
const app = new Application<ApplicationTurnState>({
  storage,
  ai: {
    planner,
    // moderator,
    promptManager,
    prompt: "chat",
    history: {
      assistantHistoryType: "text",
    },
  },
  authentication: {
    connectionName: config.connectionName,
    text: "Please Sign In",
    title: "Sign In",
    timeout: 300000,
  },
});

app.message("/history", async (context, state) => {
  const history = ConversationHistory.toString(state, 2000, "\n\n");
  await context.sendActivity(history);
});


app.ai.action(
  "readMail",
  async (context: TurnContext, state: ApplicationTurnState) => {
    console.log("reading mail");
    const graphService = new GraphService(state.temp.value.authToken);
    const mail = await graphService.getUsersMail();
    const mailCards = createMailCards(mail);

    await context.sendActivity({
      text: "Here's your last 10 emails:",
      attachments: mailCards,
      attachmentLayout: AttachmentLayoutTypes.Carousel,
    });
    return true;
  }
);

function createMailCards(mailResponse): Attachment[] {
  let cards = [];

  mailResponse.value.forEach(function (mail) {
    // "speak": "<s>Your  meeting about \"Adaptive Card design session\"<break strength='weak'/> is starting at ${formatDateTime(start.dateTime, 'HH:mm')}pm</s><s>Do you want to snooze <break strength='weak'/> or do you want to send a late notification to the attendees?</s>",
    let adaptiveCard = CardFactory.adaptiveCard(
      AdaptiveCards.declare(mailCard).render(mail)
    );
    cards.push(adaptiveCard);
  });
  return cards;
}

//app.ai.prompts.addFunction("listSites", async (context, state) => {
app.ai.action(
  "listSites",
  async (context: TurnContext, state: ApplicationTurnState, data: TData) => {
    const searchQuery = data.query;
    const graphService = new GraphService(state.temp.value.authToken);
    const mail = await graphService.getSites(searchQuery);
    const siteCards = createSiteCards(mail);

    await context.sendActivity({
      text: "Here's all the sites:",
      attachments: siteCards,
      attachmentLayout: AttachmentLayoutTypes.Carousel,
    });
    return true;
  }
);

function createSiteCards(siteResponse): Attachment[] {
  let cards = [];

  siteResponse.value.forEach(function (site) {
    // "speak": "<s>Your  meeting about \"Adaptive Card design session\"<break strength='weak'/> is starting at ${formatDateTime(start.dateTime, 'HH:mm')}pm</s><s>Do you want to snooze <break strength='weak'/> or do you want to send a late notification to the attendees?</s>",
    let adaptiveCard = CardFactory.adaptiveCard(
      AdaptiveCards.declare(siteCard).render(site)
    );
    cards.push(adaptiveCard);
  });
  return cards;
}


app.ai.action(
  "searchFiles",
  async (context: TurnContext, state: ApplicationTurnState, data: TData) => {
    
    /*const searchResultCards = createSearchCards(searchResults);

    await context.sendActivity({
      text: "Here's your search results:",
      attachments: searchResultCards,
      attachmentLayout: AttachmentLayoutTypes.Carousel,
    });*/
    state.conversation.value.searchQuery = data.query;
    await app.ai.chain(context, state, 'summarizeSearch');
    return true;
  }
);

app.ai.prompts.addFunction('searchFilesForSearchQuery', async (context: TurnContext, state: ApplicationTurnState) => {
  const conversation = state.conversation.value;
  const graphService = new GraphService(state.temp.value.authToken);
  const searchResults = await graphService.searchFiles(conversation.searchQuery);
  console.log(JSON.stringify(searchResults));
  return JSON.stringify(searchResults);
});

app.ai.action(
  "summariseDocument",
  async (context: TurnContext, state: ApplicationTurnState, data: TData) => {
    const filePath = data.docLink;
    const graphService = new GraphService(state.temp.value.authToken);
    let fileContents = await graphService.getFileContents('/sites/Conferences', '/general/commsverse/2023/No%20desk,%20no%20problem%20-%20empowering%20Frontline%20workers%20with%20Microsoft%20365.pptx');
    state.conversation.value.documentText = fileContents;
    await app.ai.chain(context, state, 'summarizeDocument');
    
    // End the current chain
    return false;
  }
);

app.ai.action(
  "analyzeCalendar",
  async (context: TurnContext, state: ApplicationTurnState) => {
    const graphService = new GraphService(state.temp.value.authToken);
    let events = await graphService.getNextTwoWeeksCalendars();
    state.conversation.value.events = events;
    await app.ai.chain(context, state, 'analyzeCalendar');
  
    // End the current chain
    return false;
  }
);


function createSearchCards(siteResponse): Attachment[] {
  let cards = [];

  siteResponse.value[0].hitsContainers[0].hits.forEach(function (searchResult) {
    // "speak": "<s>Your  meeting about \"Adaptive Card design session\"<break strength='weak'/> is starting at ${formatDateTime(start.dateTime, 'HH:mm')}pm</s><s>Do you want to snooze <break strength='weak'/> or do you want to send a late notification to the attendees?</s>",
    let adaptiveCard = CardFactory.adaptiveCard(
      AdaptiveCards.declare(searchResultCard).render(searchResult)
    );
    cards.push(adaptiveCard);
  });
  return cards;
}

// Register a handler to handle unknown actions that might be predicted
app.ai.action(
  AI.UnknownActionName,
  async (
    context: TurnContext,
    state: ApplicationTurnState,
    data: TData,
    action: string | undefined
  ) => {
    await context.sendActivity(
      "Not sure what to say about that. Hopefully soon we can just have a chat but not quite now."
    );
    return false;
  }
);

// List for /reset command and then delete the conversation state
app.message("/reset", async (context, state) => {
  state.conversation.delete();
  await context.sendActivity("Cleared the current conversation");
});

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
      console.log("incoming!");
      try {
        await app.run(context);
      }
      catch(e) {
        
        console.log(e);
      }
  });
});
