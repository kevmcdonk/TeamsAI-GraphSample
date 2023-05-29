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
  ConversationState,
  UserState,
} from "botbuilder";

// This bot's main dialog.
import { TeamsBot } from "./teamsBot";
import config from "./config";
import { TeamsBotSsoPrompt, TeamsBotSsoPromptSettings } from "@microsoft/teamsfx";
import authConfig from "./authConfig";

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
  await context.sendActivity(`The bot encountered unhandled error:\n ${error.message}`);
  await context.sendActivity("To continue to run this bot, please fix the bot source code.");
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
const convoState = new ConversationState(memoryStorage);  
const userState = new UserState(memoryStorage);
//this.dialogState = this.conversationState.createProperty("DialogState");


// Create the bot that will handle incoming messages.
const bot = new TeamsBot(convoState, userState, dialog);

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.bodyParser());
server.listen(process.env.port || process.env.PORT || 3978, () => {
  console.log(`\nBot Started, ${server.name} listening to ${server.url}`);
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  await adapter.process(req, res, async (context) => {
    await bot.run(context);
  });
});


