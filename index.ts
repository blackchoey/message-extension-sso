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
  CardFactory,
} from "botbuilder";

// This bot's main dialog.
import { TeamsBot } from "./teamsBot";
import config from "./config";
import { Application } from "@microsoft/teams-ai";
import { OnBehalfOfCredentialAuthConfig, handleMessageExtensionQueryWithSSO, OnBehalfOfUserCredential, createMicrosoftGraphClientWithCredential } from "@microsoft/teamsfx";
import "isomorphic-fetch";

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

// Create bot to handle message extension events using Teams-AI library.
const storage = new MemoryStorage();
const app = new Application({
  storage,
  adapter,
  botAppId: config.botId,
});

const oboAuthConfig: OnBehalfOfCredentialAuthConfig = {
  authorityHost: config.authorityHost,
  tenantId: config.tenantId,
  clientId: config.clientId,
  clientSecret: config.clientSecret
}
const scope= "User.Read"
const loginEndpoint = `https://${config.botDomain}/auth-start.html`

app.messageExtensions.query("searchQuery", async (context, state, query) => {
  const result = await handleMessageExtensionQueryWithSSO(context, oboAuthConfig, loginEndpoint, scope, async (token) => {
    const credential = new OnBehalfOfUserCredential(token.ssoToken, oboAuthConfig);
    const graphClient = createMicrosoftGraphClientWithCredential(credential, scope);
    const me = await graphClient.api('/me').get();
    return {
      composeExtension: {
        type:"result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(me.displayName, me.mail)]
      }
    }
  });

  if (result) {
    return result.composeExtension;
  }
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
    await app.run(context);
  });
});

server.get(
  "/auth-:name(start|end).html",
  restify.plugins.serveStatic({
    directory: path.join(__dirname, "public"),
  })
);