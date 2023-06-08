const config = {
  botId: process.env.BOT_ID,
  botPassword: process.env.BOT_PASSWORD,
  openAIKey: process.env.SECRET_OPENAI_API_KEY,
  openAIEndpoint: process.env.OPENAI_ENDPOINT,
  graphAuthorityHost: process.env.AUTHORITY_HOST,
  graphClientId: process.env.GRAPH_CLIENT_ID,
  graphTenantId: process.env.TENANT_ID,
  graphClientSecret: process.env.GRAPH_CLIENT_SECRET,
  botDomain: process.env.BOT_DOMAIN,
  connectionName: process.env.OAUTH_CONNECTION_NAME
};

export default config;
