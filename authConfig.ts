import { AppCredentialAuthConfig, OnBehalfOfCredentialAuthConfig } from "@microsoft/teamsfx";
import config from "./config";

const authConfig: OnBehalfOfCredentialAuthConfig | AppCredentialAuthConfig = {
  authorityHost: process.env.AUTHORTY_HOST,
  clientId: process.env.CLIENT_ID,
  tenantId: process.env.TENANT_ID,
  clientSecret: process.env.CLIENT_SECRET
};

export default authConfig;