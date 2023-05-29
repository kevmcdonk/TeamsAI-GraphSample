import { AppCredentialAuthConfig, OnBehalfOfCredentialAuthConfig } from "@microsoft/teamsfx";
import config from "./config";

const authConfig: OnBehalfOfCredentialAuthConfig = {
  authorityHost: config.graphAuthorityHost,
  clientId: config.graphClientId,
  tenantId: config.graphTenantId,
  clientSecret: config.graphClientSecret
};

export default authConfig;