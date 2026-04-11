import { DeviceCodeCredential } from "@azure/identity";
import { useIdentityPlugin } from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

// Enable persistent token cache so device code auth survives container restarts
useIdentityPlugin(cachePersistencePlugin);

let graphClient: Client | null = null;

function createGraphClient(): Client {
  const clientId = process.env.AZURE_CLIENT_ID;
  const tenantId = process.env.AZURE_TENANT_ID;

  if (!clientId || !tenantId) {
    throw new Error(
      "Missing required environment variables: AZURE_CLIENT_ID and AZURE_TENANT_ID must be set."
    );
  }

  const credential = new DeviceCodeCredential({
    clientId,
    tenantId,
    userPromptCallback: (info) => {
      console.log("========================================");
      console.log("DEVICE CODE AUTHENTICATION REQUIRED");
      console.log("========================================");
      console.log(info.message);
      console.log("========================================");
    },
    tokenCachePersistenceOptions: {
      enabled: true,
      name: "365-mcp-token-cache",
    },
  });

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ["https://graph.microsoft.com/.default"],
  });

  return Client.initWithMiddleware({ authProvider });
}

export function getGraphClient(): Client {
  if (!graphClient) {
    graphClient = createGraphClient();
  }
  return graphClient;
}

export async function testConnection(): Promise<void> {
  const client = getGraphClient();
  const me = await client.api("/me").get();
  console.log(`Authenticated as: ${me.displayName} (${me.userPrincipalName})`);
}
