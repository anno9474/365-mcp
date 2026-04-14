import { ConfidentialClientApplication, type Configuration } from "@azure/msal-node";
import { Client } from "@microsoft/microsoft-graph-client";

const SCOPES = [
  "https://graph.microsoft.com/Mail.Read",
  "https://graph.microsoft.com/Calendars.Read",
];

let msalApp: ConfidentialClientApplication | null = null;

function getMsalApp(): ConfidentialClientApplication {
  if (msalApp) return msalApp;

  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  const tenantId = process.env.AZURE_TENANT_ID;

  if (!clientId || !clientSecret || !tenantId) {
    throw new Error(
      "Missing required environment variables: AZURE_CLIENT_ID, AZURE_CLIENT_SECRET, and AZURE_TENANT_ID must be set."
    );
  }

  const config: Configuration = {
    auth: {
      clientId,
      clientSecret,
      authority: `https://login.microsoftonline.com/${tenantId}`,
    },
  };

  msalApp = new ConfidentialClientApplication(config);
  return msalApp;
}

/**
 * Builds a Graph API client that exchanges the caller's Azure AD access token
 * for a Graph-scoped token via the On-Behalf-Of flow.
 *
 * Prerequisites in Azure AD:
 * - App must be a confidential client (client secret set)
 * - App must Expose an API with a delegated scope (e.g. access_as_user)
 * - LibreChat OIDC must request that scope so the forwarded token has the
 *   correct audience for OBO assertion
 */
export function buildGraphClient(userAccessToken: string): Client {
  const app = getMsalApp();

  return Client.init({
    authProvider: async (done) => {
      try {
        const result = await app.acquireTokenOnBehalfOf({
          oboAssertion: userAccessToken,
          scopes: SCOPES,
        });

        if (!result?.accessToken) {
          done(new Error("OBO token exchange returned no access token"), null);
          return;
        }

        done(null, result.accessToken);
      } catch (error) {
        done(error as Error, null);
      }
    },
  });
}
