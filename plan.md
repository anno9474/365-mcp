# 365-MCP Server — Implementation Plan

## Overview

A secure, lightweight MCP (Model Context Protocol) server that connects LibreChat with Microsoft 365 via the Graph API. Read-only access to emails and calendar events using delegated user permissions and device code authentication flow.

**Transport:** Streamable HTTP (the server runs as a Docker container alongside LibreChat)
**Auth:** OAuth 2.0 Device Code Flow (headless-friendly, no redirect URI needed)
**Scope:** Read-only — `Mail.Read`, `Calendars.Read`

## Azure AD App Registration (already done)

- **Client ID:** `d8bdcfbc-843f-4228-af94-933ef5965091`
- **Tenant ID:** `f35bff13-d3f1-49c0-9e22-84bba41d31a2`
- **Type:** Public client (Allow public client flows = Yes)
- **Delegated permissions:** `Mail.Read`, `Calendars.Read`
- **No client secret** (device code flow is a public client flow)

---

## Project Structure

```
365-mcp/
├── src/
│   ├── index.ts              # Entry point — Express server + MCP setup
│   ├── auth/
│   │   └── graph-client.ts   # Device code auth + Graph client factory
│   └── tools/
│       ├── mail.ts           # Mail tools (list_emails, read_email)
│       └── calendar.ts       # Calendar tools (list_events, read_event)
├── package.json
├── tsconfig.json
├── Dockerfile
├── docker-compose.yml        # Standalone dev compose (can be merged into LibreChat's)
├── .env.example
├── .env                      # (gitignored) actual config
├── .gitignore
└── plan.md                   # This file
```

---

## Phase 1 — Project Scaffold

### Task 1.1: Initialize package.json

Create `package.json` with the following exact content:

```json
{
  "name": "365-mcp",
  "version": "0.1.0",
  "private": true,
  "type": "module",
  "scripts": {
    "build": "tsc",
    "start": "node dist/index.js",
    "dev": "tsx watch src/index.ts"
  },
  "dependencies": {
    "@azure/identity": "^4.6.0",
    "@azure/identity-cache-persistence": "^1.3.0",
    "@microsoft/microsoft-graph-client": "^3.0.7",
    "@modelcontextprotocol/server": "^1.12.0",
    "@modelcontextprotocol/node": "^1.12.0",
    "@modelcontextprotocol/express": "^0.1.0",
    "cors": "^2.8.5",
    "express": "^4.21.0",
    "zod": "^3.25.0"
  },
  "devDependencies": {
    "@types/cors": "^2.8.17",
    "@types/express": "^5.0.0",
    "@types/node": "^22.0.0",
    "tsx": "^4.19.0",
    "typescript": "^5.7.0"
  }
}
```

### Task 1.2: Create tsconfig.json

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "Node16",
    "moduleResolution": "Node16",
    "outDir": "dist",
    "rootDir": "src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true,
    "resolveJsonModule": true,
    "declaration": true,
    "declarationMap": true,
    "sourceMap": true
  },
  "include": ["src/**/*"],
  "exclude": ["node_modules", "dist"]
}
```

### Task 1.3: Create .gitignore

```
node_modules/
dist/
.env
token-cache/
```

### Task 1.4: Create .env.example

```env
# Azure AD App Registration
AZURE_CLIENT_ID=d8bdcfbc-843f-4228-af94-933ef5965091
AZURE_TENANT_ID=f35bff13-d3f1-49c0-9e22-84bba41d31a2

# MCP Server
MCP_PORT=3000

# Token cache directory (must be a mounted volume in Docker)
TOKEN_CACHE_DIR=/data/token-cache
```

### Task 1.5: Create .env (actual values)

Same content as `.env.example` but with the real values filled in (they are the same in this case since there are no secrets).

### Task 1.6: Run npm install

```bash
cd /home/bwa/git/365-mcp && npm install
```

---

## Phase 2 — Authentication

### Task 2.1: Create `src/auth/graph-client.ts`

This file handles device code authentication and creates an authenticated Microsoft Graph client.

**File:** `src/auth/graph-client.ts`

**Imports needed:**

```typescript
import { DeviceCodeCredential } from "@azure/identity";
import { useIdentityPlugin } from "@azure/identity";
import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";
```

**Implementation requirements:**

1. **Enable cache persistence plugin** — call `useIdentityPlugin(cachePersistencePlugin)` once at module load.

2. **Create and export a function `createGraphClient()`** that:
   - Reads `AZURE_CLIENT_ID` and `AZURE_TENANT_ID` from `process.env`. If either is missing, throw an error with a clear message.
   - Creates a `DeviceCodeCredential` with these options:
     ```typescript
     const credential = new DeviceCodeCredential({
       clientId: process.env.AZURE_CLIENT_ID,
       tenantId: process.env.AZURE_TENANT_ID,
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
     ```
   - Creates a `TokenCredentialAuthenticationProvider`:
     ```typescript
     const authProvider = new TokenCredentialAuthenticationProvider(credential, {
       scopes: ["https://graph.microsoft.com/.default"],
     });
     ```
   - Creates and returns the Graph `Client`:
     ```typescript
     const client = Client.initWithMiddleware({
       authProvider,
     });
     return client;
     ```

3. **Export a singleton pattern** — the Graph client should only be created once. Use a module-level variable:
   ```typescript
   let graphClient: Client | null = null;

   export function getGraphClient(): Client {
     if (!graphClient) {
       graphClient = createGraphClient();
     }
     return graphClient;
   }
   ```

4. **Export a `testConnection()` async function** that calls `client.api("/me").get()` and logs the user's display name. This is used to verify auth works and trigger the initial device code flow.

**Complete file behavior:**
- On first call to `getGraphClient()` + any Graph API request, the device code prompt appears in the container logs.
- The user opens the URL, enters the code, authenticates.
- The token (including refresh token) is cached by `@azure/identity-cache-persistence` using the OS credential store or encrypted file.
- On subsequent container restarts, the cached refresh token is used automatically — no re-auth needed unless the refresh token expires (typically 90 days).

---

## Phase 3 — MCP Server + Tools

### Task 3.1: Create `src/tools/mail.ts`

**File:** `src/tools/mail.ts`

This file exports a function that registers email tools on the MCP server.

**Function signature:**
```typescript
import type { McpServer } from "@modelcontextprotocol/server";
import type { Client } from "@microsoft/microsoft-graph-client";

export function registerMailTools(server: McpServer, graphClient: Client): void
```

**Tool 1: `list_emails`**

- **Name:** `list_emails`
- **Description:** `List recent emails from your Microsoft 365 mailbox. Returns subject, sender, date, and a preview of each email.`
- **Input schema (zod):**
  ```typescript
  z.object({
    folder: z.string().optional().describe("Mail folder to list from. Defaults to 'inbox'. Other options: 'sentitems', 'drafts', 'deleteditems', 'archive'."),
    top: z.number().min(1).max(50).optional().describe("Number of emails to return. Defaults to 10, max 50."),
    filter: z.string().optional().describe("OData filter expression. Example: \"isRead eq false\" for unread emails, or \"from/emailAddress/address eq 'user@example.com'\""),
    search: z.string().optional().describe("Search query string. Searches subject, body, and sender. Example: \"quarterly report\""),
  })
  ```
- **Implementation:**
  ```typescript
  // Build the API path
  const folder = args.folder || "inbox";
  const top = args.top || 10;

  let apiPath = `/me/mailFolders/${folder}/messages`;
  let request = graphClient.api(apiPath)
    .top(top)
    .select("id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments")
    .orderby("receivedDateTime desc");

  if (args.filter) {
    request = request.filter(args.filter);
  }
  if (args.search) {
    request = request.search(`"${args.search}"`);
  }

  const response = await request.get();
  ```
- **Return format:** Return a single `text` content block with a formatted summary:
  ```
  Found {count} emails in {folder}:

  1. **{subject}**
     From: {sender name} <{sender email}>
     Date: {receivedDateTime formatted as ISO string}
     Read: {yes/no} | Attachments: {yes/no}
     Preview: {bodyPreview, first 150 chars}
     ID: {id}

  2. ...
  ```
- **Error handling:** Wrap the Graph API call in try/catch. On error, return `isError: true` with the error message as text content.

**Tool 2: `read_email`**

- **Name:** `read_email`
- **Description:** `Read the full content of a specific email by its ID. Returns the complete email body, headers, and attachment list.`
- **Input schema (zod):**
  ```typescript
  z.object({
    messageId: z.string().describe("The ID of the email message to read. Get this from list_emails."),
  })
  ```
- **Implementation:**
  ```typescript
  const message = await graphClient.api(`/me/messages/${args.messageId}`)
    .select("id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,attachments")
    .expand("attachments($select=id,name,contentType,size)")
    .get();
  ```
- **Return format:**
  ```
  Subject: {subject}
  From: {from name} <{from email}>
  To: {comma-separated list of to recipients}
  CC: {comma-separated list of cc recipients, or "none"}
  Date: {receivedDateTime}

  --- Body ---
  {body.content — if body.contentType is "html", strip tags or return as-is; prefer returning the text}

  --- Attachments ---
  {list of attachment names with sizes, or "No attachments"}
  ```
- **Error handling:** Same as above — try/catch, return `isError: true` on failure.

### Task 3.2: Create `src/tools/calendar.ts`

**File:** `src/tools/calendar.ts`

**Function signature:**
```typescript
import type { McpServer } from "@modelcontextprotocol/server";
import type { Client } from "@microsoft/microsoft-graph-client";

export function registerCalendarTools(server: McpServer, graphClient: Client): void
```

**Tool 1: `list_events`**

- **Name:** `list_events`
- **Description:** `List calendar events from your Microsoft 365 calendar within a specified date range.`
- **Input schema (zod):**
  ```typescript
  z.object({
    startDateTime: z.string().describe("Start of the time range in ISO 8601 format. Example: '2026-04-11T00:00:00'. Defaults to now."),
    endDateTime: z.string().describe("End of the time range in ISO 8601 format. Example: '2026-04-18T23:59:59'. Defaults to 7 days from now."),
    top: z.number().min(1).max(50).optional().describe("Number of events to return. Defaults to 20, max 50."),
  })
  ```
  Note: `startDateTime` and `endDateTime` should both be optional with defaults computed at runtime.
- **Implementation:**
  ```typescript
  const now = new Date();
  const startDateTime = args.startDateTime || now.toISOString();
  const endDateTime = args.endDateTime || new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString();
  const top = args.top || 20;

  const response = await graphClient.api("/me/calendarView")
    .query({
      startDateTime,
      endDateTime,
    })
    .top(top)
    .select("id,subject,start,end,location,organizer,isAllDay,isCancelled,bodyPreview")
    .orderby("start/dateTime asc")
    .get();
  ```
- **Return format:**
  ```
  Found {count} events from {startDateTime} to {endDateTime}:

  1. **{subject}**
     When: {start.dateTime} — {end.dateTime} ({timezone})
     All day: {yes/no}
     Location: {location.displayName or "No location"}
     Organizer: {organizer name} <{organizer email}>
     Status: {cancelled/confirmed}
     Preview: {bodyPreview, first 150 chars}
     ID: {id}

  2. ...
  ```
- **Error handling:** try/catch, `isError: true` on failure.

**Tool 2: `read_event`**

- **Name:** `read_event`
- **Description:** `Read the full details of a specific calendar event by its ID.`
- **Input schema (zod):**
  ```typescript
  z.object({
    eventId: z.string().describe("The ID of the calendar event to read. Get this from list_events."),
  })
  ```
- **Implementation:**
  ```typescript
  const event = await graphClient.api(`/me/events/${args.eventId}`)
    .select("id,subject,body,start,end,location,organizer,attendees,isAllDay,isCancelled,recurrence,onlineMeeting")
    .get();
  ```
- **Return format:**
  ```
  Subject: {subject}
  When: {start.dateTime} — {end.dateTime} ({timezone})
  All day: {yes/no}
  Location: {location.displayName or "No location"}
  Organizer: {organizer name} <{organizer email}>
  Status: {cancelled/confirmed}

  --- Attendees ---
  - {name} <{email}> — {response status: accepted/declined/tentative/none}
  ...

  --- Body ---
  {body.content}

  --- Online Meeting ---
  {join URL if available, or "No online meeting link"}
  ```
- **Error handling:** try/catch, `isError: true` on failure.

### Task 3.3: Create `src/index.ts`

**File:** `src/index.ts`

This is the entry point. It wires together the MCP server, Express HTTP transport, auth, and tools.

**Imports:**
```typescript
import { randomUUID } from "node:crypto";
import { McpServer, isInitializeRequest } from "@modelcontextprotocol/server";
import { NodeStreamableHTTPServerTransport } from "@modelcontextprotocol/node";
import cors from "cors";
import express from "express";
import type { Request, Response } from "express";
import { getGraphClient, testConnection } from "./auth/graph-client.js";
import { registerMailTools } from "./tools/mail.js";
import { registerCalendarTools } from "./tools/calendar.js";
```

**Implementation — step by step:**

1. **Read config from environment:**
   ```typescript
   const PORT = process.env.MCP_PORT ? parseInt(process.env.MCP_PORT, 10) : 3000;
   ```

2. **Create Express app:**
   ```typescript
   const app = express();
   app.use(express.json());
   app.use(cors({
     exposedHeaders: ["Mcp-Session-Id"],
     origin: "*", // In production, restrict to LibreChat's origin
   }));
   ```

3. **Create a factory function `getServer()`** that creates a new McpServer instance per session and registers all tools on it:
   ```typescript
   function getServer(): McpServer {
     const server = new McpServer({
       name: "365-mcp",
       version: "0.1.0",
     });

     const graphClient = getGraphClient();
     registerMailTools(server, graphClient);
     registerCalendarTools(server, graphClient);

     return server;
   }
   ```

4. **Session management** — use a `Map<string, NodeStreamableHTTPServerTransport>` to store active transports by session ID:
   ```typescript
   const transports = new Map<string, NodeStreamableHTTPServerTransport>();
   ```

5. **Health check endpoint:**
   ```typescript
   app.get("/health", (_req, res) => {
     res.json({ status: "ok" });
   });
   ```

6. **POST /mcp handler** — handles JSON-RPC requests:
   ```typescript
   app.post("/mcp", async (req: Request, res: Response) => {
     const sessionId = req.headers["mcp-session-id"] as string | undefined;

     try {
       let transport: NodeStreamableHTTPServerTransport;

       if (sessionId && transports.has(sessionId)) {
         // Reuse existing transport
         transport = transports.get(sessionId)!;
       } else if (!sessionId && isInitializeRequest(req.body)) {
         // New session
         transport = new NodeStreamableHTTPServerTransport({
           sessionIdGenerator: () => randomUUID(),
           onsessioninitialized: (id) => {
             transports.set(id, transport);
           },
         });

         transport.onclose = () => {
           const sid = transport.sessionId;
           if (sid) transports.delete(sid);
         };

         const server = getServer();
         await server.connect(transport);
         await transport.handleRequest(req, res, req.body);
         return;
       } else {
         res.status(400).json({
           jsonrpc: "2.0",
           error: { code: -32000, message: "Bad request: no valid session" },
           id: null,
         });
         return;
       }

       await transport.handleRequest(req, res, req.body);
     } catch (error) {
       console.error("Error handling MCP request:", error);
       if (!res.headersSent) {
         res.status(500).json({
           jsonrpc: "2.0",
           error: { code: -32603, message: "Internal server error" },
           id: null,
         });
       }
     }
   });
   ```

7. **GET /mcp handler** — for SSE streams:
   ```typescript
   app.get("/mcp", async (req: Request, res: Response) => {
     const sessionId = req.headers["mcp-session-id"] as string | undefined;
     if (!sessionId || !transports.has(sessionId)) {
       res.status(400).send("Invalid or missing session ID");
       return;
     }
     const transport = transports.get(sessionId)!;
     await transport.handleRequest(req, res);
   });
   ```

8. **DELETE /mcp handler** — for session termination:
   ```typescript
   app.delete("/mcp", async (req: Request, res: Response) => {
     const sessionId = req.headers["mcp-session-id"] as string | undefined;
     if (!sessionId || !transports.has(sessionId)) {
       res.status(400).send("Invalid or missing session ID");
       return;
     }
     const transport = transports.get(sessionId)!;
     await transport.handleRequest(req, res);
   });
   ```

9. **Startup sequence:**
   ```typescript
   async function main() {
     console.log("365-MCP Server starting...");

     // Test Graph API connection (triggers device code auth on first run)
     try {
       await testConnection();
       console.log("Graph API connection verified.");
     } catch (error) {
       console.error("Graph API connection failed:", error);
       console.error("The server will start, but tools will fail until auth is completed.");
     }

     app.listen(PORT, () => {
       console.log(`365-MCP Server listening on port ${PORT}`);
       console.log(`MCP endpoint: http://localhost:${PORT}/mcp`);
       console.log(`Health check: http://localhost:${PORT}/health`);
     });
   }

   main().catch((error) => {
     console.error("Fatal error:", error);
     process.exit(1);
   });
   ```

10. **Graceful shutdown:**
    ```typescript
    process.on("SIGINT", async () => {
      console.log("Shutting down...");
      for (const [id, transport] of transports) {
        await transport.close();
        transports.delete(id);
      }
      process.exit(0);
    });

    process.on("SIGTERM", async () => {
      console.log("Shutting down...");
      for (const [id, transport] of transports) {
        await transport.close();
        transports.delete(id);
      }
      process.exit(0);
    });
    ```

---

## Phase 4 — Containerize

### Task 4.1: Create Dockerfile

**File:** `Dockerfile`

```dockerfile
FROM node:22-alpine AS builder

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY tsconfig.json ./
COPY src/ ./src/
RUN npm run build

FROM node:22-alpine

WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci --omit=dev
COPY --from=builder /app/dist ./dist

# Create directory for token cache persistence
RUN mkdir -p /data/token-cache

EXPOSE 3000

HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
  CMD wget -qO- http://localhost:3000/health || exit 1

CMD ["node", "dist/index.js"]
```

### Task 4.2: Create docker-compose.yml

**File:** `docker-compose.yml`

This is a standalone development compose file. For production, merge the `365-mcp` service into LibreChat's compose file.

```yaml
services:
  365-mcp:
    build: .
    container_name: 365-mcp
    ports:
      - "3000:3000"
    environment:
      - AZURE_CLIENT_ID=${AZURE_CLIENT_ID}
      - AZURE_TENANT_ID=${AZURE_TENANT_ID}
      - MCP_PORT=3000
    volumes:
      - token-cache:/data/token-cache
    restart: unless-stopped

volumes:
  token-cache:
```

### Task 4.3: Create .dockerignore

**File:** `.dockerignore`

```
node_modules
dist
.env
.git
*.md
```

---

## Phase 5 — LibreChat Integration

### Task 5.1: Configure LibreChat

Add this to LibreChat's `librechat.yaml` under `mcpServers`:

```yaml
mcpServers:
  365-mcp:
    type: streamable-http
    url: http://365-mcp:3000/mcp
    timeout: 30000
    title: "Microsoft 365"
    description: "Read emails and calendar events from Microsoft 365"
```

If running in the same Docker Compose network as LibreChat, `365-mcp` resolves via Docker DNS. If running separately, use `http://localhost:3000/mcp` or the appropriate host.

### Task 5.2: Add to LibreChat's docker-compose (if applicable)

Add the `365-mcp` service definition from Task 4.2 into LibreChat's `docker-compose.override.yml` and ensure it shares the same Docker network.

---

## Phase 6 — Testing & Verification

### Task 6.1: Local dev test

```bash
# Start in dev mode
cd /home/bwa/git/365-mcp
npm run dev

# First run: watch logs for device code prompt, authenticate in browser

# Test health endpoint
curl http://localhost:3000/health

# Test MCP endpoint with an initialize request
curl -X POST http://localhost:3000/mcp \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
      "protocolVersion": "2025-03-26",
      "capabilities": {},
      "clientInfo": { "name": "test", "version": "1.0.0" }
    }
  }'
```

### Task 6.2: Docker test

```bash
docker compose up --build
# Watch logs for device code prompt on first run
# Test same curl commands against http://localhost:3000
```

### Task 6.3: LibreChat integration test

1. Restart LibreChat after adding MCP config
2. Create an agent in LibreChat with the 365-mcp tools enabled
3. Test: "List my recent emails"
4. Test: "What meetings do I have this week?"

---

## Important Notes for Implementation

### Token cache persistence in Docker

The `@azure/identity-cache-persistence` plugin uses the OS credential store (libsecret on Linux). In Alpine containers, this may not be available. If token cache fails to persist:

**Fallback approach:** Instead of `@azure/identity-cache-persistence`, use `@azure/msal-node` directly with a custom file-based cache serializer that writes to `/data/token-cache/msal-cache.json` on the mounted volume.

If this fallback is needed, the `src/auth/graph-client.ts` implementation changes significantly. Use `@azure/msal-node`'s `PublicClientApplication` with `DeviceCodeRequest` and a custom `ICachePlugin` that reads/writes a JSON file. This is a known pattern documented in the MSAL Node samples.

### Import paths

All local imports must use `.js` extensions because of Node16 module resolution:
```typescript
import { getGraphClient } from "./auth/graph-client.js";  // NOT .ts
```

### zod version

The MCP SDK uses `zod/v4` internally. Use `zod` v3.25+ which ships both `zod` and `zod/v4` entry points. In tool schemas, import from `zod`:
```typescript
import { z } from "zod";
```
The SDK handles the v4 conversion internally.

### Error patterns

Every tool handler must follow this pattern:
```typescript
async (args) => {
  try {
    // ... Graph API call ...
    return {
      content: [{ type: "text", text: formattedResult }],
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      content: [{ type: "text", text: `Error: ${message}` }],
      isError: true,
    };
  }
}
```

### Package resolution note

If `@modelcontextprotocol/express` or `@modelcontextprotocol/node` do not exist as separate packages (the SDK structure may have changed), the transports may be available from `@modelcontextprotocol/sdk` directly:
```typescript
// Alternative imports if separate packages don't exist:
import { McpServer, isInitializeRequest } from "@modelcontextprotocol/sdk/server/index.js";
import { NodeStreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.node.js";
```
Check the actual published package structure during `npm install` and adjust imports accordingly.
