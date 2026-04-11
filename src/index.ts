import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import cors from "cors";
import express from "express";
import type { Request, Response } from "express";
import { getGraphClient, testConnection } from "./auth/graph-client.js";
import { registerMailTools } from "./tools/mail.js";
import { registerCalendarTools } from "./tools/calendar.js";

const PORT = process.env.MCP_PORT
  ? parseInt(process.env.MCP_PORT, 10)
  : 3000;

const app = express();
app.use(express.json());
app.use(
  cors({
    exposedHeaders: ["Mcp-Session-Id"],
    origin: "*",
  })
);

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

const transports = new Map<string, StreamableHTTPServerTransport>();

// Health check
app.get("/health", (_req, res) => {
  res.json({ status: "ok" });
});

// POST /mcp — JSON-RPC requests
app.post("/mcp", async (req: Request, res: Response) => {
  const sessionId = req.headers["mcp-session-id"] as string | undefined;

  try {
    let transport: StreamableHTTPServerTransport;

    if (sessionId && transports.has(sessionId)) {
      transport = transports.get(sessionId)!;
    } else if (!sessionId && isInitializeRequest(req.body)) {
      transport = new StreamableHTTPServerTransport({
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

// GET /mcp — SSE stream for existing sessions
app.get("/mcp", async (req: Request, res: Response) => {
  const sessionId = req.headers["mcp-session-id"] as string | undefined;
  if (!sessionId || !transports.has(sessionId)) {
    res.status(400).send("Invalid or missing session ID");
    return;
  }
  const transport = transports.get(sessionId)!;
  await transport.handleRequest(req, res);
});

// DELETE /mcp — session termination
app.delete("/mcp", async (req: Request, res: Response) => {
  const sessionId = req.headers["mcp-session-id"] as string | undefined;
  if (!sessionId || !transports.has(sessionId)) {
    res.status(400).send("Invalid or missing session ID");
    return;
  }
  const transport = transports.get(sessionId)!;
  await transport.handleRequest(req, res);
});

async function main() {
  console.log("365-MCP Server starting...");

  try {
    await testConnection();
    console.log("Graph API connection verified.");
  } catch (error) {
    console.error("Graph API connection failed:", error);
    console.error(
      "The server will start, but tools will fail until auth is completed."
    );
  }

  app.listen(PORT, () => {
    console.log(`365-MCP Server listening on port ${PORT}`);
    console.log(`MCP endpoint: http://localhost:${PORT}/mcp`);
    console.log(`Health check: http://localhost:${PORT}/health`);
  });
}

async function shutdown() {
  console.log("Shutting down...");
  for (const [id, transport] of transports) {
    await transport.close();
    transports.delete(id);
  }
  process.exit(0);
}

process.on("SIGINT", shutdown);
process.on("SIGTERM", shutdown);

main().catch((error) => {
  console.error("Fatal error:", error);
  process.exit(1);
});
