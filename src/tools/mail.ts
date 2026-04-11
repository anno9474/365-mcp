import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { Client } from "@microsoft/microsoft-graph-client";

interface EmailAddress {
  emailAddress: { name: string; address: string };
}

interface Message {
  id: string;
  subject: string;
  from: EmailAddress;
  toRecipients: EmailAddress[];
  ccRecipients: EmailAddress[];
  receivedDateTime: string;
  bodyPreview: string;
  isRead: boolean;
  hasAttachments: boolean;
  body?: { contentType: string; content: string };
  attachments?: { id: string; name: string; contentType: string; size: number }[];
}

function formatAddress(addr: EmailAddress): string {
  return `${addr.emailAddress.name} <${addr.emailAddress.address}>`;
}

export function registerMailTools(server: McpServer, graphClient: Client): void {
  server.registerTool(
    "list_emails",
    {
      description:
        "List recent emails from your Microsoft 365 mailbox. Returns subject, sender, date, and a preview of each email.",
      inputSchema: z.object({
        folder: z
          .string()
          .optional()
          .describe(
            "Mail folder to list from. Defaults to 'inbox'. Other options: 'sentitems', 'drafts', 'deleteditems', 'archive'."
          ),
        top: z
          .number()
          .min(1)
          .max(50)
          .optional()
          .describe("Number of emails to return. Defaults to 10, max 50."),
        filter: z
          .string()
          .optional()
          .describe(
            'OData filter expression. Example: "isRead eq false" for unread emails.'
          ),
        search: z
          .string()
          .optional()
          .describe(
            'Search query string. Searches subject, body, and sender. Example: "quarterly report"'
          ),
      }),
    },
    async (args) => {
      try {
        const folder = args.folder || "inbox";
        const top = args.top || 10;

        let request = graphClient
          .api(`/me/mailFolders/${folder}/messages`)
          .top(top)
          .select(
            "id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments"
          )
          .orderby("receivedDateTime desc");

        if (args.filter) {
          request = request.filter(args.filter);
        }
        if (args.search) {
          request = request.search(`"${args.search}"`);
        }

        const response = await request.get();
        const messages: Message[] = response.value;

        if (messages.length === 0) {
          return {
            content: [{ type: "text" as const, text: `No emails found in ${folder}.` }],
          };
        }

        const lines = [`Found ${messages.length} emails in ${folder}:\n`];
        messages.forEach((msg, i) => {
          const preview = msg.bodyPreview
            ? msg.bodyPreview.substring(0, 150)
            : "(no preview)";
          lines.push(
            `${i + 1}. **${msg.subject || "(no subject)"}**`,
            `   From: ${formatAddress(msg.from)}`,
            `   Date: ${msg.receivedDateTime}`,
            `   Read: ${msg.isRead ? "yes" : "no"} | Attachments: ${msg.hasAttachments ? "yes" : "no"}`,
            `   Preview: ${preview}`,
            `   ID: ${msg.id}`,
            ""
          );
        });

        return {
          content: [{ type: "text" as const, text: lines.join("\n") }],
        };
      } catch (error) {
        const message =
          error instanceof Error ? error.message : String(error);
        return {
          content: [{ type: "text" as const, text: `Error listing emails: ${message}` }],
          isError: true,
        };
      }
    }
  );

  server.registerTool(
    "read_email",
    {
      description:
        "Read the full content of a specific email by its ID. Returns the complete email body, headers, and attachment list.",
      inputSchema: z.object({
        messageId: z
          .string()
          .describe(
            "The ID of the email message to read. Get this from list_emails."
          ),
      }),
    },
    async (args) => {
      try {
        const msg: Message = await graphClient
          .api(`/me/messages/${args.messageId}`)
          .select(
            "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,attachments"
          )
          .expand("attachments($select=id,name,contentType,size)")
          .get();

        const to = msg.toRecipients?.map(formatAddress).join(", ") || "none";
        const cc = msg.ccRecipients?.map(formatAddress).join(", ") || "none";
        const bodyContent = msg.body?.content || "(no body)";

        const lines = [
          `Subject: ${msg.subject || "(no subject)"}`,
          `From: ${formatAddress(msg.from)}`,
          `To: ${to}`,
          `CC: ${cc}`,
          `Date: ${msg.receivedDateTime}`,
          "",
          "--- Body ---",
          bodyContent,
          "",
          "--- Attachments ---",
        ];

        if (msg.attachments && msg.attachments.length > 0) {
          msg.attachments.forEach((att) => {
            const sizeKb = Math.round(att.size / 1024);
            lines.push(`- ${att.name} (${att.contentType}, ${sizeKb} KB)`);
          });
        } else {
          lines.push("No attachments");
        }

        return {
          content: [{ type: "text" as const, text: lines.join("\n") }],
        };
      } catch (error) {
        const message =
          error instanceof Error ? error.message : String(error);
        return {
          content: [{ type: "text" as const, text: `Error reading email: ${message}` }],
          isError: true,
        };
      }
    }
  );
}
