import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { Client } from "@microsoft/microsoft-graph-client";

interface CalendarEvent {
  id: string;
  subject: string;
  start: { dateTime: string; timeZone: string };
  end: { dateTime: string; timeZone: string };
  location?: { displayName?: string };
  organizer?: { emailAddress: { name: string; address: string } };
  attendees?: {
    emailAddress: { name: string; address: string };
    status: { response: string };
  }[];
  isAllDay: boolean;
  isCancelled: boolean;
  bodyPreview?: string;
  body?: { contentType: string; content: string };
  recurrence?: unknown;
  onlineMeeting?: { joinUrl?: string };
}

export function registerCalendarTools(
  server: McpServer,
  graphClient: Client
): void {
  server.registerTool(
    "list_events",
    {
      description:
        "List calendar events from your Microsoft 365 calendar within a specified date range.",
      inputSchema: z.object({
        startDateTime: z
          .string()
          .optional()
          .describe(
            "Start of the time range in ISO 8601 format. Example: '2026-04-11T00:00:00'. Defaults to now."
          ),
        endDateTime: z
          .string()
          .optional()
          .describe(
            "End of the time range in ISO 8601 format. Example: '2026-04-18T23:59:59'. Defaults to 7 days from now."
          ),
        top: z
          .number()
          .min(1)
          .max(50)
          .optional()
          .describe("Number of events to return. Defaults to 20, max 50."),
      }),
    },
    async (args) => {
      try {
        const now = new Date();
        const startDateTime =
          args.startDateTime || now.toISOString();
        const endDateTime =
          args.endDateTime ||
          new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString();
        const top = args.top || 20;

        const response = await graphClient
          .api("/me/calendarView")
          .query({ startDateTime, endDateTime })
          .top(top)
          .select(
            "id,subject,start,end,location,organizer,isAllDay,isCancelled,bodyPreview"
          )
          .orderby("start/dateTime asc")
          .get();

        const events: CalendarEvent[] = response.value;

        if (events.length === 0) {
          return {
            content: [
              {
                type: "text" as const,
                text: `No events found from ${startDateTime} to ${endDateTime}.`,
              },
            ],
          };
        }

        const lines = [
          `Found ${events.length} events from ${startDateTime} to ${endDateTime}:\n`,
        ];

        events.forEach((evt, i) => {
          const location =
            evt.location?.displayName || "No location";
          const organizer = evt.organizer
            ? `${evt.organizer.emailAddress.name} <${evt.organizer.emailAddress.address}>`
            : "Unknown";
          const preview = evt.bodyPreview
            ? evt.bodyPreview.substring(0, 150)
            : "(no preview)";
          const status = evt.isCancelled ? "Cancelled" : "Confirmed";

          lines.push(
            `${i + 1}. **${evt.subject || "(no subject)"}**`,
            `   When: ${evt.start.dateTime} — ${evt.end.dateTime} (${evt.start.timeZone})`,
            `   All day: ${evt.isAllDay ? "yes" : "no"}`,
            `   Location: ${location}`,
            `   Organizer: ${organizer}`,
            `   Status: ${status}`,
            `   Preview: ${preview}`,
            `   ID: ${evt.id}`,
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
          content: [
            { type: "text" as const, text: `Error listing events: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );

  server.registerTool(
    "read_event",
    {
      description:
        "Read the full details of a specific calendar event by its ID.",
      inputSchema: z.object({
        eventId: z
          .string()
          .describe(
            "The ID of the calendar event to read. Get this from list_events."
          ),
      }),
    },
    async (args) => {
      try {
        const evt: CalendarEvent = await graphClient
          .api(`/me/events/${args.eventId}`)
          .select(
            "id,subject,body,start,end,location,organizer,attendees,isAllDay,isCancelled,recurrence,onlineMeeting"
          )
          .get();

        const location =
          evt.location?.displayName || "No location";
        const organizer = evt.organizer
          ? `${evt.organizer.emailAddress.name} <${evt.organizer.emailAddress.address}>`
          : "Unknown";
        const status = evt.isCancelled ? "Cancelled" : "Confirmed";
        const bodyContent = evt.body?.content || "(no body)";

        const lines = [
          `Subject: ${evt.subject || "(no subject)"}`,
          `When: ${evt.start.dateTime} — ${evt.end.dateTime} (${evt.start.timeZone})`,
          `All day: ${evt.isAllDay ? "yes" : "no"}`,
          `Location: ${location}`,
          `Organizer: ${organizer}`,
          `Status: ${status}`,
          "",
          "--- Attendees ---",
        ];

        if (evt.attendees && evt.attendees.length > 0) {
          evt.attendees.forEach((att) => {
            lines.push(
              `- ${att.emailAddress.name} <${att.emailAddress.address}> — ${att.status.response}`
            );
          });
        } else {
          lines.push("No attendees");
        }

        lines.push("", "--- Body ---", bodyContent);

        lines.push("", "--- Online Meeting ---");
        if (evt.onlineMeeting?.joinUrl) {
          lines.push(evt.onlineMeeting.joinUrl);
        } else {
          lines.push("No online meeting link");
        }

        return {
          content: [{ type: "text" as const, text: lines.join("\n") }],
        };
      } catch (error) {
        const message =
          error instanceof Error ? error.message : String(error);
        return {
          content: [
            { type: "text" as const, text: `Error reading event: ${message}` },
          ],
          isError: true,
        };
      }
    }
  );
}
