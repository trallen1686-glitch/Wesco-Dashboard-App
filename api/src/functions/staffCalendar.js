const { app } = require("@azure/functions");
const { graphFetch } = require("../graphClient");

const DEFAULT_OWNER = "tallen@wesconc.com";
const CALENDAR_ID =
  "AAMkADYzOTRjM2FiLTNmOGMtNDE4ZS1hOTJiLWNhN2U5ZGEwZTg5YgBGAAAAAAC6DW7Kzg_sS5BcQPqo4yPNBwCkVNVFY9M3Q4AcfAVFudcgAAAAAAEGAACkVNVFY9M3Q4AcfAVFudcgAACl8R88AAA=";

function calendarPath(owner, suffix) {
  return (
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}` +
    `/calendars/${encodeURIComponent(CALENDAR_ID)}${suffix}`
  );
}

app.http("staffCalendar", {
  methods: ["GET", "POST", "PATCH", "DELETE"],
  authLevel: "anonymous",
  route: "staff-calendar",
  handler: async (request, context) => {
    try {
      const owner = process.env.STAFF_CALENDAR_OWNER_USER || DEFAULT_OWNER;

      if (request.method === "GET") {
        const url = new URL(request.url);
        const start = url.searchParams.get("start");
        const end = url.searchParams.get("end");
        if (!start || !end) {
          return {
            status: 400,
            jsonBody: { error: "A calendar date range is required." },
          };
        }
        const data = await graphFetch(
          calendarPath(owner, "/calendarView") +
            `?startDateTime=${encodeURIComponent(start)}` +
            `&endDateTime=${encodeURIComponent(end)}` +
            "&$orderby=start/dateTime&$top=250",
          { headers: { Prefer: 'outlook.timezone="Eastern Standard Time"' } }
        );
        return {
          headers: { "Cache-Control": "no-store" },
          jsonBody: { events: data.value || [] },
        };
      }

      const body = await request.json();
      if (request.method === "POST") {
        if (!body.event || typeof body.event !== "object") {
          return { status: 400, jsonBody: { error: "Calendar event details are required." } };
        }
        const created = await graphFetch(calendarPath(owner, "/events"), {
          method: "POST",
          body: JSON.stringify(body.event),
        });
        return { status: 201, jsonBody: { event: created } };
      }

      if (!body.eventId) {
        return { status: 400, jsonBody: { error: "An event ID is required." } };
      }

      if (request.method === "DELETE") {
        await graphFetch(
          calendarPath(owner, `/events/${encodeURIComponent(body.eventId)}`),
          { method: "DELETE" }
        );
        return { status: 204 };
      }

      if (!body.event || typeof body.event !== "object") {
        return { status: 400, jsonBody: { error: "Calendar event details are required." } };
      }
      const updated = await graphFetch(
        calendarPath(owner, `/events/${encodeURIComponent(body.eventId)}`),
        { method: "PATCH", body: JSON.stringify(body.event) }
      );
      return { jsonBody: { event: updated } };
    } catch (err) {
      context.error(err);
      return {
        status: err.status || 500,
        headers: { "Cache-Control": "no-store" },
        jsonBody: { error: "The staff calendar service is temporarily unavailable." },
      };
    }
  },
});
