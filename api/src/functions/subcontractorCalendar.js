const { app } = require("@azure/functions");
const { graphFetch } = require("../graphClient");

const DEFAULT_OWNER = "tallen@wesconc.com";
const CALENDARS = [
  {
    name: "El Bruzz - Esteban",
    owner: "Oscar Centeno Nieto",
    id: "AAMkADYzOTRjM2FiLTNmOGMtNDE4ZS1hOTJiLWNhN2U5ZGEwZTg5YgBGAAAAAAC6DW7Kzg_sS5BcQPqo4yPNBwCkVNVFY9M3Q4AcfAVFudcgAAAAAAEGAACkVNVFY9M3Q4AcfAVFudcgAAA5sLVhAAA=",
  },
  {
    name: "Anzola Trim - Antonio",
    owner: "Oscar Centeno Nieto",
    id: "AAMkADYzOTRjM2FiLTNmOGMtNDE4ZS1hOTJiLWNhN2U5ZGEwZTg5YgBGAAAAAAC6DW7Kzg_sS5BcQPqo4yPNBwCkVNVFY9M3Q4AcfAVFudcgAAAAAAEGAACkVNVFY9M3Q4AcfAVFudcgAAA5sLViAAA=",
  },
  {
    name: "Guillermo",
    owner: "Oscar Centeno Nieto",
    id: "AAMkADYzOTRjM2FiLTNmOGMtNDE4ZS1hOTJiLWNhN2U5ZGEwZTg5YgBGAAAAAAC6DW7Kzg_sS5BcQPqo4yPNBwCkVNVFY9M3Q4AcfAVFudcgAAAAAAEGAACkVNVFY9M3Q4AcfAVFudcgAAA5sLVjAAA=",
  },
  {
    name: "Yohan - Rooster Roofing",
    owner: "Oscar Centeno Nieto",
    id: "AAMkADYzOTRjM2FiLTNmOGMtNDE4ZS1hOTJiLWNhN2U5ZGEwZTg5YgBGAAAAAAC6DW7Kzg_sS5BcQPqo4yPNBwCkVNVFY9M3Q4AcfAVFudcgAAAAAAEGAACkVNVFY9M3Q4AcfAVFudcgAAA5sLVkAAA=",
  },
  {
    name: "Rogelio",
    owner: "Wes Atkinson",
    id: "AAMkADYzOTRjM2FiLTNmOGMtNDE4ZS1hOTJiLWNhN2U5ZGEwZTg5YgBGAAAAAAC6DW7Kzg_sS5BcQPqo4yPNBwCkVNVFY9M3Q4AcfAVFudcgAAAAAAEGAACkVNVFY9M3Q4AcfAVFudcgAABQeZW0AAA=",
  },
  {
    name: "Towa Construction (Rigo)",
    owner: "Wes Atkinson",
    id: "AAMkADYzOTRjM2FiLTNmOGMtNDE4ZS1hOTJiLWNhN2U5ZGEwZTg5YgBGAAAAAAC6DW7Kzg_sS5BcQPqo4yPNBwCkVNVFY9M3Q4AcfAVFudcgAAAAAAEGAACkVNVFY9M3Q4AcfAVFudcgAABQeZW1AAA=",
  },
  {
    name: "Pristine/Sik/Tulio/Chris/Josue/Commie/Samir",
    owner: "Wes Atkinson",
    id: "AAMkADYzOTRjM2FiLTNmOGMtNDE4ZS1hOTJiLWNhN2U5ZGEwZTg5YgBGAAAAAAC6DW7Kzg_sS5BcQPqo4yPNBwCkVNVFY9M3Q4AcfAVFudcgAAAAAAEGAACkVNVFY9M3Q4AcfAVFudcgAABQeZW2AAA=",
  },
];

function graphCalendarPath(owner, calendarId, suffix) {
  return (
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(owner)}` +
    `/calendars/${encodeURIComponent(calendarId)}${suffix}`
  );
}

function allowedCalendar(calendarId) {
  return CALENDARS.find((calendar) => calendar.id === calendarId);
}

app.http("subcontractorCalendar", {
  methods: ["GET", "POST", "PATCH"],
  authLevel: "anonymous",
  route: "subcontractor-calendar",
  handler: async (request, context) => {
    try {
      const owner = process.env.SUBCONTRACTOR_CALENDAR_OWNER_USER || DEFAULT_OWNER;

      if (request.method === "GET") {
        const url = new URL(request.url);
        const start = url.searchParams.get("start");
        const end = url.searchParams.get("end");
        const calendarId = url.searchParams.get("calendarId") || "all";
        if (!start || !end) {
          return { status: 400, jsonBody: { error: "A calendar date range is required." } };
        }

        const targets =
          calendarId === "all"
            ? CALENDARS
            : CALENDARS.filter((calendar) => calendar.id === calendarId);
        if (!targets.length) {
          return { status: 400, jsonBody: { error: "Unknown subcontractor calendar." } };
        }

        const groups = await Promise.allSettled(
          targets.map(async (calendar) => {
            const path =
              graphCalendarPath(owner, calendar.id, "/calendarView") +
              `?startDateTime=${encodeURIComponent(start)}` +
              `&endDateTime=${encodeURIComponent(end)}` +
              "&$orderby=start/dateTime";
            const response = await graphFetch(path, {
              headers: { Prefer: 'outlook.timezone="Eastern Standard Time"' },
            });
            return (response.value || []).map((event) => ({
              ...event,
              calendarName: calendar.name,
              calendarOwner: calendar.owner,
              calendarId: calendar.id,
            }));
          })
        );

        const events = groups
          .filter((result) => result.status === "fulfilled")
          .flatMap((result) => result.value);
        const failedCalendars = groups
          .map((result, index) =>
            result.status === "rejected" ? targets[index].name : null
          )
          .filter(Boolean);

        if (failedCalendars.length === targets.length) {
          throw new Error("No subcontractor calendars could be loaded.");
        }

        return {
          headers: {
            "Cache-Control": "no-store",
            "Content-Type": "application/json; charset=utf-8",
          },
          jsonBody: { events, failedCalendars },
        };
      }

      const body = await request.json();
      const calendar = allowedCalendar(body.calendarId);
      if (!calendar) {
        return { status: 400, jsonBody: { error: "Unknown subcontractor calendar." } };
      }
      if (!body.event || typeof body.event !== "object") {
        return { status: 400, jsonBody: { error: "Calendar event details are required." } };
      }

      if (request.method === "POST") {
        const created = await graphFetch(
          graphCalendarPath(owner, calendar.id, "/events"),
          { method: "POST", body: JSON.stringify(body.event) }
        );
        return { status: 201, jsonBody: { event: created } };
      }

      if (!body.eventId) {
        return { status: 400, jsonBody: { error: "An event ID is required." } };
      }
      const updated = await graphFetch(
        graphCalendarPath(
          owner,
          calendar.id,
          `/events/${encodeURIComponent(body.eventId)}`
        ),
        { method: "PATCH", body: JSON.stringify(body.event) }
      );
      return { jsonBody: { event: updated } };
    } catch (err) {
      context.error(err);
      return {
        status: err.status || 500,
        headers: { "Cache-Control": "no-store" },
        jsonBody: { error: "The subcontractor calendar service is temporarily unavailable." },
      };
    }
  },
});
