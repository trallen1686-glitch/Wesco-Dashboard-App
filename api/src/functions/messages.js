const { app } = require("@azure/functions");
const { graphFetch } = require("../graphClient");

const SHEET_NAME = "Messages";
const TABLE_NAME = "WescoMessages";
const HEADERS = [["MessageId", "From", "To", "Text", "SentAt", "Read"]];
const STAFF = new Set([
  "Oscar Centeno",
  "Joseph St. John",
  "Chris O'Keefe",
  "Theo Allen",
  "Jesse Warren",
  "Wes Atkinson",
  "Stephanie Atkinson",
  "Rebecca Kirkman",
]);

function worksheetSegment() {
  return `worksheets('${encodeURIComponent(SHEET_NAME)}')`;
}

function tableSegment() {
  return `${worksheetSegment()}/tables('${encodeURIComponent(TABLE_NAME)}')`;
}

function isMissing(err) {
  return err && err.status === 404;
}

async function ensureMessagesTable() {
  try {
    await graphFetch(`/${tableSegment()}`);
    return;
  } catch (err) {
    if (!isMissing(err)) throw err;
  }

  try {
    await graphFetch("/worksheets/add", {
      method: "POST",
      body: JSON.stringify({ name: SHEET_NAME }),
    });
  } catch (err) {
    if (!(err.status === 400 || err.status === 409)) throw err;
  }

  await graphFetch(`/${worksheetSegment()}/range(address='A1:F1')`, {
    method: "PATCH",
    body: JSON.stringify({ values: HEADERS }),
  });

  try {
    await graphFetch(`/${worksheetSegment()}/tables/add`, {
      method: "POST",
      body: JSON.stringify({ address: "A1:F1", hasHeaders: true }),
    });
  } catch (err) {
    if (!(err.status === 400 || err.status === 409)) throw err;
  }

  const tables = await graphFetch(`/${worksheetSegment()}/tables`);
  const current = tables.value.find((t) => t.name === TABLE_NAME) || tables.value[0];
  if (!current) throw new Error("Could not create the Messages table.");

  if (current.name !== TABLE_NAME) {
    await graphFetch(
      `/${worksheetSegment()}/tables('${encodeURIComponent(current.name)}')`,
      {
        method: "PATCH",
        body: JSON.stringify({ name: TABLE_NAME }),
      }
    );
  }
}

function mapMessage(row) {
  const values = row.values && row.values[0] ? row.values[0] : [];
  return {
    index: row.index,
    id: String(values[0] || ""),
    from: String(values[1] || ""),
    to: String(values[2] || ""),
    text: String(values[3] || ""),
    sentAt: String(values[4] || ""),
    read: values[5] === true || String(values[5]).toLowerCase() === "true",
  };
}

function validStaff(name) {
  return STAFF.has(String(name || "").trim());
}

app.http("messages", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "messages",
  handler: async (request, context) => {
    try {
      await ensureMessagesTable();

      if (request.method === "GET") {
        const person = String(request.query.get("person") || "").trim();
        if (!validStaff(person)) {
          return { status: 400, jsonBody: { error: "A valid staff member is required." } };
        }

        const rows = await graphFetch(`/${tableSegment()}/rows`);
        const messages = rows.value
          .map(mapMessage)
          .filter((m) => m.id && (m.from === person || m.to === person))
          .sort((a, b) => a.sentAt.localeCompare(b.sentAt))
          .slice(-500)
          .map(({ index, ...message }) => message);

        return { jsonBody: messages };
      }

      const body = await request.json();
      const from = String(body.from || "").trim();
      const to = String(body.to || "").trim();
      const text = String(body.text || "").trim();

      if (!validStaff(from) || !validStaff(to) || from === to) {
        return { status: 400, jsonBody: { error: "Valid sender and recipient are required." } };
      }
      if (!text || text.length > 2000) {
        return { status: 400, jsonBody: { error: "Message must be between 1 and 2000 characters." } };
      }

      const message = {
        id: String(body.id || `msg_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`),
        from,
        to,
        text,
        sentAt: new Date().toISOString(),
        read: false,
      };

      await graphFetch(`/${tableSegment()}/rows`, {
        method: "POST",
        body: JSON.stringify({
          values: [[message.id, message.from, message.to, message.text, message.sentAt, false]],
        }),
      });

      return { status: 201, jsonBody: message };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: "Messages are temporarily unavailable." } };
    }
  },
});

app.http("markMessageRead", {
  methods: ["PATCH"],
  authLevel: "anonymous",
  route: "messages/{id}/read",
  handler: async (request, context) => {
    try {
      await ensureMessagesTable();

      const id = String(request.params.id || "").trim();
      const recipient = String(request.query.get("recipient") || "").trim();
      if (!id || !validStaff(recipient)) {
        return { status: 400, jsonBody: { error: "Message and recipient are required." } };
      }

      const rows = await graphFetch(`/${tableSegment()}/rows`);
      const row = rows.value.map(mapMessage).find((m) => m.id === id && m.to === recipient);
      if (!row) return { status: 404, jsonBody: { error: "Message not found." } };

      await graphFetch(`/${tableSegment()}/rows/itemAt(index=${row.index})/range`, {
        method: "PATCH",
        body: JSON.stringify({
          values: [[row.id, row.from, row.to, row.text, row.sentAt, true]],
        }),
      });

      return { jsonBody: { ok: true } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: "Could not update message status." } };
    }
  },
});
