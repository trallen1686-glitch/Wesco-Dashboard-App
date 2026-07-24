const { app } = require("@azure/functions");
const { graphFetch, visitsTableName } = require("../graphClient");

function tableSegment(sheetName) {
  const table = visitsTableName(sheetName);
  return `worksheets('${encodeURIComponent(sheetName)}')/tables('${encodeURIComponent(table)}')`;
}

// Excel table rows return raw cell values, so a date-formatted cell comes back
// as a serial day-count number (e.g. 46227) instead of a string. Convert for display.
function excelSerialToISODate(value) {
  if (typeof value !== "number") return value || "";
  const utcDays = Math.floor(value - 25569);
  return new Date(utcDays * 86400 * 1000).toISOString().slice(0, 10);
}

app.http("listVisits", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "equipment/{sheet}/visits",
  handler: async (request, context) => {
    try {
      const sheetName = decodeURIComponent(request.params.sheet);
      const rowsResp = await graphFetch(`/${tableSegment(sheetName)}/rows`);
      const visits = rowsResp.value
        .map((r) => {
          const [date, hours, task, cost, notes] = r.values[0];
          return {
            index: r.index,
            date: excelSerialToISODate(date),
            hours: hours ?? "",
            task: task || "",
            cost: cost || 0,
            notes: notes || "",
          };
        })
        // Hide fully-blank rows (e.g. the placeholder row every new Visits table
        // starts with) from display without deleting anything in the workbook.
        .filter((v) => v.date || v.task || v.cost || v.notes);
      return { jsonBody: visits };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});

app.http("addVisit", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "equipment/{sheet}/visits",
  handler: async (request, context) => {
    try {
      const sheetName = decodeURIComponent(request.params.sheet);
      const body = await request.json();
      const row = [body.date || "", body.hours ?? "", body.task || "", body.cost ?? 0, body.notes || ""];
      await graphFetch(`/${tableSegment(sheetName)}/rows`, {
        method: "POST",
        body: JSON.stringify({ values: [row] }),
      });
      return { status: 201, jsonBody: { ok: true } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});

app.http("deleteVisit", {
  methods: ["DELETE"],
  authLevel: "anonymous",
  route: "equipment/{sheet}/visits/{index}",
  handler: async (request, context) => {
    try {
      const sheetName = decodeURIComponent(request.params.sheet);
      const index = Number(request.params.index);
      await graphFetch(`/${tableSegment(sheetName)}/rows/itemAt(index=${index})`, {
        method: "DELETE",
      });
      return { status: 204 };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
