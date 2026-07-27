const { app } = require("@azure/functions");
const { getSheetTableName, worksheetSegment, gf } = require("./employees");

// Same raw-value quirk as the other logs' Date columns, but "Date Purchased"
// here is inconsistently typed in the real data (some literal strings like
// "05-2026", some real dates) so only convert numeric (real date) values.
function excelSerialToISODate(value) {
  if (typeof value !== "number") return value || "";
  const utcDays = Math.floor(value - 25569);
  return new Date(utcDays * 86400 * 1000).toISOString().slice(0, 10);
}

function rowToItem(r) {
  const [num, device, model, modelNo, serial, datePurchased, warranty, used, notes] = r.values[0];
  return {
    index: r.index,
    num: num ?? "",
    device: device || "",
    model: model || "",
    modelNo: modelNo || "",
    serial: serial || "",
    datePurchased: excelSerialToISODate(datePurchased),
    warranty: warranty || "",
    used: used || "",
    notes: notes || "",
  };
}

app.http("listEmployeeHardware", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "employees/{sheet}/hardware",
  handler: async (request, context) => {
    try {
      const sheetName = decodeURIComponent(request.params.sheet);
      const tableName = await getSheetTableName(sheetName);
      const seg = `${worksheetSegment(sheetName)}/tables('${encodeURIComponent(tableName)}')`;
      const rowsResp = await gf(`/${seg}/rows`);
      const items = rowsResp.value.map(rowToItem).filter((item) => item.device);
      return { jsonBody: items };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});

app.http("addEmployeeHardware", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "employees/{sheet}/hardware",
  handler: async (request, context) => {
    try {
      const sheetName = decodeURIComponent(request.params.sheet);
      const body = await request.json();
      const tableName = await getSheetTableName(sheetName);
      const seg = `${worksheetSegment(sheetName)}/tables('${encodeURIComponent(tableName)}')`;

      const rowsResp = await gf(`/${seg}/rows`);
      const existingCount = rowsResp.value.filter((r) => r.values[0][1]).length;
      const nextNum = existingCount + 1;

      const row = [
        nextNum,
        body.device || "",
        body.model || "",
        body.modelNo || "",
        body.serial || "",
        body.datePurchased || "",
        body.warranty || "",
        body.used || "",
        body.notes || "",
      ];
      await gf(`/${seg}/rows`, {
        method: "POST",
        body: JSON.stringify({ values: [row] }),
      });
      return { status: 201, jsonBody: { ok: true, num: nextNum } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});

app.http("deleteEmployeeHardware", {
  methods: ["DELETE"],
  authLevel: "anonymous",
  route: "employees/{sheet}/hardware/{index}",
  handler: async (request, context) => {
    try {
      const sheetName = decodeURIComponent(request.params.sheet);
      const index = Number(request.params.index);
      const tableName = await getSheetTableName(sheetName);
      const seg = `${worksheetSegment(sheetName)}/tables('${encodeURIComponent(tableName)}')`;
      await gf(`/${seg}/rows/itemAt(index=${index})`, { method: "DELETE" });
      return { status: 204 };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
