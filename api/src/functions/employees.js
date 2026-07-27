const { app } = require("@azure/functions");
const { graphFetchFor, sanitizeForSheetName } = require("../graphClient");

const WORKBOOK = "hardware";
const TEMPLATE_SHEET = "TEMPLATE";

function gf(path, options) {
  return graphFetchFor(WORKBOOK, path, options);
}

function worksheetSegment(sheetName) {
  return `worksheets('${encodeURIComponent(sheetName)}')`;
}

// Unlike the equipment/trailer/vehicle logs, this workbook's per-sheet tables
// already existed with Excel-assigned names (HW_TEMPLATE4, HW_TEMPLATE5, ...)
// before this API existed, so there's no deterministic naming scheme to rely
// on. Each sheet has exactly one table, so just look it up by position.
async function getSheetTableName(sheetName) {
  const tablesResp = await gf(`/${worksheetSegment(sheetName)}/tables`);
  if (!tablesResp.value || tablesResp.value.length === 0) {
    throw new Error(`Sheet '${sheetName}' has no table`);
  }
  return tablesResp.value[0].name;
}

app.http("listEmployees", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "employees",
  handler: async (request, context) => {
    try {
      const sheetsResp = await gf("/worksheets");
      const sheets = sheetsResp.value.filter((s) => s.name.trim().toUpperCase() !== TEMPLATE_SHEET);

      const employees = await Promise.all(
        sheets.map(async (sheet) => {
          const seg = worksheetSegment(sheet.name);
          const nameRange = await gf(`/${seg}/range(address='A3')`);
          const tableName = await getSheetTableName(sheet.name);
          const rowsResp = await gf(`/${seg}/tables('${encodeURIComponent(tableName)}')/rows`);
          const itemCount = rowsResp.value.filter((r) => r.values[0][1]).length; // column B = device name

          return {
            sheet: sheet.name,
            name: (nameRange.values && nameRange.values[0] && nameRange.values[0][0]) || sheet.name.trim(),
            itemCount,
          };
        })
      );

      return { jsonBody: employees };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});

async function findUniqueSheetName(baseName) {
  const sheetsResp = await gf("/worksheets");
  const existing = new Set(sheetsResp.value.map((s) => s.name.trim()));
  let candidate = sanitizeForSheetName(baseName);
  let i = 2;
  while (existing.has(candidate)) {
    const suffix = ` (${i})`;
    candidate = sanitizeForSheetName(baseName).slice(0, 31 - suffix.length) + suffix;
    i++;
  }
  return candidate;
}

app.http("addEmployee", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "employees",
  handler: async (request, context) => {
    try {
      const body = await request.json();
      if (!body.name || !String(body.name).trim()) {
        return { status: 400, jsonBody: { error: "employee name is required" } };
      }

      const beforeResp = await gf("/worksheets");
      const beforeNames = new Set(beforeResp.value.map((s) => s.name));

      await gf(`/${worksheetSegment(TEMPLATE_SHEET)}/copy`, {
        method: "POST",
        body: JSON.stringify({ positionType: "End" }),
      });

      const afterResp = await gf("/worksheets");
      const newSheet = afterResp.value.find((s) => !beforeNames.has(s.name));
      if (!newSheet) throw new Error("Could not identify the newly copied worksheet");

      const finalSheetName = await findUniqueSheetName(body.name);
      await gf(`/${worksheetSegment(newSheet.name)}`, {
        method: "PATCH",
        body: JSON.stringify({ name: finalSheetName }),
      });

      await gf(`/${worksheetSegment(finalSheetName)}/range(address='A3')`, {
        method: "PATCH",
        body: JSON.stringify({ values: [[body.name]] }),
      });

      return { status: 201, jsonBody: { sheet: finalSheetName, name: body.name } };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});

module.exports = { getSheetTableName, worksheetSegment, gf };
