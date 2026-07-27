const { app } = require("@azure/functions");
const { graphFetchFor, visitsTableName, sanitizeForSheetName, odataQuote } = require("../graphClient");

const WORKBOOK = "vehicle";
const EXCLUDED_SHEETS = new Set(["Template"]);

function gf(path, options) {
  return graphFetchFor(WORKBOOK, path, options);
}

function worksheetSegment(sheetName) {
  return `worksheets('${odataQuote(sheetName)}')`;
}

function parseHeaderRange(values) {
  // values: rows of [label, value] for range G3:H12
  const fields = {};
  for (const row of values) {
    const label = row[0];
    const value = row[1];
    if (label) {
      const key = String(label).trim().replace(/:$/, "").trim();
      fields[key] = value;
    }
  }
  return fields;
}

app.http("listVehicles", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "vehicles",
  handler: async (request, context) => {
    try {
      const sheetsResp = await gf("/worksheets");
      const sheets = sheetsResp.value.filter((s) => !EXCLUDED_SHEETS.has(s.name.trim()));

      const vehicles = await Promise.all(
        sheets.map(async (sheet) => {
          const seg = worksheetSegment(sheet.name);
          const [machineRange, headerRange] = await Promise.all([
            gf(`/${seg}/range(address='B3')`),
            gf(`/${seg}/range(address='G3:H12')`),
          ]);
          const fields = parseHeaderRange(headerRange.values);
          return {
            sheet: sheet.name,
            tableName: visitsTableName(sheet.name),
            machine: (machineRange.values && machineRange.values[0] && machineRange.values[0][0]) || sheet.name.trim(),
            ownerName: fields["OWNER NAME"] || "",
            make: fields["VEHICLE MAKE"] || "",
            year: fields["YEAR"] || "",
            model: fields["MODEL"] || "",
            vin: fields["VIN NO"] || "",
            condition: fields["Condition"] || "",
            lastLocation: fields["Last Location"] || fields["Location"] || "",
            licenseNo: fields["LICENSE NO"] || "",
            periodStart: fields["PERIOD START"] || "",
            periodEnd: fields["PERIOD END"] || "",
            legacyCost: Number(fields["TOTAL SERVICE COSTS"]) || 0,
          };
        })
      );

      return { jsonBody: vehicles };
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

async function findAndRenameVisitsTable(sheetName, desiredTableName) {
  const seg = worksheetSegment(sheetName);
  const tablesResp = await gf(`/${seg}/tables`);
  let visitsTable = null;
  for (const table of tablesResp.value) {
    const tableSeg = `${seg}/tables('${odataQuote(table.name)}')`;
    const headerRange = await gf(`/${tableSeg}/headerRowRange`);
    const firstHeader = headerRange.values && headerRange.values[0] && headerRange.values[0][0];
    if (String(firstHeader).trim().toLowerCase() === "date") {
      visitsTable = table;
      break;
    }
  }
  if (!visitsTable) throw new Error(`Could not find the Visits table on new sheet '${sheetName}'`);

  let finalName = desiredTableName;
  for (let attempt = 0; attempt < 5; attempt++) {
    try {
      await gf(`/${seg}/tables('${odataQuote(visitsTable.name)}')`, {
        method: "PATCH",
        body: JSON.stringify({ name: finalName }),
      });
      return finalName;
    } catch (err) {
      if (err.status === 400 || err.status === 409) {
        finalName = `${desiredTableName}_${attempt + 2}`;
        continue;
      }
      throw err;
    }
  }
  throw new Error(`Could not rename Visits table on '${sheetName}' after retries`);
}

app.http("addVehicle", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "vehicles",
  handler: async (request, context) => {
    try {
      const body = await request.json();
      if (!body.machine || !String(body.machine).trim()) {
        return { status: 400, jsonBody: { error: "machine name is required" } };
      }

      const beforeResp = await gf("/worksheets");
      const beforeNames = new Set(beforeResp.value.map((s) => s.name));

      await gf("/worksheets('Template')/copy", {
        method: "POST",
        body: JSON.stringify({ positionType: "End" }),
      });

      const afterResp = await gf("/worksheets");
      const newSheet = afterResp.value.find((s) => !beforeNames.has(s.name));
      if (!newSheet) throw new Error("Could not identify the newly copied worksheet");

      const finalSheetName = await findUniqueSheetName(body.machine);
      await gf(`/${worksheetSegment(newSheet.name)}`, {
        method: "PATCH",
        body: JSON.stringify({ name: finalSheetName }),
      });

      const seg = worksheetSegment(finalSheetName);

      const today = new Date();
      const periodEnd = new Date(today.getTime() + 60 * 24 * 60 * 60 * 1000);
      const iso = (d) => d.toISOString().slice(0, 10);

      await Promise.all([
        gf(`/${seg}/range(address='B3')`, {
          method: "PATCH",
          body: JSON.stringify({ values: [[body.machine]] }),
        }),
        gf(`/${seg}/range(address='H4:H12')`, {
          method: "PATCH",
          body: JSON.stringify({
            values: [
              [iso(today)],
              [iso(periodEnd)],
              [body.ownerName || "Wesco Exteriors"],
              [body.make || ""],
              [body.year || ""],
              [body.model || ""],
              [body.vin || ""],
              [body.condition || ""],
              [body.licenseNo || ""],
            ],
          }),
        }),
      ]);

      const desiredTableName = visitsTableName(finalSheetName);
      const finalTableName = await findAndRenameVisitsTable(finalSheetName, desiredTableName);

      return {
        status: 201,
        jsonBody: {
          sheet: finalSheetName,
          tableName: finalTableName,
          machine: body.machine,
        },
      };
    } catch (err) {
      context.error(err);
      return { status: 500, jsonBody: { error: err.message } };
    }
  },
});
