const { app } = require("@azure/functions");
const { graphFetch } = require("../graphClient");

const SHEET_NAME = "Site Visits";
const TABLE_NAME = "WescoSiteVisits";
const HEADERS = [[
  "VisitId", "Status", "Project", "Customer", "Address", "Scope",
  "ScheduledDate", "AssignedSelected", "AssignedOther", "AssignedDisplay",
  "VisitedDate", "VisitedSelected", "VisitedOther", "VisitedDisplay", "UpdatedAt",
]];
const VALID_STATUSES = new Set(["unscheduled", "pending", "audit"]);

function worksheetSegment() {
  return `worksheets('${encodeURIComponent(SHEET_NAME)}')`;
}

function tableSegment() {
  return `${worksheetSegment()}/tables('${encodeURIComponent(TABLE_NAME)}')`;
}

function missing(error) {
  return error && (error.status === 400 || error.status === 404);
}

async function ensureTable() {
  try {
    await graphFetch(`/${tableSegment()}`);
    return;
  } catch (error) {
    if (!missing(error)) throw error;
  }

  try {
    await graphFetch("/worksheets/add", {
      method: "POST",
      body: JSON.stringify({ name: SHEET_NAME }),
    });
  } catch (error) {
    if (!(error.status === 400 || error.status === 409)) throw error;
  }

  await graphFetch(`/${worksheetSegment()}/range(address='A1:O1')`, {
    method: "PATCH",
    body: JSON.stringify({ values: HEADERS }),
  });

  try {
    await graphFetch(`/${worksheetSegment()}/tables/add`, {
      method: "POST",
      body: JSON.stringify({ address: "A1:O1", hasHeaders: true }),
    });
  } catch (error) {
    if (!(error.status === 400 || error.status === 409)) throw error;
  }

  const tables = await graphFetch(`/${worksheetSegment()}/tables`);
  const current = (tables.value || []).find((table) => table.name === TABLE_NAME) ||
    (tables.value || []).find((table) => table.showHeaders);
  if (!current) throw new Error("Could not create the Site Visits table.");
  if (current.name !== TABLE_NAME) {
    await graphFetch(`/${worksheetSegment()}/tables('${encodeURIComponent(current.name)}')`, {
      method: "PATCH",
      body: JSON.stringify({ name: TABLE_NAME }),
    });
  }
}

function text(value, maximum = 2000) {
  return String(value == null ? "" : value).trim().slice(0, maximum);
}

function list(value) {
  return Array.isArray(value) ? value.map((item) => text(item, 100)).filter(Boolean).slice(0, 25) : [];
}

function cleanItem(item, status) {
  return {
    id: text(item.id, 100) || `visit_${Date.now()}_${Math.random().toString(36).slice(2, 10)}`,
    project: text(item.project, 300),
    customer: text(item.customer, 300),
    address: text(item.address, 500),
    scope: text(item.scope, 3000),
    scheduledDate: text(item.scheduledDate, 10),
    assignedSelected: list(item.assignedSelected),
    assignedOther: text(item.assignedOther, 200),
    assignedDisplay: text(item.assignedDisplay, 500),
    visitedDate: text(item.visitedDate, 10),
    visitedSelected: list(item.visitedSelected),
    visitedOther: text(item.visitedOther, 200),
    visitedDisplay: text(item.visitedDisplay, 500),
    status,
  };
}

function mapRow(row) {
  const values = row.values && row.values[0] ? row.values[0] : [];
  const status = text(values[1], 20).toLowerCase();
  if (!VALID_STATUSES.has(status) || !text(values[0], 100)) return null;
  return {
    id: text(values[0], 100), status,
    project: text(values[2], 300), customer: text(values[3], 300),
    address: text(values[4], 500), scope: text(values[5], 3000),
    scheduledDate: text(values[6], 10),
    assignedSelected: text(values[7], 1000).split("|").filter(Boolean),
    assignedOther: text(values[8], 200), assignedDisplay: text(values[9], 500),
    visitedDate: text(values[10], 10),
    visitedSelected: text(values[11], 1000).split("|").filter(Boolean),
    visitedOther: text(values[12], 200), visitedDisplay: text(values[13], 500),
    updatedAt: text(values[14], 40),
  };
}

function stateFromRows(rows) {
  const state = { unscheduled: [], pending: [], audit: [] };
  for (const row of rows) {
    const item = mapRow(row);
    if (!item) continue;
    const { status, updatedAt, ...record } = item;
    state[status].push(record);
  }
  return state;
}

function rowFromItem(item) {
  return [[
    item.id, item.status, item.project, item.customer, item.address, item.scope,
    item.scheduledDate, item.assignedSelected.join("|"), item.assignedOther,
    item.assignedDisplay, item.visitedDate, item.visitedSelected.join("|"),
    item.visitedOther, item.visitedDisplay, new Date().toISOString(),
  ]];
}

async function findRow(id) {
  const rows = await graphFetch(`/${tableSegment()}/rows`);
  const values = rows.value || [];
  const index = values.findIndex((row) => text(row.values?.[0]?.[0], 100) === id);
  return { index, rows: values };
}

async function createRecord(raw, status) {
  const item = cleanItem(raw, status);
  const existing = await findRow(item.id);
  if (existing.index >= 0) {
    const error = new Error("A Site Visit with this ID already exists.");
    error.status = 409;
    throw error;
  }
  await graphFetch(`/${tableSegment()}/rows`, {
    method: "POST",
    body: JSON.stringify({ values: rowFromItem(item) }),
  });
  return item;
}

async function updateRecord(id, raw, status) {
  const item = cleanItem({ ...raw, id }, status);
  const existing = await findRow(id);
  if (existing.index < 0) {
    await graphFetch(`/${tableSegment()}/rows`, {
      method: "POST",
      body: JSON.stringify({ values: rowFromItem(item) }),
    });
    return item;
  }
  await graphFetch(`/${tableSegment()}/rows/itemAt(index=${existing.index})/range`, {
    method: "PATCH",
    body: JSON.stringify({ values: rowFromItem(item) }),
  });
  return item;
}

app.http("siteVisits", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "site-visits",
  handler: async (request, context) => {
    try {
      await ensureTable();
      if (request.method === "GET") {
        const rows = await graphFetch(`/${tableSegment()}/rows`);
        return {
          headers: { "Cache-Control": "no-store" },
          jsonBody: { state: stateFromRows(rows.value || []) },
        };
      }
      const body = await request.json();
      const status = text(body.status, 20).toLowerCase();
      if (!VALID_STATUSES.has(status)) {
        return { status: 400, jsonBody: { error: "Invalid Site Visit status." } };
      }
      const item = await createRecord(body.item || body, status);
      return { status: 201, jsonBody: { item } };
    } catch (error) {
      context.error(error);
      return {
        status: 500,
        jsonBody: { error: "The shared Site Visit Planner is temporarily unavailable." },
      };
    }
  },
});

app.http("siteVisitRecord", {
  methods: ["PUT", "DELETE"],
  authLevel: "anonymous",
  route: "site-visits/{id}",
  handler: async (request, context) => {
    const id = text(request.params.id, 100);
    if (!id) return { status: 400, jsonBody: { error: "Invalid Site Visit ID." } };
    try {
      await ensureTable();
      const existing = await findRow(id);
      if (request.method === "DELETE") {
        if (existing.index < 0) return { status: 404, jsonBody: { error: "Site Visit not found." } };
        await graphFetch(`/${tableSegment()}/rows/itemAt(index=${existing.index})`, { method: "DELETE" });
        return { status: 204 };
      }
      const body = await request.json();
      const status = text(body.status, 20).toLowerCase();
      if (!VALID_STATUSES.has(status)) {
        return { status: 400, jsonBody: { error: "Invalid Site Visit status." } };
      }
      const item = await updateRecord(id, body.item || body, status);
      return { jsonBody: { item } };
    } catch (error) {
      context.error(error);
      return {
        status: error.status === 409 ? 409 : 500,
        jsonBody: { error: error.status === 409 ? error.message : "The shared Site Visit Planner is temporarily unavailable." },
      };
    }
  },
});
