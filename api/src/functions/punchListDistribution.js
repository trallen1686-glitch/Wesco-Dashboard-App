const { app } = require("@azure/functions");
const { dataverseFetch, dataverseJson } = require("../dataverseClient");

const ENTITY_SET = "new_punchlistdistributionitems";
const SELECT = [
  "new_punchlistdistributionitemid",
  "new_distributionitemkey",
  "new_punchlistkey",
  "new_itemkey",
  "new_jobnumber",
  "new_jobname",
  "new_punchdate",
  "new_projectmanager",
  "new_recipients",
  "new_description",
  "new_location",
  "new_urgent",
  "new_completed",
  "new_issent",
  "new_filedat",
  "new_clientupdatedat",
  "createdon",
  "modifiedon",
].join(",");

function text(value, maximum = 250) {
  return String(value == null ? "" : value).trim().slice(0, maximum);
}

function keyText(value, maximum = 110) {
  const candidate = text(value, maximum);
  return /^[A-Za-z0-9._:-]{1,110}$/.test(candidate) ? candidate : "";
}

function dateOnly(value) {
  const candidate = text(value, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(candidate) ? candidate : null;
}

function dateTime(value) {
  const parsed = new Date(value || "");
  return Number.isNaN(parsed.getTime()) ? new Date().toISOString() : parsed.toISOString();
}

function odataString(value) {
  return String(value).replace(/'/g, "''");
}

function recipients(value) {
  const list = Array.isArray(value) ? value : [];
  return [...new Set(list.map((name) => text(name, 200)).filter(Boolean))].slice(0, 50);
}

function mapRecord(record) {
  let recipientList = [];
  try {
    recipientList = JSON.parse(record.new_recipients || "[]");
  } catch (_error) {
    recipientList = text(record.new_recipients, 4000)
      .split(/\r?\n|,\s*/)
      .map((name) => name.trim())
      .filter(Boolean);
  }

  return {
    id: record.new_punchlistdistributionitemid,
    distributionItemKey: record.new_distributionitemkey || "",
    punchListKey: record.new_punchlistkey || "",
    itemKey: record.new_itemkey || "",
    jobNumber: record.new_jobnumber || "",
    jobName: record.new_jobname || "",
    punchDate: String(record.new_punchdate || "").slice(0, 10),
    projectManager: record.new_projectmanager || "",
    recipients: recipientList,
    description: record.new_description || "",
    location: record.new_location || "",
    urgent: Boolean(record.new_urgent),
    completed: Boolean(record.new_completed),
    isSent: Boolean(record.new_issent),
    filedAt: record.new_filedat || "",
    updatedAt:
      record.new_clientupdatedat ||
      record.modifiedon ||
      record.createdon ||
      "",
  };
}

function itemPayload(body, item) {
  const punchListKey = keyText(body.punchListKey);
  const itemKey = keyText(item.itemKey || item.id);
  if (!punchListKey || !itemKey || !text(item.description, 4000)) return null;

  const distributionItemKey = `${punchListKey}:${itemKey}`.slice(0, 220);
  const payload = {
    new_distributionitemkey: distributionItemKey,
    new_punchlistkey: punchListKey,
    new_itemkey: itemKey,
    new_jobnumber: text(body.jobNumber, 100),
    new_jobname: text(body.jobName, 250),
    new_projectmanager: text(body.projectManager, 200),
    new_recipients: JSON.stringify(recipients(body.recipients)).slice(0, 4000),
    new_description: text(item.description, 4000),
    new_location: text(item.location, 250),
    new_urgent: Boolean(item.urgent),
    new_completed: Boolean(item.completed),
    new_issent: Boolean(body.isSent),
    new_clientupdatedat: dateTime(body.updatedAt),
  };

  const punchDate = dateOnly(body.punchDate || body.date);
  if (punchDate) payload.new_punchdate = punchDate;
  if (body.filedAt) payload.new_filedat = dateTime(body.filedAt);
  return payload;
}

async function listRecords(includeUnsent) {
  const filter = includeUnsent ? "" : "&$filter=new_issent eq true";
  const result = await dataverseJson(
    `${ENTITY_SET}?$select=${SELECT}${filter}&$orderby=modifiedon desc`
  );
  return (result.value || []).map(mapRecord);
}

async function syncPunchList(body) {
  const punchListKey = keyText(body.punchListKey);
  const sourceItems = Array.isArray(body.items) ? body.items : [];
  if (!punchListKey || !sourceItems.length) {
    const error = new Error("A punch list key and at least one item are required.");
    error.status = 400;
    throw error;
  }

  const payloads = sourceItems.map((item) => itemPayload(body, item)).filter(Boolean);
  if (!payloads.length) {
    const error = new Error("At least one described punch list item is required.");
    error.status = 400;
    throw error;
  }

  const existingResult = await dataverseJson(
    `${ENTITY_SET}?$select=new_punchlistdistributionitemid,new_distributionitemkey` +
      `&$filter=new_punchlistkey eq '${odataString(punchListKey)}'`
  );
  const retainedKeys = new Set(payloads.map((payload) => payload.new_distributionitemkey));
  const removed = (existingResult.value || []).filter(
    (row) => !retainedKeys.has(row.new_distributionitemkey)
  );

  await Promise.all(
    payloads.map((payload) =>
      dataverseJson(
        `${ENTITY_SET}(new_distributionitemkey='${odataString(
          payload.new_distributionitemkey
        )}')?$select=${SELECT}`,
        {
          method: "PATCH",
          headers: { Prefer: "return=representation" },
          body: JSON.stringify(payload),
        }
      )
    )
  );

  await Promise.all(
    removed.map((row) =>
      dataverseFetch(`${ENTITY_SET}(${row.new_punchlistdistributionitemid})`, {
        method: "DELETE",
      })
    )
  );

  const refreshed = await dataverseJson(
    `${ENTITY_SET}?$select=${SELECT}` +
      `&$filter=new_punchlistkey eq '${odataString(punchListKey)}'` +
      "&$orderby=modifiedon desc"
  );
  return (refreshed.value || []).map(mapRecord);
}

async function updateCompletion(body) {
  const distributionItemKey = text(body.distributionItemKey, 220);
  if (!distributionItemKey || typeof body.completed !== "boolean") {
    const error = new Error("Distribution item key and completed value are required.");
    error.status = 400;
    throw error;
  }

  const saved = await dataverseJson(
    `${ENTITY_SET}(new_distributionitemkey='${odataString(
      distributionItemKey
    )}')?$select=${SELECT}`,
    {
      method: "PATCH",
      headers: { Prefer: "return=representation" },
      body: JSON.stringify({
        new_completed: body.completed,
        new_clientupdatedat: dateTime(body.updatedAt),
      }),
    }
  );
  return mapRecord(saved);
}

async function deletePunchList(body) {
  const punchListKey = keyText(body.punchListKey);
  if (!punchListKey) {
    const error = new Error("Punch list key is required.");
    error.status = 400;
    throw error;
  }

  const existing = await dataverseJson(
    `${ENTITY_SET}?$select=new_punchlistdistributionitemid` +
      `&$filter=new_punchlistkey eq '${odataString(punchListKey)}'`
  );
  await Promise.all(
    (existing.value || []).map((row) =>
      dataverseFetch(`${ENTITY_SET}(${row.new_punchlistdistributionitemid})`, {
        method: "DELETE",
      })
    )
  );
}

app.http("punchListDistribution", {
  methods: ["GET", "POST", "PATCH", "DELETE"],
  authLevel: "anonymous",
  route: "punch-list-distribution",
  handler: async (request, context) => {
    try {
      if (request.method === "GET") {
        const includeUnsent = request.query.get("includeUnsent") === "true";
        return {
          headers: { "Cache-Control": "no-store" },
          jsonBody: { items: await listRecords(includeUnsent) },
        };
      }

      const body = await request.json();
      if (request.method === "DELETE") {
        await deletePunchList(body);
        return { status: 204 };
      }
      if (request.method === "PATCH") {
        return { jsonBody: { item: await updateCompletion(body) } };
      }

      const items = await syncPunchList(body);
      return { status: 201, jsonBody: { items } };
    } catch (error) {
      context.error(error);
      return {
        status: error.status || 500,
        jsonBody: {
          error:
            error.status === 400
              ? error.message
              : "The Punch List Distribution service is temporarily unavailable.",
        },
      };
    }
  },
});

module.exports = { itemPayload, mapRecord, recipients };
