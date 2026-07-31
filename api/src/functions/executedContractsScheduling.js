const { app } = require("@azure/functions");
const { ConfidentialClientApplication } = require("@azure/msal-node");

const HOSTNAME = "wesconc.sharepoint.com";
const SITE_PATH = "/sites/Wesco";
const LIST_NAME = "Executed Contracts and Scheduling";
const GRAPH = "https://graph.microsoft.com/v1.0";

let schemaCache = null;
const authClients = new Map();
const authTokens = new Map();

function credentialSets() {
  const candidates = [
    {
      name: "equipment",
      clientId: process.env.EQUIPMENT_GRAPH_CLIENT_ID,
      tenantId: process.env.EQUIPMENT_GRAPH_TENANT_ID,
      clientSecret: process.env.EQUIPMENT_GRAPH_CLIENT_SECRET,
    },
    {
      name: "graph",
      clientId: process.env.GRAPH_CLIENT_ID,
      tenantId: process.env.GRAPH_TENANT_ID,
      clientSecret: process.env.GRAPH_CLIENT_SECRET,
    },
  ].filter((candidate) => candidate.clientId && candidate.tenantId && candidate.clientSecret);

  return candidates.filter((candidate, index) =>
    candidates.findIndex((other) =>
      other.clientId === candidate.clientId &&
      other.tenantId === candidate.tenantId &&
      other.clientSecret === candidate.clientSecret
    ) === index
  );
}

async function accessToken(credentials) {
  const cached = authTokens.get(credentials.name);
  if (cached && cached.expires > Date.now() + 60_000) return cached.value;
  let client = authClients.get(credentials.name);
  if (!client) {
    client = new ConfidentialClientApplication({
      auth: {
        clientId: credentials.clientId,
        authority: `https://login.microsoftonline.com/${credentials.tenantId}`,
        clientSecret: credentials.clientSecret,
      },
    });
    authClients.set(credentials.name, client);
  }
  const result = await client.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  authTokens.set(credentials.name, {
    value: result.accessToken,
    expires: new Date(result.expiresOn).getTime(),
  });
  return result.accessToken;
}

async function graph(path, options = {}) {
  const url = path.startsWith("http") ? path : `${GRAPH}/${path}`;
  const candidates = credentialSets();
  if (!candidates.length) throw new Error("Microsoft Graph credentials are not configured in Azure.");

  let lastPermissionError = null;
  for (const credentials of candidates) {
    const response = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${await accessToken(credentials)}`,
        Accept: "application/json",
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
    });
    if (response.ok) {
      if (response.status === 204) return null;
      return response.json();
    }
    const detail = await response.text().catch(() => "");
    const error = new Error(
      `Microsoft Graph ${options.method || "GET"} failed (${response.status}): ${detail.slice(0, 600)}`
    );
    error.status = response.status;
    if (response.status === 401 || response.status === 403) {
      lastPermissionError = error;
      continue;
    }
    throw error;
  }
  throw lastPermissionError || new Error("No configured Microsoft Graph connection can access the Wesco SharePoint site.");
}

function normalized(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

const ALIASES = {
  projectNumber: ["Project Number", "Project #", "Job Number"],
  projectName: ["Project Name", "Title", "Job Name"],
  customer: ["Customer", "Customer Name"],
  location: ["Location", "Project Location", "Address"],
  approvalStatus: ["Approval Status", "Contract Status", "Status"],
  estimateCreated: ["Estimate Created", "Estimate Created?", "Estimate Complete", "Estimate Completed", "Estimate Made", "Estimate"],
  siteVisited: ["Site Visited", "Site Visited?"],
  pmAssigned: ["PM Assigned", "Project Manager Assigned"],
  projectManagerName: ["PM Assigned To", "Assigned To", "Project Manager Name", "Project Manager", "PM Name"],
  scope: ["Scope", "Scope of Work", "Description"],
};

function findColumn(columns, key) {
  const exact = columns.find((candidate) =>
    ALIASES[key].some((alias) =>
      normalized(candidate.displayName) === normalized(alias) ||
      normalized(candidate.name) === normalized(alias)
    )
  );
  if (exact) return exact;

  // Existing SharePoint lists sometimes preserve older display names or encode
  // punctuation in internal names. Use a narrow fallback without changing the list.
  if (key === "estimateCreated") {
    return columns.find((candidate) => {
      const name = normalized(candidate.displayName || candidate.name);
      return name.includes("estimate") &&
        (name.includes("create") || name.includes("complete") || name === "estimate");
    });
  }
  if (key === "projectManagerName") {
    return columns.find((candidate) => {
      const name = normalized(candidate.displayName || candidate.name);
      return name === "pmassignedto" || name === "assignedto" ||
        name.includes("projectmanagername");
    });
  }
  return undefined;
}

async function schema() {
  if (schemaCache && schemaCache.expires > Date.now()) return schemaCache.value;
  const site = await graph(`sites/${HOSTNAME}:${SITE_PATH}`);
  const lists = await graph(`sites/${site.id}/lists?$select=id,displayName`);
  const list = (lists.value || []).find((item) => normalized(item.displayName) === normalized(LIST_NAME));
  if (!list) throw new Error(`SharePoint list "${LIST_NAME}" was not found.`);
  const columns = (await graph(
    `sites/${site.id}/lists/${list.id}/columns?$select=name,displayName,hidden,readOnly,boolean,choice,text`
  )).value || [];

  const mapping = {};
  for (const key of Object.keys(ALIASES)) {
    const column = findColumn(columns, key);
    if (column) {
      mapping[key] = {
        name: column.name,
        type: column.boolean ? "boolean" : column.choice ? "choice" : "text",
      };
    }
  }
  const missing = ["projectNumber", "projectName", "customer", "approvalStatus"].filter((key) => !mapping[key]);
  if (missing.length) throw new Error(`SharePoint list is missing required columns: ${missing.join(", ")}.`);
  const value = { siteId: site.id, listId: list.id, mapping };
  schemaCache = { value, expires: Date.now() + 10 * 60 * 1000 };
  return value;
}

function mapItem(item, mapping) {
  const fields = item.fields || {};
  const record = { id: String(item.id), createdAt: item.createdDateTime || "", updatedAt: item.lastModifiedDateTime || "" };
  for (const [key, column] of Object.entries(mapping)) {
    const value = fields[column.name];
    record[key] = column.type === "boolean"
      ? (value === true || value === "true" || value === 1 ? "Yes" : "No")
      : (value == null ? "" : String(value));
  }
  return record;
}

function clean(value, maximum = 1500) {
  return String(value == null ? "" : value).trim().slice(0, maximum);
}

function fieldsFrom(body, mapping) {
  const fields = {};
  for (const [key, column] of Object.entries(mapping)) {
    if (!Object.prototype.hasOwnProperty.call(body, key)) continue;
    fields[column.name] = column.type === "boolean"
      ? String(body[key]).toLowerCase() === "yes"
      : clean(body[key]);
  }
  return fields;
}

function valid(body) {
  return clean(body.projectNumber, 200) && clean(body.projectName, 300) &&
    clean(body.customer, 300) && clean(body.approvalStatus, 100);
}

app.http("executedContractsScheduling", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "executed-contracts-scheduling",
  handler: async (request, context) => {
    try {
      const current = await schema();
      if (request.method === "GET") {
        const data = await graph(
          `sites/${current.siteId}/lists/${current.listId}/items?expand=fields&$top=999`
        );
        return { jsonBody: { items: (data.value || []).map((item) => mapItem(item, current.mapping)) } };
      }
      const body = await request.json();
      if (!valid(body)) return { status: 400, jsonBody: { error: "Complete every required field." } };
      const created = await graph(`sites/${current.siteId}/lists/${current.listId}/items`, {
        method: "POST",
        body: JSON.stringify({ fields: fieldsFrom(body, current.mapping) }),
      });
      const item = await graph(`sites/${current.siteId}/lists/${current.listId}/items/${created.id}?expand=fields`);
      return { status: 201, jsonBody: { item: mapItem(item, current.mapping) } };
    } catch (error) {
      context.error(error);
      return { status: 500, jsonBody: { error: "Executed Contracts and Scheduling could not sync with SharePoint.", detail: error.message } };
    }
  },
});

app.http("executedContractsSchedulingRecord", {
  methods: ["PATCH", "DELETE"],
  authLevel: "anonymous",
  route: "executed-contracts-scheduling/{id}",
  handler: async (request, context) => {
    const id = String(request.params.id || "");
    if (!/^\d+$/.test(id)) return { status: 400, jsonBody: { error: "Invalid SharePoint item ID." } };
    try {
      const current = await schema();
      const path = `sites/${current.siteId}/lists/${current.listId}/items/${id}`;
      if (request.method === "DELETE") {
        await graph(path, { method: "DELETE" });
        return { status: 204 };
      }
      const body = await request.json();
      if (!valid(body)) return { status: 400, jsonBody: { error: "Complete every required field." } };
      await graph(`${path}/fields`, {
        method: "PATCH",
        body: JSON.stringify(fieldsFrom(body, current.mapping)),
      });
      const item = await graph(`${path}?expand=fields`);
      return { jsonBody: { item: mapItem(item, current.mapping) } };
    } catch (error) {
      context.error(error);
      return { status: 500, jsonBody: { error: "The SharePoint record could not be updated.", detail: error.message } };
    }
  },
});
