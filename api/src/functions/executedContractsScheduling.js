const { app } = require("@azure/functions");
const { graphFetch } = require("../graphClient");

const HOSTNAME = "wesconc.sharepoint.com";
const SITE_PATH = "/sites/Wesco";
const LIST_NAME = "Executed Contracts and Scheduling";
const GRAPH = "https://graph.microsoft.com/v1.0";

let schemaCache = null;

function graph(path, options = {}) {
  const url = path.startsWith("http") ? path : `${GRAPH}/${path}`;
  return graphFetch(url, options);
}

function normalized(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function authorized(request) {
  const encoded = request.headers.get("x-ms-client-principal");
  if (!encoded) return false;
  try {
    const principal = JSON.parse(Buffer.from(encoded, "base64").toString("utf8"));
    const roles = (principal.userRoles || []).map((role) => String(role).toLowerCase());
    const email = String(principal.userDetails || "").trim().toLowerCase();
    return roles.includes("authenticated") && email.endsWith("@wesconc.com");
  } catch {
    return false;
  }
}

function unauthorized() {
  return { status: 401, jsonBody: { error: "Sign in with your Wesco account." } };
}

const ALIASES = {
  projectNumber: ["Project Number", "Project #", "Job Number"],
  projectName: ["Project Name", "Title", "Job Name"],
  customer: ["Customer", "Customer Name"],
  location: ["Location", "Project Location", "Address"],
  approvalStatus: ["Approval Status", "Contract Status", "Status"],
  estimateCreated: ["Estimate Created", "Estimate Created?"],
  siteVisited: ["Site Visited", "Site Visited?"],
  pmAssigned: ["PM Assigned", "Project Manager Assigned"],
  projectManagerName: ["Project Manager Name", "Project Manager", "PM Name"],
  scope: ["Scope", "Scope of Work", "Description"],
};

async function schema() {
  if (schemaCache && schemaCache.expires > Date.now()) return schemaCache.value;
  const site = await graph(`sites/${HOSTNAME}:${SITE_PATH}`);
  const lists = await graph(`sites/${site.id}/lists?$select=id,displayName`);
  const list = (lists.value || []).find((item) => normalized(item.displayName) === normalized(LIST_NAME));
  if (!list) throw new Error(`SharePoint list "${LIST_NAME}" was not found.`);
  const columns = await graph(`sites/${site.id}/lists/${list.id}/columns?$select=name,displayName,hidden,readOnly`);
  const mapping = {};
  for (const [key, aliases] of Object.entries(ALIASES)) {
    const column = (columns.value || []).find((candidate) =>
      !candidate.hidden && aliases.some((alias) =>
        normalized(candidate.displayName) === normalized(alias) ||
        normalized(candidate.name) === normalized(alias)
      )
    );
    if (column) mapping[key] = column.name;
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
  for (const [key, name] of Object.entries(mapping)) record[key] = fields[name] == null ? "" : String(fields[name]);
  return record;
}

function clean(value, maximum = 1500) {
  return String(value == null ? "" : value).trim().slice(0, maximum);
}

function fieldsFrom(body, mapping) {
  const fields = {};
  for (const [key, name] of Object.entries(mapping)) {
    if (Object.prototype.hasOwnProperty.call(body, key)) fields[name] = clean(body[key]);
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
    if (!authorized(request)) return unauthorized();
    try {
      const current = await schema();
      if (request.method === "GET") {
        const data = await graph(
          `sites/${current.siteId}/lists/${current.listId}/items?expand=fields&$top=999&$orderby=lastModifiedDateTime desc`
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
    if (!authorized(request)) return unauthorized();
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
