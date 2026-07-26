const { app } = require("@azure/functions");
const { randomUUID } = require("crypto");
const { dataverseFetch, dataverseJson } = require("../dataverseClient");

const ENTITY_SET = "new_executivetasks";
const WES_EMAIL = "watkinson@wesconc.com";
const PORTAL_ADMINS = new Set([WES_EMAIL, "tallen@wesconc.com"]);
const STAFF = new Map([
  ["ocenteno@wesconc.com", "Oscar Centeno"],
  ["jstjohn@wesconc.com", "Joseph St. John"],
  ["cokeefe@wesconc.com", "Chris O'Keefe"],
  ["tallen@wesconc.com", "Theo Allen"],
  ["jwarren@wesconc.com", "Jesse Warren"],
  [WES_EMAIL, "Wes Atkinson"],
  ["satkinson@wesconc.com", "Stephanie Atkinson"],
  ["rkirkman@wesconc.com", "Rebecca Kirkman"],
]);
const NAME_TO_EMAIL = new Map(
  Array.from(STAFF.entries()).map(([email, name]) => [name.toLowerCase(), email])
);
const STAFF_OBJECT_IDS = new Map([
  ["3461c317-ee66-49f4-b8e7-3ffd19773981", "cokeefe@wesconc.com"],
  ["b2eefa79-939f-4320-a303-de39516c3940", "jwarren@wesconc.com"],
  ["968d9e2e-6f07-4caa-8692-a498cad822b4", "jstjohn@wesconc.com"],
  ["94547ecd-135d-47bf-9ee7-081cc79fd0d4", "ocenteno@wesconc.com"],
  ["c2416b7e-20c3-4c8b-9b3b-9b14654de70a", "rkirkman@wesconc.com"],
  ["8e711b0a-77d9-4b29-a304-8dfecb1f5c75", "satkinson@wesconc.com"],
  ["129c350e-b8ef-4d57-beed-17f9a1e05f3e", "tallen@wesconc.com"],
  ["e64a8120-b2b4-4ee0-9f6d-b5e7b110c4d2", WES_EMAIL],
]);
function normalizeIdentity(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9@]/g, "");
}
const STAFF_ALIASES = new Map();
for (const [email, name] of STAFF) {
  STAFF_ALIASES.set(normalizeIdentity(email), email);
  STAFF_ALIASES.set(normalizeIdentity(email.split("@")[0]), email);
  STAFF_ALIASES.set(normalizeIdentity(name), email);
}
for (const [objectId, email] of STAFF_OBJECT_IDS) {
  STAFF_ALIASES.set(normalizeIdentity(objectId), email);
}
function resolveStaffEmail(principal) {
  const candidates = [principal.userDetails, principal.userId];
  for (const claim of principal.claims || []) {
    candidates.push(claim && (claim.val || claim.value));
  }
  for (const candidate of candidates) {
    const email = STAFF_ALIASES.get(normalizeIdentity(candidate));
    if (email) return email;
  }
  // Azure AD guest and federated accounts can arrive as
  // tallen_wesconc.com#EXT#@tenant.onmicrosoft.com instead of the normal UPN.
  // Match the compact company email embedded in that value without granting
  // access to an unrelated account that merely shares a display name.
  for (const candidate of candidates) {
    const compactCandidate = normalizeIdentity(candidate).replace(/@/g, "");
    for (const email of STAFF.keys()) {
      const compactEmail = normalizeIdentity(email).replace(/@/g, "");
      if (compactCandidate.includes(compactEmail)) return email;
    }
  }
  return "";
}
const SELECT = [
  "new_executivetaskid",
  "new_taskkey",
  "new_groupkey",
  "new_tasktitle",
  "new_assignedto",
  "new_assigneeemail",
  "new_projectnumber",
  "new_location",
  "new_dateassigned",
  "new_duedate",
  "new_communicationmethod",
  "new_priority",
  "new_status",
  "new_completedat",
  "new_photoname",
  "new_documentname",
  "new_instructions",
].join(",");

function currentUser(request) {
  const encoded = request.headers.get("x-ms-client-principal");
  if (!encoded) return null;
  try {
    const principal = JSON.parse(Buffer.from(encoded, "base64").toString("utf8"));
    const roles = (principal.userRoles || []).map((role) => String(role).toLowerCase());
    const email = resolveStaffEmail(principal);
    if (!email || !roles.includes("authenticated")) return null;
    return { email, name: STAFF.get(email) };
  } catch {
    return null;
  }
}

function identityBody(user) {
  return {
    email: user.email,
    name: user.name,
    canViewAll: PORTAL_ADMINS.has(user.email),
    canCreateTasks: PORTAL_ADMINS.has(user.email),
  };
}

function unauthorized() {
  return { status: 401, jsonBody: { error: "Sign in with your Wesco account." } };
}

function odataEscape(value) {
  return String(value).replace(/'/g, "''");
}

function validId(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
    String(value || "")
  );
}

function mapTask(record) {
  const id = record.new_executivetaskid;
  return {
    id,
    taskKey: record.new_taskkey || "",
    groupKey: record.new_groupkey || "",
    title: record.new_tasktitle || "",
    assignee: record.new_assignedto || "",
    assigneeEmail: record.new_assigneeemail || "",
    projectNumber: record.new_projectnumber || "",
    location: record.new_location || "",
    taskDate: record.new_dateassigned || "",
    dueDate: record.new_duedate || "",
    channel: record.new_communicationmethod || "",
    priority: record.new_priority || "Medium",
    status: record.new_status || "Open",
    completedAt: record.new_completedat || "",
    photoName: record.new_photoname || "",
    documentName: record.new_documentname || "",
    notes: record.new_instructions || "",
    photoUrl: record.new_photoname
      ? `/api/executive-tasks/${id}/files/photo`
      : "",
    documentUrl: record.new_documentname
      ? `/api/executive-tasks/${id}/files/document`
      : "",
  };
}

async function getRecord(id) {
  if (!validId(id)) return null;
  try {
    return await dataverseJson(`${ENTITY_SET}(${id})?$select=${SELECT}`);
  } catch (error) {
    if (error.status === 404) return null;
    throw error;
  }
}

function canAccess(user, record) {
  return (
    PORTAL_ADMINS.has(user.email) ||
    String(record.new_assigneeemail || "").toLowerCase() === user.email
  );
}

function decodeUpload(upload) {
  if (!upload) return null;
  const name = String(upload.name || "attachment").trim().slice(0, 255);
  const data = String(upload.data || "").replace(/^data:[^,]*,/, "");
  const bytes = Buffer.from(data, "base64");
  if (!bytes.length) throw new Error("The selected attachment is empty.");
  if (bytes.length > 4 * 1024 * 1024) {
    throw new Error("Each attachment must be 4 MB or smaller.");
  }
  return { name, bytes };
}

async function uploadFile(id, column, upload) {
  const file = decodeUpload(upload);
  if (!file) return;
  await dataverseFetch(`${ENTITY_SET}(${id})/${column}`, {
    method: "PATCH",
    headers: {
      "Content-Type": "application/octet-stream",
      "x-ms-file-name": file.name.replace(/[^\x20-\x7E]/g, "") || "attachment",
      "If-None-Match": "null",
    },
    body: file.bytes,
  });
}

app.http("executiveTaskIdentity", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "executive-tasks/me",
  handler: async (request) => {
    const user = currentUser(request);
    return user ? { jsonBody: identityBody(user) } : unauthorized();
  },
});

app.http("executiveTasks", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "executive-tasks",
  handler: async (request, context) => {
    const user = currentUser(request);
    if (!user) return unauthorized();

    try {
      if (request.method === "GET") {
        const requested = String(request.query.get("person") || "").trim();
        let filter = "";
        if (PORTAL_ADMINS.has(user.email)) {
          if (requested) {
            const email = NAME_TO_EMAIL.get(requested.toLowerCase()) || requested.toLowerCase();
            filter = `new_assigneeemail eq '${odataEscape(email)}'`;
          }
        } else {
          filter = `new_assigneeemail eq '${odataEscape(user.email)}'`;
        }

        let query = `?$select=${SELECT}&$orderby=createdon desc`;
        if (filter) query += `&$filter=${encodeURIComponent(filter)}`;
        const result = await dataverseJson(`${ENTITY_SET}${query}`);
        return {
          jsonBody: {
            tasks: (result.value || []).map(mapTask),
            viewer: identityBody(user),
          },
        };
      }

      if (!PORTAL_ADMINS.has(user.email)) {
        return {
          status: 403,
          jsonBody: { error: "Only Wes Atkinson and Theo Allen can create executive tasks." },
        };
      }

      const body = await request.json();
      const assignees = Array.isArray(body.assignees) ? body.assignees : [];
      if (!assignees.length) {
        return { status: 400, jsonBody: { error: "Choose at least one staff member." } };
      }

      const groupKey = randomUUID();
      const created = [];
      for (const assignee of assignees) {
        const email = String(assignee.email || "").trim().toLowerCase();
        const name = STAFF.get(email);
        if (!name) {
          return {
            status: 400,
            jsonBody: {
              error: `${email || "A recipient"} is not a registered staff login.`,
            },
          };
        }

        const payload = {
          new_taskkey: randomUUID(),
          new_groupkey: groupKey,
          new_tasktitle: String(body.title || "").trim().slice(0, 100),
          new_assignedto: name,
          new_assigneeemail: email,
          new_projectnumber: String(body.projectNumber || "").trim().slice(0, 100),
          new_location: String(body.location || "").trim().slice(0, 100),
          new_dateassigned: String(body.taskDate || "").slice(0, 50),
          new_duedate: String(body.dueDate || "").slice(0, 50),
          new_communicationmethod: String(body.channel || "").slice(0, 100),
          new_priority: String(body.priority || "Medium").slice(0, 50),
          new_status: "Open",
          new_completedat: "",
          new_photoname: String((body.photo || {}).name || "").slice(0, 100),
          new_documentname: String((body.document || {}).name || "").slice(0, 100),
          new_instructions: String(body.notes || "").slice(0, 10000),
        };

        const createdRecord = await dataverseJson(
          `${ENTITY_SET}?$select=new_executivetaskid`,
          {
            method: "POST",
            headers: { Prefer: "return=representation" },
            body: JSON.stringify(payload),
          }
        );
        const id = createdRecord.new_executivetaskid;
        try {
          await uploadFile(id, "new_photo", body.photo);
          await uploadFile(id, "new_document", body.document);
        } catch (error) {
          await dataverseFetch(`${ENTITY_SET}(${id})`, { method: "DELETE" }).catch(() => {});
          throw error;
        }
        created.push(id);
      }
      return { status: 201, jsonBody: { created, count: created.length } };
    } catch (error) {
      context.error(error);
      const clientError =
        /attachment|recipient|JSON/i.test(String(error.message || ""));
      return {
        status: clientError ? 400 : 500,
        jsonBody: {
          error: clientError
            ? error.message
            : "The task service is temporarily unavailable.",
        },
      };
    }
  },
});

app.http("executiveTaskRecord", {
  methods: ["PATCH", "DELETE"],
  authLevel: "anonymous",
  route: "executive-tasks/{id}",
  handler: async (request, context) => {
    const user = currentUser(request);
    if (!user) return unauthorized();

    try {
      const id = String(request.params.id || "");
      const record = await getRecord(id);
      if (!record) return { status: 404, jsonBody: { error: "Task not found." } };
      if (!canAccess(user, record)) {
        return {
          status: 403,
          jsonBody: { error: "You can only update your own task list." },
        };
      }

      if (request.method === "DELETE") {
        if (!PORTAL_ADMINS.has(user.email)) {
          return {
            status: 403,
            jsonBody: { error: "Only Wes Atkinson and Theo Allen can delete executive tasks." },
          };
        }
        await dataverseFetch(`${ENTITY_SET}(${id})`, {
          method: "DELETE",
          headers: { "If-Match": "*" },
        });
        return { status: 204 };
      }

      const body = await request.json();
      const status = String(body.status || "");
      if (!["Open", "Completed"].includes(status)) {
        return { status: 400, jsonBody: { error: "Invalid task status." } };
      }
      const completedAt = status === "Completed" ? new Date().toISOString() : "";
      await dataverseFetch(`${ENTITY_SET}(${id})`, {
        method: "PATCH",
        headers: { "If-Match": "*" },
        body: JSON.stringify({
          new_status: status,
          new_completedat: completedAt,
        }),
      });
      return { jsonBody: { status, completedAt } };
    } catch (error) {
      context.error(error);
      return {
        status: 500,
        jsonBody: { error: "The task service is temporarily unavailable." },
      };
    }
  },
});

app.http("executiveTaskFile", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "executive-tasks/{id}/files/{kind}",
  handler: async (request, context) => {
    const user = currentUser(request);
    if (!user) return unauthorized();

    try {
      const id = String(request.params.id || "");
      const kind = String(request.params.kind || "");
      if (!["photo", "document"].includes(kind)) {
        return { status: 404, jsonBody: { error: "File not found." } };
      }
      const record = await getRecord(id);
      if (!record) return { status: 404, jsonBody: { error: "Task not found." } };
      if (!canAccess(user, record)) {
        return {
          status: 403,
          jsonBody: { error: "You can only open your own task files." },
        };
      }

      const column = kind === "photo" ? "new_photo" : "new_document";
      const nameField = kind === "photo" ? "new_photoname" : "new_documentname";
      const response = await dataverseFetch(
        `${ENTITY_SET}(${id})/${column}/$value`,
        { headers: { "Content-Type": "application/octet-stream" } }
      );
      const bytes = Buffer.from(await response.arrayBuffer());
      const filename = String(record[nameField] || `task-${kind}`).replace(/["\r\n]/g, "");
      return {
        status: 200,
        body: bytes,
        headers: {
          "Content-Type": response.headers.get("content-type") || "application/octet-stream",
          "Content-Disposition": `${
            kind === "photo" ? "inline" : "attachment"
          }; filename="${filename}"`,
        },
      };
    } catch (error) {
      context.error(error);
      return { status: 500, jsonBody: { error: "The file could not be opened." } };
    }
  },
});

module.exports = {
  STAFF,
  PORTAL_ADMINS,
  resolveStaffEmail,
  currentUser,
  identityBody,
  canAccess,
  mapTask,
};
