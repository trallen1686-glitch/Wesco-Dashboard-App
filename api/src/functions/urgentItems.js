const { app } = require("@azure/functions");
const { dataverseJson } = require("../dataverseClient");

const ENTITY_SET = "cre09_urgents";
const STAFF = new Set([
  "watkinson@wesconc.com",
  "satkinson@wesconc.com",
  "rkirkman@wesconc.com",
  "ocenteno@wesconc.com",
  "jstjohn@wesconc.com",
  "cokeefe@wesconc.com",
  "tallen@wesconc.com",
  "jwarren@wesconc.com",
]);
const SELECT = [
  "cre09_urgentid",
  "cre09_urgentissuetitle",
  "cre09_reportedby",
  "cre09_duedate",
  "cre09_assignedto",
  "cre09_issuedescription",
  "cre09_notes",
  "createdon",
  "modifiedon",
].join(",");

function normalizeIdentity(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9@]/g, "");
}

function currentUser(request) {
  const encoded = request.headers.get("x-ms-client-principal");
  if (!encoded) return null;
  try {
    const principal = JSON.parse(Buffer.from(encoded, "base64").toString("utf8"));
    const roles = (principal.userRoles || []).map((role) => String(role).toLowerCase());
    const candidates = [principal.userDetails, principal.userId];
    for (const claim of principal.claims || []) {
      candidates.push(claim && (claim.val || claim.value));
    }
    for (const candidate of candidates) {
      const normalized = normalizeIdentity(candidate);
      for (const email of STAFF) {
        const compactEmail = normalizeIdentity(email).replace(/@/g, "");
        if (
          normalized === normalizeIdentity(email) ||
          normalized.replace(/@/g, "").includes(compactEmail)
        ) {
          return roles.includes("authenticated") ? email : null;
        }
      }
    }
  } catch {
    return null;
  }
  return null;
}

function unauthorized() {
  return { status: 401, jsonBody: { error: "Sign in with your Wesco account." } };
}

function validId(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
    String(value || "")
  );
}

function metaFrom(record) {
  const notes = String(record.cre09_notes || "");
  if (!notes.startsWith("WESCO_URGENT_META:")) return {};
  try {
    return JSON.parse(notes.slice("WESCO_URGENT_META:".length));
  } catch {
    return {};
  }
}

function mapRecord(record) {
  const meta = metaFrom(record);
  const title = String(record.cre09_urgentissuetitle || "");
  const type = title.toLowerCase() === "urgent" ? "Urgent" : "Issue";
  return {
    id: record.cre09_urgentid,
    issuer: record.cre09_reportedby || "",
    dateIssued:
      meta.dateIssued ||
      String(record.createdon || "").slice(0, 10),
    dueDate: record.cre09_duedate || "",
    issuedFor: String(record.cre09_assignedto || "")
      .split(",")
      .map((value) => value.trim())
      .filter(Boolean),
    type,
    description: record.cre09_issuedescription || "",
    resolved: Boolean(meta.resolved),
    resolvedAt: meta.resolvedAt || null,
    createdAt: meta.createdAt || record.createdon || "",
  };
}

function cleanText(value, maximum) {
  return String(value || "").trim().slice(0, maximum);
}

app.http("urgentItems", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "urgent-items",
  handler: async (request, context) => {
    try {
      if (request.method === "GET") {
        const result = await dataverseJson(
          `${ENTITY_SET}?$select=${SELECT}&$orderby=createdon desc`
        );
        return { jsonBody: { items: (result.value || []).map(mapRecord) } };
      }

      const body = await request.json();
      const issuedFor = Array.isArray(body.issuedFor)
        ? body.issuedFor.map((value) => cleanText(value, 100)).filter(Boolean)
        : [];
      const type = body.type === "Urgent" ? "Urgent" : body.type === "Issue" ? "Issue" : "";
      if (
        !cleanText(body.issuer, 100) ||
        !cleanText(body.dateIssued, 10) ||
        !cleanText(body.dueDate, 10) ||
        !issuedFor.length ||
        !type ||
        !cleanText(body.description, 1500)
      ) {
        return { status: 400, jsonBody: { error: "Complete every required field." } };
      }
      const createdAt = new Date().toISOString();
      const payload = {
        cre09_urgentissuetitle: type,
        cre09_reportedby: cleanText(body.issuer, 100),
        cre09_duedate: cleanText(body.dueDate, 10),
        cre09_assignedto: issuedFor.join(", ").slice(0, 100),
        cre09_issuedescription: cleanText(body.description, 1500),
        cre09_notes: `WESCO_URGENT_META:${JSON.stringify({
          resolved: false,
          resolvedAt: null,
          createdAt,
          dateIssued: cleanText(body.dateIssued, 10),
        })}`,
      };
      const created = await dataverseJson(
        `${ENTITY_SET}?$select=${SELECT}`,
        {
          method: "POST",
          headers: { Prefer: "return=representation" },
          body: JSON.stringify(payload),
        }
      );
      return { status: 201, jsonBody: { item: mapRecord(created) } };
    } catch (error) {
      context.error(error);
      return {
        status: 500,
        jsonBody: { error: "The urgent items service is temporarily unavailable." },
      };
    }
  },
});

app.http("urgentItemRecord", {
  methods: ["PATCH", "DELETE"],
  authLevel: "anonymous",
  route: "urgent-items/{id}",
  handler: async (request, context) => {
    const id = request.params.id;
    if (!validId(id)) return { status: 400, jsonBody: { error: "Invalid item ID." } };
    try {
      if (request.method === "DELETE") {
        await dataverseJson(`${ENTITY_SET}(${id})`, { method: "DELETE" });
        return { status: 204 };
      }
      const body = await request.json();
      const resolved = Boolean(body.resolved);
      const current = await dataverseJson(
        `${ENTITY_SET}(${id})?$select=cre09_notes,createdon`
      );
      const currentMeta = metaFrom(current);
      await dataverseJson(`${ENTITY_SET}(${id})`, {
        method: "PATCH",
        body: JSON.stringify({
          cre09_notes: `WESCO_URGENT_META:${JSON.stringify({
            resolved,
            resolvedAt: resolved ? new Date().toISOString() : null,
            createdAt:
              currentMeta.createdAt ||
              cleanText(body.createdAt, 50) ||
              current.createdon ||
              new Date().toISOString(),
            dateIssued: currentMeta.dateIssued || String(current.createdon || "").slice(0, 10),
          })}`,
        }),
      });
      return { jsonBody: { ok: true } };
    } catch (error) {
      context.error(error);
      return {
        status: 500,
        jsonBody: { error: "The urgent item could not be updated." },
      };
    }
  },
});
