const { app } = require("@azure/functions");
const { dataverseJson } = require("../dataverseClient");

const ENTITY_SET = "wsc_approvalrequests";
const SELECT = [
  "wsc_approvalrequestid",
  "wsc_requestnumber",
  "wsc_datesubmitted",
  "wsc_submittedby",
  "wsc_requestdepartment",
  "wsc_projectname",
  "wsc_projectnumber",
  "wsc_requesttype",
  "wsc_vendorcustomer",
  "wsc_amount",
  "wsc_description",
  "wsc_approvername",
  "wsc_approvaltitle",
  "wsc_approvaldepartment",
  "wsc_approvalstatus",
  "wsc_decision",
  "wsc_approvercomments",
  "wsc_datereviewed",
  "wsc_electronicsignature",
  "wsc_clientupdatedat",
  "createdon",
  "modifiedon",
].join(",");

function text(value, maximum = 250) {
  return String(value || "").trim().slice(0, maximum);
}

function dateOnly(value) {
  const candidate = text(value, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(candidate) ? candidate : null;
}

function dateTime(value) {
  const parsed = new Date(value || "");
  return Number.isNaN(parsed.getTime()) ? new Date().toISOString() : parsed.toISOString();
}

function requestNumber(value) {
  const candidate = text(value, 100);
  return /^[A-Za-z0-9._-]{1,100}$/.test(candidate) ? candidate : "";
}

function statusValue(value) {
  if (value === "Approved") return "Approved";
  if (value === "Denied" || value === "Rejected") return "Denied";
  return "Pending";
}

function mapRecord(record) {
  return {
    id: record.wsc_approvalrequestid,
    requestId: record.wsc_requestnumber || "",
    dateSubmitted: String(record.wsc_datesubmitted || "").slice(0, 10),
    submittedBy: record.wsc_submittedby || "",
    requestDepartment: record.wsc_requestdepartment || "",
    projectName: record.wsc_projectname || "",
    projectNumber: record.wsc_projectnumber || "",
    requestType: record.wsc_requesttype || "",
    vendorCustomer: record.wsc_vendorcustomer || "",
    amount: record.wsc_amount == null ? "" : record.wsc_amount,
    description: record.wsc_description || "",
    approverName: record.wsc_approvername || "",
    approvalTitle: record.wsc_approvaltitle || "",
    approvalDepartment: record.wsc_approvaldepartment || "",
    status: statusValue(record.wsc_approvalstatus),
    decision: record.wsc_decision || "Pending",
    approverComments: record.wsc_approvercomments || "",
    dateReviewed: String(record.wsc_datereviewed || "").slice(0, 10),
    electronicSignature: record.wsc_electronicsignature || "",
    updatedAt:
      record.wsc_clientupdatedat ||
      record.modifiedon ||
      record.createdon ||
      "",
  };
}

function payloadFrom(body) {
  const requestId = requestNumber(body.requestId);
  if (
    !requestId ||
    !text(body.submittedBy, 200) ||
    !text(body.requestType, 150) ||
    !text(body.description, 4000)
  ) {
    return null;
  }

  const payload = {
    wsc_requestnumber: requestId,
    wsc_submittedby: text(body.submittedBy, 200),
    wsc_requestdepartment: text(body.requestDepartment, 150),
    wsc_projectname: text(body.projectName, 250),
    wsc_projectnumber: text(body.projectNumber, 100),
    wsc_requesttype: text(body.requestType, 150),
    wsc_vendorcustomer: text(body.vendorCustomer, 250),
    wsc_description: text(body.description, 4000),
    wsc_approvername: text(body.approverName, 200),
    wsc_approvaltitle: text(body.approvalTitle, 150),
    wsc_approvaldepartment: text(body.approvalDepartment, 150),
    wsc_approvalstatus: statusValue(body.status),
    wsc_decision: text(body.decision || "Pending", 50),
    wsc_approvercomments: text(body.approverComments, 4000),
    wsc_electronicsignature: text(body.electronicSignature, 200),
    wsc_clientupdatedat: dateTime(body.updatedAt),
  };

  const submitted = dateOnly(body.dateSubmitted);
  const reviewed = dateOnly(body.dateReviewed);
  if (submitted) payload.wsc_datesubmitted = submitted;
  if (reviewed) payload.wsc_datereviewed = reviewed;

  if (body.amount !== "" && body.amount != null) {
    const amount = Number(body.amount);
    if (Number.isFinite(amount) && amount >= 0) payload.wsc_amount = amount;
  }

  return payload;
}

app.http("approvalRequests", {
  methods: ["GET", "POST", "PATCH"],
  authLevel: "anonymous",
  route: "approval-requests",
  handler: async (request, context) => {
    try {
      if (request.method === "GET") {
        const result = await dataverseJson(
          `${ENTITY_SET}?$select=${SELECT}&$orderby=modifiedon desc`
        );
        return { jsonBody: { items: (result.value || []).map(mapRecord) } };
      }

      const body = await request.json();
      const payload = payloadFrom(body);
      if (!payload) {
        return {
          status: 400,
          jsonBody: {
            error: "Request number, submitted by, request type, and description are required.",
          },
        };
      }

      const key = payload.wsc_requestnumber;
      const saved = await dataverseJson(
        `${ENTITY_SET}(wsc_requestnumber='${key}')?$select=${SELECT}`,
        {
          method: "PATCH",
          headers: {
            Prefer: "return=representation",
          },
          body: JSON.stringify(payload),
        }
      );
      return {
        status: request.method === "POST" ? 201 : 200,
        jsonBody: { item: mapRecord(saved) },
      };
    } catch (error) {
      context.error(error);
      return {
        status: 500,
        jsonBody: {
          error: "The approval request service is temporarily unavailable.",
        },
      };
    }
  },
});
