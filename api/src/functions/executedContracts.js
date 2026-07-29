const { app } = require("@azure/functions");
const { graphFetch, graphUploadContent, odataQuote } = require("../graphClient");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const TABLE_NAME = "Table1";
const WORKBOOK_SHARE_URL = "https://wesconc.sharepoint.com/:x:/r/sites/Wesco/_layouts/15/Doc.aspx?sourcedoc=%7BB6546919-508C-4E64-A2AC-82E9B323B7CD%7D&file=Assign%20New%20Executed%20Project%20Template.xlsx&action=default&mobileredirect=true";
const UPLOAD_FOLDER_SHARE_URL = "https://wesconc.sharepoint.com/:f:/s/Wesco/IgDLZsZf1D8URru5-Ta4tanJAVePK8vUAIcqw9d69AMetqI?e=fWaYhD";
const MAX_FILE_BYTES = 12 * 1024 * 1024;

let workbookPromise;
let uploadFolderPromise;

function shareId(url) {
  return `u!${Buffer.from(url).toString("base64")
    .replace(/=/g, "").replace(/\+/g, "-").replace(/\//g, "_")}`;
}

async function resolveShare(url) {
  return graphFetch(`${GRAPH_BASE}/shares/${shareId(url)}/driveItem` +
    "?$select=id,name,webUrl,parentReference,folder,file");
}

function workbook() {
  if (!workbookPromise) workbookPromise = resolveShare(WORKBOOK_SHARE_URL);
  return workbookPromise;
}

function uploadFolder() {
  if (!uploadFolderPromise) uploadFolderPromise = resolveShare(UPLOAD_FOLDER_SHARE_URL);
  return uploadFolderPromise;
}

function cleanText(value, label, max = 250) {
  const text = String(value || "").trim();
  if (!text || text.length > max) {
    throw Object.assign(new Error(`${label} is required.`), { status: 400 });
  }
  return text;
}

function cleanFile(value, label) {
  if (!value || typeof value !== "object") {
    throw Object.assign(new Error(`${label} is required.`), { status: 400 });
  }
  const fileName = cleanText(value.fileName, `${label} file name`, 180);
  if (/[\/\\]/.test(fileName) || fileName === "." || fileName === "..") {
    throw Object.assign(new Error(`${label} file name is invalid.`), { status: 400 });
  }
  const contentType = cleanText(value.contentType, `${label} file type`, 120);
  const documentBase64 = String(value.documentBase64 || "").trim();
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(documentBase64)) {
    throw Object.assign(new Error(`${label} file contents are invalid.`), { status: 400 });
  }
  const content = Buffer.from(documentBase64, "base64");
  if (!content.length || content.length > MAX_FILE_BYTES) {
    throw Object.assign(new Error(`${label} must be smaller than 12 MB.`), { status: 400 });
  }
  return { fileName, contentType, content };
}

function safeSegment(value) {
  return String(value).replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").slice(0, 180);
}

function excelDate(value) {
  if (typeof value !== "number") return value || "";
  const utc = Date.UTC(1899, 11, 30) + Math.round(value * 86_400_000);
  return new Date(utc).toISOString().slice(0, 10);
}

async function uploadSupportingFile(file, record) {
  const folder = await uploadFolder();
  if (!folder.folder) {
    throw Object.assign(new Error("The SharePoint upload location is not a folder."), { status: 500 });
  }
  const driveId = folder.parentReference?.driveId;
  const prefix = `${safeSegment(record.projectNumber)}_${safeSegment(record.projectName)}`;
  const fileName = `${prefix}_${safeSegment(file.fileName)}`;
  return graphUploadContent(
    `${GRAPH_BASE}/drives/${driveId}/items/${folder.id}:/${encodeURIComponent(fileName)}:/content`,
    file.content,
    file.contentType
  );
}

async function addWorkbookRow(record, contractItem, estimateItem) {
  const item = await workbook();
  const driveId = item.parentReference?.driveId;
  const values = [[
    record.projectName, record.customer, record.projectAddress,
    record.executionDate, record.executedBy, "", record.contactNumber,
    record.estimateNumber, record.invoiceNumber, record.projectNumber,
    contractItem.webUrl, estimateItem.webUrl, record.submittedAt, record.id,
  ]];
  return graphFetch(
    `${GRAPH_BASE}/drives/${driveId}/items/${item.id}/workbook/` +
      `tables('${odataQuote(TABLE_NAME)}')/rows/add`,
    { method: "POST", body: JSON.stringify({ values }) }
  );
}

async function listWorkbookRows() {
  const item = await workbook();
  const driveId = item.parentReference?.driveId;
  const data = await graphFetch(
    `${GRAPH_BASE}/drives/${driveId}/items/${item.id}/workbook/` +
      `tables('${odataQuote(TABLE_NAME)}')/rows?$top=500`
  );
  return (data.value || []).map((row) => {
    const v = row.values?.[0] || [];
    return {
      projectName: v[0] || "", customer: v[1] || "",
      projectAddress: v[2] || "", executionDate: excelDate(v[3]),
      executedBy: v[4] || "", contactNumber: v[6] || "",
      estimateNumber: v[7] || "", invoiceNumber: v[8] || "",
      projectNumber: v[9] || "", contractUrl: v[10] || "",
      estimateUrl: v[11] || "", submittedAt: v[12] || "",
      id: v[13] || `row-${row.index}`,
    };
  }).filter((record) =>
    record.projectName || record.customer || record.projectNumber || record.id.startsWith("row-") === false
  ).reverse();
}

app.http("executedContracts", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "executed-contracts",
  handler: async (request, context) => {
    try {
      if (request.method === "GET") {
        return {
          headers: { "Cache-Control": "no-store" },
          jsonBody: { records: await listWorkbookRows() },
        };
      }
      const body = await request.json();
      const record = {
        id: crypto.randomUUID(),
        projectName: cleanText(body.projectName, "Project Name"),
        customer: cleanText(body.customer, "Customer"),
        contactNumber: cleanText(body.contactNumber, "Contact Number", 50),
        projectAddress: cleanText(body.projectAddress, "Project Address", 500),
        estimateNumber: cleanText(body.estimateNumber, "Estimate Number", 100),
        invoiceNumber: cleanText(body.invoiceNumber, "Invoice Number", 100),
        projectNumber: cleanText(body.projectNumber, "Project Number", 100),
        executionDate: cleanText(body.executionDate, "Date of Execution", 20),
        executedBy: cleanText(body.executedBy, "Executed By", 100),
        submittedAt: new Date().toISOString(),
      };
      const contract = cleanFile(body.contractFile, "Executed Contract");
      const estimate = cleanFile(body.estimateFile, "Approved Estimate");
      const [contractItem, estimateItem] = await Promise.all([
        uploadSupportingFile(contract, record),
        uploadSupportingFile(estimate, record),
      ]);
      await addWorkbookRow(record, contractItem, estimateItem);
      return {
        status: 201,
        jsonBody: {
          record: {
            ...record,
            contractUrl: contractItem.webUrl,
            estimateUrl: estimateItem.webUrl,
          },
        },
      };
    } catch (err) {
      context.error(err);
      return {
        status: err.status || 500,
        jsonBody: {
          error: err.status
            ? err.message
            : "The executed contract could not be saved. Please try again.",
        },
      };
    }
  },
});
