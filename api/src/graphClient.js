const { ConfidentialClientApplication } = require("@azure/msal-node");

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

// Each entry is a separate Excel workbook this API can read/write, sharing the
// same app-only Graph credentials (granted Sites.Selected on the whole site).
const WORKBOOKS = {
  equipment: { driveId: process.env.EXCEL_DRIVE_ID, itemId: process.env.EXCEL_ITEM_ID },
  trailer: { driveId: process.env.TRAILER_EXCEL_DRIVE_ID, itemId: process.env.TRAILER_EXCEL_ITEM_ID },
  vehicle: { driveId: process.env.VEHICLE_EXCEL_DRIVE_ID, itemId: process.env.VEHICLE_EXCEL_ITEM_ID },
  hardware: { driveId: process.env.HARDWARE_EXCEL_DRIVE_ID, itemId: process.env.HARDWARE_EXCEL_ITEM_ID },
};

function workbookBase(workbookKey) {
  const cfg = WORKBOOKS[workbookKey];
  if (!cfg || !cfg.driveId || !cfg.itemId) {
    throw new Error(`Missing drive/item ID app settings for workbook '${workbookKey}'`);
  }
  return `${GRAPH_BASE}/drives/${cfg.driveId}/items/${cfg.itemId}/workbook`;
}

let msalClient = null;
function getMsalClient() {
  if (!msalClient) {
    msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: process.env.EQUIPMENT_GRAPH_CLIENT_ID || process.env.GRAPH_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${
          process.env.EQUIPMENT_GRAPH_TENANT_ID || process.env.GRAPH_TENANT_ID
        }`,
        clientSecret:
          process.env.EQUIPMENT_GRAPH_CLIENT_SECRET || process.env.GRAPH_CLIENT_SECRET,
      },
    });
  }
  return msalClient;
}

let cachedToken = null; // { token, expiresOn }
async function getAccessToken() {
  if (cachedToken && cachedToken.expiresOn > Date.now() + 60_000) {
    return cachedToken.token;
  }
  const result = await getMsalClient().acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"],
  });
  cachedToken = { token: result.accessToken, expiresOn: new Date(result.expiresOn).getTime() };
  return cachedToken.token;
}

async function rawFetch(url, options = {}) {
  const token = await getAccessToken();
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    const err = new Error(`Graph ${options.method || "GET"} ${url} failed: ${res.status} ${text}`);
    err.status = res.status;
    throw err;
  }
  if (res.status === 204) return null;
  return res.json();
}

async function graphUploadContent(url, content, contentType = "application/octet-stream") {
  const token = await getAccessToken();
  const res = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": contentType,
    },
    body: content,
  });
  if (!res.ok) {
    const text = await res.text().catch(() => "");
    const err = new Error(`Graph PUT ${url} failed: ${res.status} ${text}`);
    err.status = res.status;
    throw err;
  }
  return res.json();
}

// Original signature: relative path against the "equipment" workbook.
// Kept as-is so equipment.js/visits.js don't need to change.
async function graphFetch(path, options = {}) {
  const url = path.startsWith("http") ? path : `${workbookBase("equipment")}${path}`;
  return rawFetch(url, options);
}

// New signature for additional workbooks (e.g. "trailer").
async function graphFetchFor(workbookKey, path, options = {}) {
  const url = path.startsWith("http") ? path : `${workbookBase(workbookKey)}${path}`;
  return rawFetch(url, options);
}

// Must exactly match the sanitize() logic used in restructure_excel.py so
// table names computed here line up with the tables already created in the workbook.
function sanitizeForTableName(name) {
  let s = name.trim().replace(/[^A-Za-z0-9_]/g, "_");
  if (!s || !/^[A-Za-z]/.test(s)) s = "T_" + s;
  return s.slice(0, 60);
}

function visitsTableName(sheetName) {
  return "Visits_" + sanitizeForTableName(sheetName);
}

// Excel worksheet names: max 31 chars, cannot contain : \ / ? * [ ]
function sanitizeForSheetName(name) {
  let s = name.trim().replace(/[:\\/?*\[\]]/g, "-");
  return s.slice(0, 31) || "Equipment";
}

module.exports = {
  graphFetch,
  graphFetchFor,
  graphUploadContent,
  workbookBase,
  visitsTableName,
  sanitizeForTableName,
  sanitizeForSheetName,
};
