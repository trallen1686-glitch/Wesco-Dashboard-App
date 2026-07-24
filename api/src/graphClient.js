const { ConfidentialClientApplication } = require("@azure/msal-node");

const DRIVE_ID = process.env.EXCEL_DRIVE_ID;
const ITEM_ID = process.env.EXCEL_ITEM_ID;
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
const WORKBOOK_BASE = `${GRAPH_BASE}/drives/${DRIVE_ID}/items/${ITEM_ID}/workbook`;

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

async function graphFetch(path, options = {}) {
  const token = await getAccessToken();
  const url = path.startsWith("http") ? path : `${WORKBOOK_BASE}${path}`;
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
    const err = new Error(`Graph ${options.method || "GET"} ${path} failed: ${res.status} ${text}`);
    err.status = res.status;
    throw err;
  }
  if (res.status === 204) return null;
  return res.json();
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
  visitsTableName,
  sanitizeForTableName,
  sanitizeForSheetName,
  WORKBOOK_BASE,
};
