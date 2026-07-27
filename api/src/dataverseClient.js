const { ConfidentialClientApplication } = require("@azure/msal-node");

let msalClient = null;
let cachedToken = null;

function dataverseUrl() {
  const value = String(process.env.DATAVERSE_URL || "").replace(/\/+$/, "");
  if (!value) throw new Error("DATAVERSE_URL is not configured.");
  return value;
}

function getMsalClient() {
  if (!msalClient) {
    msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: process.env.DATAVERSE_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.DATAVERSE_TENANT_ID}`,
        clientSecret: process.env.DATAVERSE_CLIENT_SECRET,
      },
    });
  }
  return msalClient;
}

async function getAccessToken() {
  if (cachedToken && cachedToken.expiresOn > Date.now() + 60_000) {
    return cachedToken.token;
  }
  const result = await getMsalClient().acquireTokenByClientCredential({
    scopes: [`${dataverseUrl()}/.default`],
  });
  cachedToken = {
    token: result.accessToken,
    expiresOn: new Date(result.expiresOn).getTime(),
  };
  return cachedToken.token;
}

async function dataverseFetch(path, options = {}) {
  const token = await getAccessToken();
  const response = await fetch(
    path.startsWith("http")
      ? path
      : `${dataverseUrl()}/api/data/v9.2/${path.replace(/^\/+/, "")}`,
    {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Content-Type": "application/json",
        ...(options.headers || {}),
      },
    }
  );

  if (!response.ok) {
    const text = await response.text().catch(() => "");
    const error = new Error(
      `Dataverse ${options.method || "GET"} failed: ${response.status} ${text.slice(0, 500)}`
    );
    error.status = response.status;
    throw error;
  }
  return response;
}

async function dataverseJson(path, options = {}) {
  const response = await dataverseFetch(path, options);
  if (response.status === 204) return null;
  return response.json();
}

module.exports = { dataverseFetch, dataverseJson };
