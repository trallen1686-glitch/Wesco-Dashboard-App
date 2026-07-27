const { ConfidentialClientApplication } = require("@azure/msal-node");

let msalClient = null;
let cachedToken = null;

function dataverseUrl() {
  const value = String(process.env.DATAVERSE_URL || "").replace(/\/+$/, "");
  if (!value) throw new Error("DATAVERSE_URL is not configured.");
  return value;
}

function credentialCandidates(profile = "default") {
  const defaultCandidates = [
    {
      clientId: process.env.DATAVERSE_CLIENT_ID,
      tenantId: process.env.DATAVERSE_TENANT_ID,
      clientSecret: process.env.DATAVERSE_CLIENT_SECRET,
    },
    {
      clientId: process.env.EQUIPMENT_GRAPH_CLIENT_ID,
      tenantId: process.env.EQUIPMENT_GRAPH_TENANT_ID,
      clientSecret: process.env.EQUIPMENT_GRAPH_CLIENT_SECRET,
    },
    {
      clientId: process.env.GRAPH_CLIENT_ID,
      tenantId: process.env.GRAPH_TENANT_ID,
      clientSecret: process.env.GRAPH_CLIENT_SECRET,
    },
  ];
  const equipment = defaultCandidates[1];
  const candidates = profile === "equipment"
    ? [equipment, ...defaultCandidates.filter((candidate) => candidate !== equipment)]
    : defaultCandidates;
  const seen = new Set();
  return candidates.filter(({ clientId, tenantId, clientSecret }) => {
    if (!clientId || !tenantId || !clientSecret) return false;
    const key = `${clientId}|${tenantId}|${clientSecret}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

async function getAccessToken(profile = "default") {
  if (cachedToken && cachedToken.expiresOn > Date.now() + 60_000) {
    return cachedToken.token;
  }
  let lastError = null;
  for (const credential of credentialCandidates(profile)) {
    try {
      const client = new ConfidentialClientApplication({
        auth: {
          clientId: credential.clientId,
          authority: `https://login.microsoftonline.com/${credential.tenantId}`,
          clientSecret: credential.clientSecret,
        },
      });
      const result = await client.acquireTokenByClientCredential({
        scopes: [`${dataverseUrl()}/.default`],
      });
      msalClient = client;
      cachedToken = {
        token: result.accessToken,
        expiresOn: new Date(result.expiresOn).getTime(),
      };
      return cachedToken.token;
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("No Dataverse application credential is configured.");
}

async function dataverseFetch(path, options = {}) {
  const profile = options.credentialProfile || "default";
  const { credentialProfile, ...fetchOptions } = options;
  const token = await getAccessToken(profile);
  const response = await fetch(
    path.startsWith("http")
      ? path
      : `${dataverseUrl()}/api/data/v9.2/${path.replace(/^\/+/, "")}`,
    {
      ...fetchOptions,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0",
        "Content-Type": "application/json",
        ...(fetchOptions.headers || {}),
      },
    }
  );

  if (!response.ok) {
    const text = await response.text().catch(() => "");
    const error = new Error(
      `Dataverse ${fetchOptions.method || "GET"} failed: ${response.status} ${text.slice(0, 500)}`
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
