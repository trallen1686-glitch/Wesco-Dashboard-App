const docusign = require("docusign-esign");
const { DefaultAzureCredential } = require("@azure/identity");
const { SecretClient } = require("@azure/keyvault-secrets");

const REQUIRED_SETTINGS = [
  "DOCUSIGN_INTEGRATION_KEY",
  "DOCUSIGN_USER_ID",
  "DOCUSIGN_ACCOUNT_ID",
  "DOCUSIGN_KEY_VAULT_URL",
];

let secretClient = null;
let cachedPrivateKey = null;
let cachedToken = null;

function requireSetting(name) {
  const value = String(process.env[name] || "").trim();
  if (!value) {
    throw Object.assign(new Error(`Missing server setting: ${name}`), {
      status: 503,
    });
  }
  return value;
}

function assertConfigured() {
  REQUIRED_SETTINGS.forEach(requireSetting);
}

function getSecretClient() {
  if (!secretClient) {
    secretClient = new SecretClient(
      requireSetting("DOCUSIGN_KEY_VAULT_URL"),
      new DefaultAzureCredential()
    );
  }
  return secretClient;
}

async function getPrivateKey() {
  if (process.env.DOCUSIGN_PRIVATE_KEY) {
    const configuredKey = process.env.DOCUSIGN_PRIVATE_KEY.trim();
    if (/^[A-Za-z0-9+/=]+$/.test(configuredKey)) {
      const decodedKey = Buffer.from(configuredKey, "base64").toString("utf8");
      if (decodedKey.includes("PRIVATE KEY")) return decodedKey;
    }
    return configuredKey.replace(/\\n/g, "\n");
  }
  if (!cachedPrivateKey) {
    const secretName =
      process.env.DOCUSIGN_PRIVATE_KEY_SECRET || "Docusign-PrivateKey";
    const secret = await getSecretClient().getSecret(secretName);
    if (!secret.value) {
      throw Object.assign(
        new Error(`Azure Key Vault secret ${secretName} is empty.`),
        { status: 503 }
      );
    }
    cachedPrivateKey = secret.value;
  }
  return cachedPrivateKey;
}

function createApiClient() {
  const apiClient = new docusign.ApiClient();
  apiClient.setOAuthBasePath(
    process.env.DOCUSIGN_OAUTH_HOST || "account-d.docusign.com"
  );
  return apiClient;
}

async function getAccessToken() {
  if (cachedToken && cachedToken.expiresAt > Date.now() + 120_000) {
    return cachedToken.value;
  }

  assertConfigured();
  const apiClient = createApiClient();
  try {
    const result = await apiClient.requestJWTUserToken(
      requireSetting("DOCUSIGN_INTEGRATION_KEY"),
      requireSetting("DOCUSIGN_USER_ID"),
      ["signature", "impersonation"],
      Buffer.from(await getPrivateKey(), "utf8"),
      3600
    );
    cachedToken = {
      value: result.body.access_token,
      expiresAt: Date.now() + Number(result.body.expires_in || 3600) * 1000,
    };
    return cachedToken.value;
  } catch (error) {
    const responseBody = error?.response?.body || error?.response?.data;
    const code = responseBody?.error || error?.body?.error;
    if (code === "consent_required") {
      const redirectUri =
        process.env.DOCUSIGN_CONSENT_REDIRECT_URI ||
        "https://wonderful-field-08b34ac0f.7.azurestaticapps.net/";
      throw Object.assign(
        new Error(
          "Docusign administrator consent is required before agreements can be sent."
        ),
        {
          status: 409,
          code: "DOCUSIGN_CONSENT_REQUIRED",
          consentUrl: apiClient.getJWTUri(
            requireSetting("DOCUSIGN_INTEGRATION_KEY"),
            redirectUri,
            process.env.DOCUSIGN_OAUTH_HOST || "account-d.docusign.com"
          ),
        }
      );
    }
    throw error;
  }
}

async function getEnvelopesApi() {
  const apiClient = createApiClient();
  apiClient.setBasePath(
    process.env.DOCUSIGN_BASE_PATH ||
      "https://demo.docusign.net/restapi"
  );
  apiClient.addDefaultHeader(
    "Authorization",
    `Bearer ${await getAccessToken()}`
  );
  return new docusign.EnvelopesApi(apiClient);
}

function initialsForEveryPage(pageCount, recipient) {
  const isContractor = recipient === "contractor";
  return Array.from({ length: pageCount }, (_, index) => ({
    documentId: "1",
    pageNumber: String(index + 1),
    recipientId: isContractor ? "1" : "2",
    xPosition: isContractor ? "32" : "540",
    yPosition: "742",
    width: "24",
    height: "20",
    tabLabel: `${isContractor ? "Wesco" : "Subcontractor"} Initial ${
      index + 1
    }`,
    optional: "false",
  }));
}

function finalPageTabs(pageCount, recipient) {
  const isContractor = recipient === "contractor";
  const recipientId = isContractor ? "1" : "2";
  const xPosition = isContractor ? "82" : "355";
  return {
    signHereTabs: [
      {
        documentId: "1",
        pageNumber: String(pageCount),
        recipientId,
        xPosition,
        yPosition: "640",
        tabLabel: `${isContractor ? "Wesco" : "Subcontractor"} Signature`,
        optional: "false",
      },
    ],
    dateSignedTabs: [
      {
        documentId: "1",
        pageNumber: String(pageCount),
        recipientId,
        xPosition,
        yPosition: "700",
        width: "88",
        height: "16",
        tabLabel: `${isContractor ? "Wesco" : "Subcontractor"} Date Signed`,
      },
    ],
    fullNameTabs: [
      {
        documentId: "1",
        pageNumber: String(pageCount),
        recipientId,
        xPosition,
        yPosition: "675",
        width: "170",
        height: "16",
        tabLabel: `${isContractor ? "Wesco" : "Subcontractor"} Full Name`,
      },
    ],
  };
}

async function getConvertedPageCount(envelopesApi, envelopeId) {
  let lastError = null;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    try {
      const result = await envelopesApi.getPages(
        requireSetting("DOCUSIGN_ACCOUNT_ID"),
        envelopeId,
        "1",
        { count: "100", dpi: "72", maxWidth: "612", maxHeight: "792" }
      );
      if (Array.isArray(result.pages) && result.pages.length) {
        return result.pages.length;
      }
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 900));
  }
  if (lastError) throw lastError;
  throw new Error("Docusign did not finish converting the agreement pages.");
}

async function sendAgreementForSignature({
  documentBase64,
  fileName,
  projectName,
  subcontractorFolder,
  subcontractorName,
  subcontractorEmail,
  contractorName,
  contractorEmail,
}) {
  const envelopesApi = await getEnvelopesApi();
  const accountId = requireSetting("DOCUSIGN_ACCOUNT_ID");
  const envelopeDefinition = {
    emailSubject: `Wesco Subcontractor Agreement - ${projectName}`,
    emailBlurb:
      "Please review and complete the required initials and signature fields in this Wesco subcontractor agreement.",
    documents: [
      {
        documentBase64,
        name: fileName,
        fileExtension: "html",
        documentId: "1",
      },
    ],
    recipients: {
      signers: [
        {
          email: contractorEmail,
          name: contractorName,
          recipientId: "1",
          routingOrder: "1",
        },
        {
          email: subcontractorEmail,
          name: subcontractorName,
          recipientId: "2",
          routingOrder: "2",
        },
      ],
    },
    customFields: {
      textCustomFields: [
        {
          name: "Wesco SharePoint Folder",
          value: subcontractorFolder,
          show: "false",
          required: "false",
        },
        {
          name: "Wesco Source File",
          value: fileName,
          show: "false",
          required: "false",
        },
      ],
    },
    status: "created",
  };

  const created = await envelopesApi.createEnvelope(accountId, {
    envelopeDefinition,
  });
  const envelopeId = created.envelopeId;
  const pageCount = await getConvertedPageCount(envelopesApi, envelopeId);

  for (const recipient of ["contractor", "subcontractor"]) {
    const recipientId = recipient === "contractor" ? "1" : "2";
    const finalTabs = finalPageTabs(pageCount, recipient);
    await envelopesApi.createTabs(accountId, envelopeId, recipientId, {
      tabs: {
        initialHereTabs: initialsForEveryPage(pageCount, recipient),
        ...finalTabs,
      },
    });
  }

  await envelopesApi.update(accountId, envelopeId, {
    envelope: { status: "sent" },
  });

  return {
    envelopeId,
    pageCount,
    status: "sent",
  };
}

async function getEnvelopeStatus(envelopeId) {
  const envelopesApi = await getEnvelopesApi();
  const envelope = await envelopesApi.getEnvelope(
    requireSetting("DOCUSIGN_ACCOUNT_ID"),
    envelopeId
  );
  return {
    envelopeId: envelope.envelopeId,
    status: envelope.status,
    statusChangedDateTime: envelope.statusChangedDateTime,
    completedDateTime: envelope.completedDateTime,
  };
}

async function verifyDocuSignConnection() {
  await getAccessToken();
  return {
    ready: true,
    environment:
      process.env.DOCUSIGN_BASE_PATH ||
      "https://demo.docusign.net/restapi",
  };
}

module.exports = {
  sendAgreementForSignature,
  getEnvelopeStatus,
  verifyDocuSignConnection,
};
