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

function anchorTab(anchorString, recipientId, tabLabel, options = {}) {
  return {
    documentId: "1",
    recipientId,
    anchorString,
    anchorUnits: "pixels",
    anchorXOffset: String(options.xOffset ?? 0),
    anchorYOffset: String(options.yOffset ?? 0),
    anchorIgnoreIfNotPresent: "false",
    tabLabel,
    optional: options.optional ? "true" : "false",
    ...(options.width ? { width: String(options.width) } : {}),
    ...(options.height ? { height: String(options.height) } : {}),
    ...(options.value ? { value: options.value } : {}),
    ...(options.locked !== undefined
      ? { locked: options.locked ? "true" : "false" }
      : {}),
  };
}

function documentExtension(fileName, contentType) {
  const extension = String(fileName || "")
    .split(".")
    .pop()
    .toLowerCase();
  if (["pdf", "png", "jpg", "jpeg"].includes(extension)) {
    return extension === "jpeg" ? "jpg" : extension;
  }
  const byType = {
    "application/pdf": "pdf",
    "image/png": "png",
    "image/jpeg": "jpg",
  };
  return byType[String(contentType || "").toLowerCase()] || "";
}

function docusignErrorDetails(error) {
  let responseBody =
    error?.response?.body ||
    error?.response?.data ||
    error?.body ||
    {};
  if (typeof responseBody === "string") {
    try {
      responseBody = JSON.parse(responseBody);
    } catch {
      responseBody = { message: responseBody };
    }
  }
  const code =
    responseBody.errorCode ||
    responseBody.error ||
    error?.code ||
    "DOCUSIGN_REQUEST_FAILED";
  const message =
    responseBody.message ||
    responseBody.error_description ||
    error?.message ||
    "Docusign rejected the request.";
  const status =
    Number(error?.response?.statusCode || error?.response?.status || error?.status) ||
    502;
  return {
    status,
    code: String(code).slice(0, 120),
    message: String(message).slice(0, 700),
  };
}

async function runDocusignStage(stage, operation) {
  try {
    return await operation();
  } catch (error) {
    const details = docusignErrorDetails(error);
    throw Object.assign(
      new Error(`${stage}: ${details.message}`),
      {
        status: details.status,
        code: details.code,
        stage,
      }
    );
  }
}

function sectionTenTabs(recipient, subcontractorEmail) {
  const isContractor = recipient === "contractor";
  const recipientId = isContractor ? "1" : "2";
  if (isContractor) {
    return {
      signHereTabs: [
        anchorTab(
          "WESCO_SIGNATURE_ANCHOR",
          recipientId,
          "Wesco Signature",
          { yOffset: 4 }
        ),
      ],
      textTabs: [
        anchorTab("WESCO_NAME_ANCHOR", recipientId, "Wesco Name", {
          width: 220,
          height: 22,
        }),
        anchorTab("WESCO_TITLE_ANCHOR", recipientId, "Wesco Title", {
          width: 220,
          height: 22,
        }),
      ],
      dateSignedTabs: [
        anchorTab("WESCO_DATE_ANCHOR", recipientId, "Wesco Date Signed", {
          width: 110,
          height: 22,
        }),
      ],
    };
  }

  return {
    signHereTabs: [
      anchorTab(
        "SUB_SIGNATURE_ANCHOR",
        recipientId,
        "Subcontractor Signature",
        { yOffset: 4 }
      ),
      anchorTab(
        "EQUIP_SIGNATURE_ANCHOR",
        recipientId,
        "Equipment Agreement Signature",
        { yOffset: 4 }
      ),
    ],
    textTabs: [
      anchorTab(
        "SUB_COMPANY_ANCHOR",
        recipientId,
        "Subcontractor Company",
        { width: 220, height: 22 }
      ),
      anchorTab("SUB_NAME_ANCHOR", recipientId, "Subcontractor Name", {
        width: 220,
        height: 22,
      }),
      anchorTab("SUB_TITLE_ANCHOR", recipientId, "Subcontractor Title", {
        width: 220,
        height: 22,
        optional: true,
      }),
      anchorTab("SUB_ADDRESS_ANCHOR", recipientId, "Subcontractor Address", {
        width: 220,
        height: 42,
      }),
      anchorTab("SUB_PHONE_ANCHOR", recipientId, "Subcontractor Telephone", {
        width: 220,
        height: 22,
      }),
      anchorTab("SUB_FAX_ANCHOR", recipientId, "Subcontractor Facsimile", {
        width: 220,
        height: 22,
        optional: true,
      }),
      anchorTab("SUB_EMAIL_ANCHOR", recipientId, "Subcontractor Email", {
        width: 220,
        height: 22,
        value: subcontractorEmail,
        locked: false,
      }),
      anchorTab(
        "SUB_CONTACT_ANCHOR",
        recipientId,
        "Subcontractor Project Contact",
        { width: 220, height: 22, optional: true }
      ),
      anchorTab("EQUIP_NAME_ANCHOR", recipientId, "Equipment Employee Name", {
        width: 220,
        height: 22,
      }),
      anchorTab(
        "EQUIP_COMPANY_ANCHOR",
        recipientId,
        "Equipment Company Name",
        { width: 220, height: 22 }
      ),
    ],
    dateSignedTabs: [
      anchorTab(
        "SUB_DATE_ANCHOR",
        recipientId,
        "Subcontractor Date Signed",
        { width: 110, height: 22 }
      ),
      anchorTab(
        "EQUIP_DATE_ANCHOR",
        recipientId,
        "Equipment Agreement Date Signed",
        { width: 110, height: 22 }
      ),
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

function buildEnvelopeDefinition({
  documentBase64,
  fileName,
  purchaseOrder,
  projectName,
  subcontractorFolder,
  subcontractorName,
  subcontractorEmail,
  contractorName,
  contractorEmail,
  completionFolder,
  completionWebhookUrl,
}) {
  // DocuSign limits envelope text custom-field values to 100 characters.
  // Keep the full document name on the document and completion webhook, while
  // constraining only the hidden audit metadata sent as envelope custom fields.
  const customFieldValue = (value) => String(value || "").slice(0, 100);
  const completedFileName = fileName.replace(/\.html$/i, ".pdf");
  const documents = [
    {
      documentBase64,
      name: fileName,
      fileExtension: "html",
      documentId: "1",
    },
  ];
  if (purchaseOrder) {
    const extension = documentExtension(
      purchaseOrder.fileName,
      purchaseOrder.contentType
    );
    if (!extension) {
      throw Object.assign(
        new Error("The purchase order must be a PDF, PNG, JPG, or JPEG file."),
        { status: 400, code: "INVALID_PURCHASE_ORDER" }
      );
    }
    documents.push({
      documentBase64: purchaseOrder.documentBase64,
      name: purchaseOrder.fileName,
      fileExtension: extension,
      documentId: "2",
    });
  }

  return {
    emailSubject: `Wesco Subcontractor Agreement - ${projectName}`,
    emailBlurb:
      "Wesco Exteriors completes and signs first. After Wesco finishes, Docusign automatically emails the subcontractor to complete the right-side information, initials, and signatures.",
    documents,
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
          value: customFieldValue(subcontractorFolder),
          show: "false",
          required: "false",
        },
        {
          name: "Wesco Source File",
          value: customFieldValue(fileName),
          show: "false",
          required: "false",
        },
      ],
    },
    eventNotification: {
      url:
        `${completionWebhookUrl}?fileName=` +
        encodeURIComponent(completedFileName) +
        "&folder=" +
        encodeURIComponent(completionFolder),
      loggingEnabled: "true",
      requireAcknowledgment: "true",
      useSoapInterface: "false",
      includeDocuments: "false",
      includeCertificateOfCompletion: "false",
      eventData: {
        version: "restv2.1",
        format: "json",
        includeData: [],
      },
      envelopeEvents: [
        {
          envelopeEventStatusCode: "completed",
          includeDocuments: "false",
        },
      ],
    },
    status: "created",
  };
}

async function sendAgreementForSignature(options) {
  const envelopesApi = await getEnvelopesApi();
  const accountId = requireSetting("DOCUSIGN_ACCOUNT_ID");
  const completionWebhookUrl =
    process.env.DOCUSIGN_COMPLETION_WEBHOOK_URL ||
    "https://wonderful-field-08b34ac0f.7.azurestaticapps.net/api/subcontractor-agreement-docusign-completed";
  const envelopeDefinition = buildEnvelopeDefinition({
    ...options,
    completionWebhookUrl,
  });
  const { subcontractorEmail } = options;

  const created = await runDocusignStage(
    "Create the Docusign draft envelope",
    () =>
      envelopesApi.createEnvelope(accountId, {
        envelopeDefinition,
      })
  );
  const envelopeId = created.envelopeId;
  try {
    const pageCount = await runDocusignStage(
      "Convert the agreement into signing pages",
      () => getConvertedPageCount(envelopesApi, envelopeId)
    );

    for (const recipient of ["contractor", "subcontractor"]) {
      const recipientId = recipient === "contractor" ? "1" : "2";
      await runDocusignStage(
        recipient === "contractor"
          ? "Place the Wesco signature and initials"
          : "Place the subcontractor signature and initials",
        () =>
          envelopesApi.createTabs(accountId, envelopeId, recipientId, {
            tabs: {
              initialHereTabs: initialsForEveryPage(pageCount, recipient),
              ...sectionTenTabs(recipient, subcontractorEmail),
            },
          })
      );
    }

    await runDocusignStage("Send the Docusign envelope", () =>
      envelopesApi.update(accountId, envelopeId, {
        envelope: { status: "sent" },
      })
    );

    return {
      envelopeId,
      pageCount,
      status: "sent",
    };
  } catch (error) {
    error.envelopeId = envelopeId;
    throw error;
  }
}

async function downloadCompletedEnvelope(envelopeId) {
  const envelopesApi = await getEnvelopesApi();
  const accountId = requireSetting("DOCUSIGN_ACCOUNT_ID");
  const envelope = await envelopesApi.getEnvelope(accountId, envelopeId);
  if (String(envelope.status || "").toLowerCase() !== "completed") {
    throw Object.assign(
      new Error("The Docusign envelope is not completed yet."),
      { status: 409 }
    );
  }

  const document = await envelopesApi.getDocument(
    accountId,
    envelopeId,
    "combined",
    { certificate: "true" }
  );
  if (Buffer.isBuffer(document)) return document;
  if (document instanceof ArrayBuffer) return Buffer.from(document);
  if (Buffer.isBuffer(document?.data)) return document.data;
  if (document?.data instanceof ArrayBuffer) return Buffer.from(document.data);
  if (typeof document === "string") return Buffer.from(document, "binary");
  throw new Error("Docusign returned an unsupported completed-document format.");
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
  downloadCompletedEnvelope,
  getEnvelopeStatus,
  verifyDocuSignConnection,
  buildEnvelopeDefinition,
  docusignErrorDetails,
  documentExtension,
  sectionTenTabs,
  initialsForEveryPage,
};
