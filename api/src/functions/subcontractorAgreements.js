const { app } = require("@azure/functions");
const { graphFetch, graphUploadContent } = require("../graphClient");
const {
  sendAgreementForSignature,
  downloadCompletedEnvelope,
  getEnvelopeStatus,
  verifyDocuSignConnection,
} = require("../docuSignClient");

const SITE_HOSTNAME = "wesconc.sharepoint.com";
const SITE_PATH = "/sites/Wesco";
const SUBCONTRACTOR_ROOT = "We App/Subcontractors";
const AGREEMENT_ARCHIVE_ROOT =
  "We App/Internal Staff/Rebecca Kirkman/Saved Subcontractors Agreements";
const SENT_FOR_SIGNATURE_FOLDER =
  `${AGREEMENT_ARCHIVE_ROOT}/Sent for Signature`;

let siteIdPromise = null;

function encodeGraphPath(path) {
  return path.split("/").map(encodeURIComponent).join("/");
}

function cleanName(value, label) {
  const name = String(value || "").trim();
  if (
    !name ||
    name.length > 180 ||
    /[\/\\]/.test(name) ||
    name === "." ||
    name === ".."
  ) {
    throw Object.assign(new Error(`A valid ${label} is required.`), {
      status: 400,
    });
  }
  return name;
}

function cleanEmail(value, label) {
  const email = String(value || "").trim();
  if (
    !email ||
    email.length > 254 ||
    !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)
  ) {
    throw Object.assign(new Error(`A valid ${label} is required.`), {
      status: 400,
    });
  }
  return email;
}

function cleanBase64(value) {
  const documentBase64 = String(value || "").trim();
  if (
    !documentBase64 ||
    documentBase64.length > 20_000_000 ||
    documentBase64.length % 4 !== 0 ||
    !/^[A-Za-z0-9+/]+={0,2}$/.test(documentBase64)
  ) {
    throw Object.assign(
      new Error("A valid agreement document is required for Docusign."),
      { status: 400 }
    );
  }
  return documentBase64;
}

function cleanPurchaseOrder(value) {
  if (!value || typeof value !== "object") {
    throw Object.assign(
      new Error("A valid purchase order attachment is required for Docusign."),
      { status: 400 }
    );
  }
  const fileName = cleanName(value.fileName, "purchase order file name");
  const contentType = String(value.contentType || "").trim().toLowerCase();
  const allowedTypes = new Set([
    "application/pdf",
    "image/png",
    "image/jpeg",
  ]);
  if (!allowedTypes.has(contentType)) {
    throw Object.assign(
      new Error("The purchase order must be a PDF, PNG, JPG, or JPEG file."),
      { status: 400 }
    );
  }
  const extension = fileName.split(".").pop().toLowerCase();
  const validExtensions = {
    "application/pdf": new Set(["pdf"]),
    "image/png": new Set(["png"]),
    "image/jpeg": new Set(["jpg", "jpeg"]),
  };
  if (!validExtensions[contentType].has(extension)) {
    throw Object.assign(
      new Error("The purchase order file extension does not match its content type."),
      { status: 400 }
    );
  }
  const documentBase64 = cleanBase64(value.documentBase64);
  const header = Buffer.from(documentBase64.slice(0, 24), "base64");
  const validHeader =
    (contentType === "application/pdf" &&
      header.subarray(0, 5).toString("ascii") === "%PDF-") ||
    (contentType === "image/png" &&
      header.subarray(0, 8).equals(
        Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
      )) ||
    (contentType === "image/jpeg" &&
      header[0] === 0xff &&
      header[1] === 0xd8 &&
      header[2] === 0xff);
  if (!validHeader) {
    throw Object.assign(
      new Error("The purchase order contents do not match the selected file type."),
      { status: 400 }
    );
  }
  return {
    fileName,
    contentType,
    documentBase64,
  };
}

async function getSiteId() {
  if (!siteIdPromise) {
    siteIdPromise = graphFetch(
      `https://graph.microsoft.com/v1.0/sites/${SITE_HOSTNAME}:${SITE_PATH}`
    ).then((site) => site.id);
  }
  return siteIdPromise;
}

async function listSubcontractors() {
  const siteId = await getSiteId();
  const data = await graphFetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/` +
      `${encodeGraphPath(SUBCONTRACTOR_ROOT)}:/children` +
      "?$select=id,name,folder&$top=200"
  );
  return (data.value || [])
    .filter((item) => item.folder)
    .map((item) => ({ id: item.id, name: item.name }))
    .sort((a, b) => a.name.localeCompare(b.name));
}

app.http("subcontractorAgreements", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "subcontractor-agreements",
  handler: async (request, context) => {
    try {
      const subcontractors = await listSubcontractors();
      if (request.method === "GET") {
        return {
          headers: { "Cache-Control": "no-store" },
          jsonBody: { subcontractors },
        };
      }

      const body = await request.json();
      const requestedName = cleanName(body.subcontractor, "subcontractor");
      const fileName = cleanName(body.fileName, "file name");
      if (!fileName.toLowerCase().endsWith(".html")) {
        throw Object.assign(new Error("The agreement must be an HTML file."), {
          status: 400,
        });
      }

      const matched = subcontractors.find(
        (item) => item.name.toLowerCase() === requestedName.toLowerCase()
      );
      if (!matched) {
        throw Object.assign(
          new Error("That subcontractor folder does not exist in SharePoint."),
          { status: 404 }
        );
      }

      const siteId = await getSiteId();
      const destination = `${SENT_FOR_SIGNATURE_FOLDER}/${fileName}`;
      const session = await graphFetch(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/` +
          `${encodeGraphPath(destination)}:/createUploadSession`,
        {
          method: "POST",
          body: JSON.stringify({
            item: {
              "@microsoft.graph.conflictBehavior": "rename",
              name: fileName,
            },
          }),
        }
      );
      return {
        jsonBody: {
          uploadUrl: session.uploadUrl,
          destination: SENT_FOR_SIGNATURE_FOLDER,
        },
      };
    } catch (err) {
      context.error(err);
      const isHealthCheck =
        request.method === "GET" && request.query.get("health") === "1";
      return {
        status: err.status || 500,
        jsonBody: {
          error: err.status
            ? err.message
            : "The subcontractor agreement service is temporarily unavailable.",
        },
      };
    }
  },
});

app.http("subcontractorAgreementDocusign", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "subcontractor-agreement-docusign",
  handler: async (request, context) => {
    const isHealthCheck =
      request.method === "GET" && request.query.get("health") === "1";
    try {
      if (request.method === "GET") {
        if (isHealthCheck) {
          return {
            headers: { "Cache-Control": "no-store" },
            jsonBody: await verifyDocuSignConnection(),
          };
        }
        const envelopeId = cleanName(
          request.query.get("envelopeId"),
          "Docusign envelope ID"
        );
        return {
          headers: { "Cache-Control": "no-store" },
          jsonBody: await getEnvelopeStatus(envelopeId),
        };
      }

      const body = await request.json();
      const requestedName = cleanName(body.subcontractor, "subcontractor");
      const fileName = cleanName(body.fileName, "file name");
      const subcontractorName = cleanName(
        body.subcontractorName || requestedName,
        "subcontractor signer name"
      );
      const contractorName = cleanName(
        body.contractorName,
        "Wesco signer name"
      );
      const subcontractorEmail = cleanEmail(
        body.subcontractorEmail,
        "subcontractor email"
      );
      const contractorEmail = cleanEmail(
        body.contractorEmail,
        "Wesco signer email"
      );
      if (contractorEmail.toLowerCase() === subcontractorEmail.toLowerCase()) {
        throw Object.assign(
          new Error(
            "The subcontractor email must be different from the Wesco signer email."
          ),
          { status: 400 }
        );
      }
      const documentBase64 = cleanBase64(body.documentBase64);
      const purchaseOrder = cleanPurchaseOrder(body.purchaseOrder);
      const projectName = cleanName(body.projectName, "project name");

      if (!fileName.toLowerCase().endsWith(".html")) {
        throw Object.assign(new Error("The agreement must be an HTML file."), {
          status: 400,
        });
      }

      const subcontractors = await listSubcontractors();
      const matched = subcontractors.find(
        (item) => item.name.toLowerCase() === requestedName.toLowerCase()
      );
      if (!matched) {
        throw Object.assign(
          new Error("That subcontractor folder does not exist in SharePoint."),
          { status: 404 }
        );
      }

      const destination = SENT_FOR_SIGNATURE_FOLDER;
      const result = await sendAgreementForSignature({
        documentBase64,
        fileName,
        purchaseOrder,
        projectName,
        subcontractorFolder: destination,
        subcontractorName,
        subcontractorEmail,
        contractorName,
        contractorEmail,
        completionFolder: matched.name,
      });

      return {
        jsonBody: {
          ...result,
          destination,
        },
      };
    } catch (err) {
      context.error(err);
      return {
        status: err.status || 500,
        jsonBody: {
          error: err.status
            ? err.message
            : "Docusign could not send this agreement. Please try again.",
          code: err.code,
          stage: err.stage,
          consentUrl: err.consentUrl,
          diagnostic: isHealthCheck
            ? `${err.name || "Error"}: ${err.message || "Unknown error"}`.slice(
                0,
                700
              )
            : undefined,
        },
      };
    }
  },
});

app.http("subcontractorAgreementDocusignCompleted", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "subcontractor-agreement-docusign-completed",
  handler: async (request, context) => {
    try {
      const body = await request.json();
      const eventName = String(body.event || body.Event || "").toLowerCase();
      if (eventName && eventName !== "envelope-completed") {
        return { status: 202, jsonBody: { accepted: true } };
      }

      const envelopeId = cleanName(
        body.data?.envelopeId ||
          body.envelopeId ||
          body.EnvelopeStatus?.EnvelopeID,
        "Docusign envelope ID"
      );
      const fileName = cleanName(
        request.query.get("fileName") ||
          `Wesco-Subcontractor-Agreement-${envelopeId}.pdf`,
        "completed agreement file name"
      );
      if (!fileName.toLowerCase().endsWith(".pdf")) {
        throw Object.assign(
          new Error("The completed agreement must be a PDF file."),
          { status: 400 }
        );
      }

      const requestedFolder = cleanName(
        request.query.get("folder"),
        "subcontractor completion folder"
      );
      const subcontractors = await listSubcontractors();
      const matched = subcontractors.find(
        (item) => item.name.toLowerCase() === requestedFolder.toLowerCase()
      );
      if (!matched) {
        throw Object.assign(
          new Error(
            "The subcontractor completion folder does not exist in SharePoint."
          ),
          { status: 404 }
        );
      }

      const completedPdf = await downloadCompletedEnvelope(envelopeId);
      const siteId = await getSiteId();
      const destinationFolder = `${SUBCONTRACTOR_ROOT}/${matched.name}`;
      const destination = `${destinationFolder}/${fileName}`;
      const uploaded = await graphUploadContent(
        `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/` +
          `${encodeGraphPath(destination)}:/content`,
        completedPdf,
        "application/pdf"
      );

      return {
        jsonBody: {
          archived: true,
          envelopeId,
          destination: destinationFolder,
          itemId: uploaded.id,
        },
      };
    } catch (err) {
      context.error(err);
      return {
        status: err.status || 500,
        jsonBody: {
          error: err.status
            ? err.message
            : "The completed Docusign agreement could not be archived.",
        },
      };
    }
  },
});
