const { app } = require("@azure/functions");
const { graphFetch } = require("../graphClient");
const {
  sendAgreementForSignature,
  getEnvelopeStatus,
  verifyDocuSignConnection,
} = require("../docuSignClient");

const SITE_HOSTNAME = "wesconc.sharepoint.com";
const SITE_PATH = "/sites/Wesco";
const SUBCONTRACTOR_ROOT = "We App/Subcontractors";
const SIGNATURE_FOLDER = "Review & Signature Needed";

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
  if (!documentBase64 || documentBase64.length > 20_000_000) {
    throw Object.assign(
      new Error("A valid agreement document is required for Docusign."),
      { status: 400 }
    );
  }
  return documentBase64;
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
      const destination =
        `${SUBCONTRACTOR_ROOT}/${matched.name}/${SIGNATURE_FOLDER}/${fileName}`;
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
          destination:
            `${SUBCONTRACTOR_ROOT}/${matched.name}/${SIGNATURE_FOLDER}`,
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
    try {
      if (request.method === "GET") {
        if (request.query.get("health") === "1") {
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
      const documentBase64 = cleanBase64(body.documentBase64);
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

      const destination =
        `${SUBCONTRACTOR_ROOT}/${matched.name}/${SIGNATURE_FOLDER}`;
      const result = await sendAgreementForSignature({
        documentBase64,
        fileName,
        projectName,
        subcontractorFolder: destination,
        subcontractorName,
        subcontractorEmail,
        contractorName,
        contractorEmail,
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
