const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const {
  buildEnvelopeDefinition,
  docusignErrorDetails,
  documentExtension,
  initialsForEveryPage,
  sectionTenTabs,
} = require("../src/docuSignClient");

test("envelope contains the agreement and PO as separate ordered documents", () => {
  const envelope = buildEnvelopeDefinition({
    documentBase64: "PGh0bWw+PC9odG1sPg==",
    fileName: "Wesco Agreement.html",
    purchaseOrder: {
      fileName: "TEST-DS-001.pdf",
      contentType: "application/pdf",
      documentBase64: "JVBERi0xLjQ=",
    },
    projectName: "DOCUSIGN TEST ONLY",
    subcontractorFolder: "Sent for Signature",
    subcontractorName: "Oossie Mae",
    subcontractorEmail: "tantaneiarountree@gmail.com",
    contractorName: "Theo Allen",
    contractorEmail: "tallen@wesconc.com",
    completionFolder: "Others",
    completionWebhookUrl: "https://example.com/completed",
  });

  assert.deepEqual(
    envelope.documents.map((document) => ({
      id: document.documentId,
      extension: document.fileExtension,
      name: document.name,
    })),
    [
      { id: "1", extension: "html", name: "Wesco Agreement.html" },
      { id: "2", extension: "pdf", name: "TEST-DS-001.pdf" },
    ]
  );
  assert.deepEqual(
    envelope.recipients.signers.map((signer) => ({
      name: signer.name,
      order: signer.routingOrder,
    })),
    [
      { name: "Theo Allen", order: "1" },
      { name: "Oossie Mae", order: "2" },
    ]
  );
  assert.match(envelope.eventNotification.url, /folder=Others/);
  assert.equal(envelope.status, "created");
});

test("purchase order file extensions are normalized for Docusign", () => {
  assert.equal(documentExtension("purchase-order.PDF", "application/pdf"), "pdf");
  assert.equal(documentExtension("photo.jpeg", "image/jpeg"), "jpg");
  assert.equal(documentExtension("scan", "image/png"), "png");
  assert.equal(documentExtension("malware.exe", "application/octet-stream"), "");
});

test("both recipients receive required initials on every agreement page", () => {
  const wesco = initialsForEveryPage(3, "contractor");
  const subcontractor = initialsForEveryPage(3, "subcontractor");

  assert.equal(wesco.length, 3);
  assert.equal(subcontractor.length, 3);
  assert.deepEqual(wesco.map((tab) => tab.pageNumber), ["1", "2", "3"]);
  assert.deepEqual(subcontractor.map((tab) => tab.pageNumber), ["1", "2", "3"]);
  assert.ok(wesco.every((tab) => tab.documentId === "1"));
  assert.ok(wesco.every((tab) => tab.recipientId === "1"));
  assert.ok(subcontractor.every((tab) => tab.recipientId === "2"));
  assert.ok([...wesco, ...subcontractor].every((tab) => tab.optional === "false"));
});

test("Section 10 assigns Wesco left-side fields and subcontractor right-side fields", () => {
  const wesco = sectionTenTabs("contractor");
  const subcontractor = sectionTenTabs(
    "subcontractor",
    "tantaneiarountree@gmail.com"
  );

  assert.deepEqual(
    wesco.signHereTabs.map((tab) => tab.anchorString),
    ["WESCO_SIGNATURE_ANCHOR"]
  );
  assert.deepEqual(
    subcontractor.signHereTabs.map((tab) => tab.anchorString),
    ["SUB_SIGNATURE_ANCHOR", "EQUIP_SIGNATURE_ANCHOR"]
  );
  const subcontractorEmail = subcontractor.textTabs.find(
    (tab) => tab.anchorString === "SUB_EMAIL_ANCHOR"
  );
  assert.equal(subcontractorEmail.value, "tantaneiarountree@gmail.com");
  assert.equal(subcontractorEmail.locked, "false");
});

test("Docusign API error body is preserved for a useful diagnosis", () => {
  const details = docusignErrorDetails({
    message: "Request failed with status code 400",
    response: {
      statusCode: 400,
      body: {
        errorCode: "DOCUMENT_CONVERSION_ERROR",
        message: "The document could not be converted.",
      },
    },
  });

  assert.deepEqual(details, {
    status: 400,
    code: "DOCUMENT_CONVERSION_ERROR",
    message: "The document could not be converted.",
  });
});

test("browser payload sends the PO separately and removes the nested PDF preview", () => {
  const agreementPath = path.resolve(
    __dirname,
    "..",
    "..",
    "Wesco Subcontractor Agreement.html"
  );
  const source = fs.readFileSync(agreementPath, "utf8");

  assert.match(source, /purchaseOrder:\s*\{/);
  assert.match(
    source,
    /Purchase Order attached as a separate Docusign document:/
  );
  assert.match(source, /poPreview\.innerHTML = ''/);
});
