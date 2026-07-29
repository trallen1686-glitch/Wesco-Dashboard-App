const { app } = require("@azure/functions");
const { dataverseJson } = require("../dataverseClient");

const ENTITY_SET = "new_vendorsuppliers";
const META_PREFIX = "WESCO_VENDOR_META:";
const SELECT = [
  "new_vendorsupplierid", "new_name", "new_vendorid", "new_vendorname",
  "new_accountnumber", "new_vendorcategory", "new_primarycontactname",
  "new_contacttitle", "new_officephone", "new_cellphone", "new_emailaddress",
  "new_website", "new_physicaladdress", "new_city", "new_state", "new_zipcode",
  "new_paymentterms", "new_creditlimit", "new_salesrepresentative",
  "new_representativephone", "new_representativeemail", "new_deliveryfee",
  "new_deliverynotes", "new_productcategories", "new_pricinginformation",
  "new_specialdiscounts", "new_manufacturerprograms", "new_freightpolicies",
  "new_returnpolicies", "new_vendorcomments", "createdon", "modifiedon"
].join(",");

function clean(value, max = 4000) {
  return String(value ?? "").trim().slice(0, max);
}

function numberOrNull(value) {
  if (value === "" || value === null || value === undefined) return null;
  const number = Number(value);
  return Number.isFinite(number) && number >= 0 ? number : null;
}

function validId(value) {
  return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(String(value || ""));
}

function metaFrom(record) {
  const value = String(record.new_vendorcomments || "");
  if (!value.startsWith(META_PREFIX)) return { vendorComments: value };
  try { return JSON.parse(value.slice(META_PREFIX.length)); } catch { return {}; }
}

function mapRecord(record) {
  const meta = metaFrom(record);
  return {
    recordKey: record.new_vendorsupplierid,
    vendorId: record.new_vendorid || "",
    vendorName: record.new_vendorname || record.new_name || "",
    accountNumber: record.new_accountnumber || "",
    vendorCategory: record.new_vendorcategory || "",
    primaryContact: record.new_primarycontactname || "",
    contactTitle: record.new_contacttitle || "",
    officePhone: record.new_officephone || "",
    cellPhone: record.new_cellphone || "",
    email: record.new_emailaddress || "",
    website: record.new_website || "",
    physicalAddress: record.new_physicaladdress || "",
    city: record.new_city || "",
    state: record.new_state || "",
    zipCode: record.new_zipcode || "",
    paymentTerms: record.new_paymentterms || "",
    creditLimit: record.new_creditlimit ?? "",
    salesRepresentative: record.new_salesrepresentative || "",
    representativePhone: record.new_representativephone || "",
    representativeEmail: record.new_representativeemail || "",
    deliveryFee: record.new_deliveryfee ?? "",
    deliveryNotes: record.new_deliverynotes || "",
    productCategories: record.new_productcategories || "",
    pricingInformation: record.new_pricinginformation || "",
    specialDiscounts: record.new_specialdiscounts || "",
    manufacturerPrograms: record.new_manufacturerprograms || "",
    freightPolicies: record.new_freightpolicies || "",
    returnPolicies: record.new_returnpolicies || "",
    ...meta
  };
}

function payloadFrom(body) {
  const meta = {
    taxExemptAccepted: clean(body.taxExemptAccepted, 3),
    preferredVendor: clean(body.preferredVendor, 3),
    approvedVendor: clean(body.approvedVendor, 3),
    deliveryAvailable: clean(body.deliveryAvailable, 3),
    leadTimeDays: clean(body.leadTimeDays, 10),
    emergencyDelivery: clean(body.emergencyDelivery, 3),
    pickupAvailable: clean(body.pickupAvailable, 3),
    lastPurchaseDate: clean(body.lastPurchaseDate, 10),
    ytdPurchases: clean(body.ytdPurchases, 30),
    pricingRating: clean(body.pricingRating, 1),
    qualityRating: clean(body.qualityRating, 1),
    deliveryRating: clean(body.deliveryRating, 1),
    serviceRating: clean(body.serviceRating, 1),
    status: clean(body.status, 20),
    vendorComments: clean(body.vendorComments, 1500)
  };
  const payload = {
    new_name: clean(body.vendorName, 850),
    new_vendorid: clean(body.vendorId, 100),
    new_vendorname: clean(body.vendorName, 850),
    new_accountnumber: clean(body.accountNumber, 100),
    new_vendorcategory: clean(body.vendorCategory, 100),
    new_primarycontactname: clean(body.primaryContact, 200),
    new_contacttitle: clean(body.contactTitle, 100),
    new_officephone: clean(body.officePhone, 50),
    new_cellphone: clean(body.cellPhone, 50),
    new_emailaddress: clean(body.email, 320),
    new_website: clean(body.website, 500),
    new_physicaladdress: clean(body.physicalAddress, 500),
    new_city: clean(body.city, 100),
    new_state: clean(body.state, 2),
    new_zipcode: clean(body.zipCode, 20),
    new_paymentterms: clean(body.paymentTerms, 100),
    new_creditlimit: numberOrNull(body.creditLimit),
    new_salesrepresentative: clean(body.salesRepresentative, 200),
    new_representativephone: clean(body.representativePhone, 50),
    new_representativeemail: clean(body.representativeEmail, 320),
    new_deliveryfee: numberOrNull(body.deliveryFee),
    new_deliverynotes: clean(body.deliveryNotes, 1500),
    new_productcategories: clean(body.productCategories, 500),
    new_pricinginformation: clean(body.pricingInformation, 1500),
    new_specialdiscounts: clean(body.specialDiscounts, 1500),
    new_manufacturerprograms: clean(body.manufacturerPrograms, 1500),
    new_freightpolicies: clean(body.freightPolicies, 1500),
    new_returnpolicies: clean(body.returnPolicies, 1500),
    new_vendorcomments: META_PREFIX + JSON.stringify(meta)
  };
  for (const [key, value] of Object.entries(payload)) {
    if (value === "") delete payload[key];
  }
  if (payload.new_creditlimit === null) delete payload.new_creditlimit;
  if (payload.new_deliveryfee === null) delete payload.new_deliveryfee;
  return payload;
}

function validate(body) {
  return clean(body.vendorId, 100) && clean(body.vendorName, 850) && clean(body.vendorCategory, 100);
}

function safeFailure(error) {
  const message = String(error && error.message || "");
  if (/privilege|accesscheck|seclib|permission|not authorized/i.test(message)) return "DATAVERSE_PERMISSION";
  if (/property|attribute|payload|undeclared/i.test(message)) return "DATAVERSE_SCHEMA";
  if (/currency|transactioncurrency/i.test(message)) return "DATAVERSE_CURRENCY";
  return "DATAVERSE_WRITE";
}

function safeFailureField(error) {
  const message = String(error && error.message || "");
  const matches = message.match(/new_[a-z0-9_]+/ig) || [];
  return [...new Set(matches.map(value => value.toLowerCase()))].slice(0, 5);
}

app.http("vendors", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  route: "vendors",
  handler: async (request, context) => {
    try {
      if (request.method === "GET") {
        const result = await dataverseJson(`${ENTITY_SET}?$select=${SELECT}&$orderby=new_vendorname asc`);
        return { jsonBody: { items: (result.value || []).map(mapRecord) } };
      }
      const body = await request.json();
      if (!validate(body)) return { status: 400, jsonBody: { error: "Vendor ID, name, and category are required." } };
      const created = await dataverseJson(`${ENTITY_SET}?$select=${SELECT}`, {
        method: "POST", headers: { Prefer: "return=representation" },
        body: JSON.stringify(payloadFrom(body))
      });
      return { status: 201, jsonBody: { item: mapRecord(created) } };
    } catch (error) {
      context.error(error);
      return { status: 500, jsonBody: { error: "The vendor database is temporarily unavailable.", code: safeFailure(error), field: safeFailureField(error) } };
    }
  }
});

app.http("vendorRecord", {
  methods: ["PATCH", "DELETE"],
  authLevel: "anonymous",
  route: "vendors/{id}",
  handler: async (request, context) => {
    const id = request.params.id;
    if (!validId(id)) return { status: 400, jsonBody: { error: "Invalid vendor ID." } };
    try {
      if (request.method === "DELETE") {
        await dataverseJson(`${ENTITY_SET}(${id})`, { method: "DELETE" });
        return { status: 204 };
      }
      const body = await request.json();
      if (!validate(body)) return { status: 400, jsonBody: { error: "Vendor ID, name, and category are required." } };
      await dataverseJson(`${ENTITY_SET}(${id})`, { method: "PATCH", body: JSON.stringify(payloadFrom(body)) });
      const updated = await dataverseJson(`${ENTITY_SET}(${id})?$select=${SELECT}`);
      return { jsonBody: { item: mapRecord(updated) } };
    } catch (error) {
      context.error(error);
      return { status: 500, jsonBody: { error: "The vendor record could not be updated.", code: safeFailure(error), field: safeFailureField(error) } };
    }
  }
});
