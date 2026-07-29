const test = require("node:test");
const assert = require("node:assert/strict");
const {
  itemPayload,
  mapRecord,
  recipients,
} = require("../src/functions/punchListDistribution");

test("punch list item payload uses a stable shared key and normalized values", () => {
  const payload = itemPayload(
    {
      punchListKey: "pl_100",
      jobNumber: "24-101",
      jobName: "Wesco Test Project",
      punchDate: "2026-07-28",
      projectManager: "Theo Allen",
      recipients: ["Theo Allen", "Theo Allen", "Wes Atkinson"],
      isSent: true,
      filedAt: "2026-07-28T18:00:00Z",
      updatedAt: "2026-07-28T19:00:00Z",
    },
    {
      itemKey: "pl_item_1",
      description: "Repair trim",
      location: "North elevation",
      urgent: true,
      completed: false,
    }
  );

  assert.equal(payload.new_distributionitemkey, "pl_100:pl_item_1");
  assert.equal(payload.new_punchlistkey, "pl_100");
  assert.equal(payload.new_description, "Repair trim");
  assert.equal(payload.new_urgent, true);
  assert.equal(payload.new_completed, false);
  assert.equal(payload.new_issent, true);
  assert.deepEqual(JSON.parse(payload.new_recipients), [
    "Theo Allen",
    "Wes Atkinson",
  ]);
});

test("Dataverse records map back to Distribution Log rows", () => {
  const row = mapRecord({
    new_punchlistdistributionitemid: "row-id",
    new_distributionitemkey: "pl_100:pl_item_1",
    new_punchlistkey: "pl_100",
    new_itemkey: "pl_item_1",
    new_jobnumber: "24-101",
    new_jobname: "Wesco Test Project",
    new_punchdate: "2026-07-28T00:00:00Z",
    new_recipients: '["Theo Allen"]',
    new_description: "Repair trim",
    new_urgent: true,
    new_completed: true,
    new_issent: true,
    modifiedon: "2026-07-28T19:00:00Z",
  });

  assert.equal(row.punchDate, "2026-07-28");
  assert.deepEqual(row.recipients, ["Theo Allen"]);
  assert.equal(row.completed, true);
  assert.equal(row.isSent, true);
});

test("recipient names are de-duplicated and blank names are removed", () => {
  assert.deepEqual(recipients(["Theo Allen", "", "Theo Allen", "Wes Atkinson"]), [
    "Theo Allen",
    "Wes Atkinson",
  ]);
});
