const test = require("node:test");
const assert = require("node:assert/strict");

const {
  buildReminderEmail,
  classifyReminderItems,
  isActiveReminderStatus,
  isCompletedStatus,
  isFirstMonday,
  isOnHoldStatus,
  isOverdue,
  resolveRecipientOverride,
  runEnquiryReminderJob,
} = require("../api/_lib/enquiry-reminders");

test("classifies active, completed, lost, and on-hold statuses", () => {
  assert.equal(isActiveReminderStatus("7. Initial Enquiry"), true);
  assert.equal(isActiveReminderStatus("5. Arranging Assessment"), true);
  assert.equal(isActiveReminderStatus("Won"), false);
  assert.equal(isActiveReminderStatus("Lost - Pre Assessment"), false);
  assert.equal(isActiveReminderStatus("Didn't Enquire"), false);
  assert.equal(isActiveReminderStatus("Not Qualified"), false);
  assert.equal(isActiveReminderStatus("On Hold - awaiting family"), false);

  assert.equal(isCompletedStatus("Won"), true);
  assert.equal(isCompletedStatus("Lost - Post Assessment"), true);
  assert.equal(isOnHoldStatus("Client On Hold"), true);
});

test("detects overdue enquiries using the Modified timestamp and seven-day cutoff", () => {
  const now = new Date("2026-05-18T07:00:00.000Z");

  assert.equal(isOverdue("2026-05-10T06:59:59.000Z", now), true);
  assert.equal(isOverdue("2026-05-11T07:00:00.000Z", now), false);
  assert.equal(isOverdue("not a date", now), false);
});

test("includes on-hold reminders only on the first Monday of the month", () => {
  assert.equal(isFirstMonday(new Date("2026-06-01T07:00:00.000Z")), true);
  assert.equal(isFirstMonday(new Date("2026-06-08T07:00:00.000Z")), false);
  assert.equal(isFirstMonday(new Date("2026-06-02T07:00:00.000Z")), false);
});

test("classifies overdue active and on-hold items separately", () => {
  const now = new Date("2026-06-01T07:00:00.000Z");
  const items = [
    {
      title: "Active overdue",
      status: "7. Initial Enquiry",
      modified: "2026-05-20T09:00:00.000Z",
      modifiedTime: Date.parse("2026-05-20T09:00:00.000Z"),
    },
    {
      title: "On hold overdue",
      status: "On Hold",
      modified: "2026-05-19T09:00:00.000Z",
      modifiedTime: Date.parse("2026-05-19T09:00:00.000Z"),
    },
    {
      title: "Recently updated",
      status: "7. Initial Enquiry",
      modified: "2026-05-31T09:00:00.000Z",
      modifiedTime: Date.parse("2026-05-31T09:00:00.000Z"),
    },
    {
      title: "Completed",
      status: "Won",
      modified: "2026-05-19T09:00:00.000Z",
      modifiedTime: Date.parse("2026-05-19T09:00:00.000Z"),
    },
  ];

  assert.deepEqual(
    classifyReminderItems(items, { now, includeOnHold: true }).active.map((item) => item.title),
    ["Active overdue"]
  );
  assert.deepEqual(
    classifyReminderItems(items, { now, includeOnHold: true }).onHold.map((item) => item.title),
    ["On hold overdue"]
  );
  assert.deepEqual(classifyReminderItems(items, { now, includeOnHold: false }).onHold, []);
});

test("builds grouped reminder email content with last-updated details", () => {
  const email = buildReminderEmail(
    {
      active: [
        {
          title: "Anne Example",
          ownerName: "Chris",
          status: "7. Initial Enquiry",
          modified: "2026-05-20T09:00:00.000Z",
          followUp: "Call daughter",
          detailsExcerpt: "Left voicemail.",
          webUrl: "https://example.test/item/1",
        },
      ],
      onHold: [
        {
          title: "Bob Example",
          ownerName: "Chris",
          status: "On Hold",
          modified: "2026-05-19T09:00:00.000Z",
          followUp: "",
          detailsExcerpt: "",
          webUrl: "",
        },
      ],
    },
    {
      now: new Date("2026-06-01T07:00:00.000Z"),
      includeOnHold: true,
      listUrl: "https://example.test/list",
    }
  );

  assert.match(email.subject, /1 active, 1 on hold/);
  assert.match(email.html, /Modified/);
  assert.match(email.html, /Active enquiries overdue this week/);
  assert.match(email.html, /On-hold enquiries overdue this month/);
  assert.match(email.html, /Chris/);
  assert.match(email.html, /Anne Example/);
  assert.match(email.html, /Last updated:/);
  assert.match(email.html, /Call daughter/);
  assert.match(email.html, /Left voicemail/);
});

test("uses recipient override for trial delivery", () => {
  assert.deepEqual(
    resolveRecipientOverride({
      ENQUIRY_REMINDER_RECIPIENT_OVERRIDE: "chris@planwithcare.co.uk; Chris@planwithcare.co.uk, ops@example.test",
    }),
    ["chris@planwithcare.co.uk", "ops@example.test"]
  );
  assert.deepEqual(resolveRecipientOverride({}), []);
});

test("rejects malformed recipient override before Graph send", () => {
  assert.throws(
    () =>
      resolveRecipientOverride({
        ENQUIRY_REMINDER_RECIPIENT_OVERRIDE: "chris@planwithcare",
      }),
    /invalid email address/i
  );
});

test("dry-run job builds email payload without sending Graph mail", async () => {
  let sent = false;
  const result = await runEnquiryReminderJob({
    dryRun: true,
    now: new Date("2026-06-01T07:00:00.000Z"),
    env: {
      ENQUIRY_REMINDER_FROM_EMAIL: "sender@example.test",
      ENQUIRY_REMINDER_RECIPIENT_OVERRIDE: "chris@planwithcare.co.uk",
    },
    graphClient: {},
    listUrl: "https://example.test/list",
    items: [
      {
        title: "Anne Example",
        ownerName: "Chris",
        status: "7. Initial Enquiry",
        modified: "2026-05-20T09:00:00.000Z",
        modifiedTime: Date.parse("2026-05-20T09:00:00.000Z"),
        followUp: "Call daughter",
        detailsExcerpt: "Needs a check-in.",
        webUrl: "https://example.test/item/1",
      },
    ],
    sendMail: async () => {
      sent = true;
    },
  });

  assert.equal(sent, false);
  assert.equal(result.dryRun, true);
  assert.equal(result.activeCount, 1);
  assert.equal(result.onHoldCount, 0);
  assert.deepEqual(result.recipients, ["chris@planwithcare.co.uk"]);
  assert.match(result.email.html, /Anne Example/);
});
