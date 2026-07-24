/*
 * Job Number Enforcer for Outlook
 * ---------------------------------
 * Blocks a meeting/appointment from being sent unless BOTH are present:
 *   1. a valid job number (in the subject or notes), and
 *   2. a meeting purpose (set via the "Meeting details" side panel).
 * Runs on the OnMessageSend and OnAppointmentSend "Smart Alert" events.
 * With SendMode="Block" in the manifest this is a HARD block: the item
 * cannot leave Drafts until both are supplied.
 *
 * =====================================================================
 *  CONFIG — change these to match your rules, then re-upload this file.
 * =====================================================================
 */

// What a valid job number looks like. Default: "J" + 4 to 6 digits (e.g. J1234).
var JOB_NUMBER_REGEX = /\bJ\d{4,6}\b/i;

// Also accept the job number in the meeting body/notes (not just the subject)?
var ALSO_CHECK_BODY = true;

// Require a meeting purpose too? Set to false to enforce only the job number.
var REQUIRE_PURPOSE = true;

// The marker the side panel writes into the body to record the purpose.
var PURPOSE_REGEX = /Meeting purpose:\s*\S+/i;

// Messages shown when something's missing.
var MSG_JOB =
  "This meeting needs a job number (e.g. J1234). Open the 'Meeting details' " +
  "button on the ribbon to add it, then send again.";
var MSG_PURPOSE =
  "Please set what this meeting is for. Open the 'Meeting details' button on " +
  "the ribbon, pick a purpose, then send again.";
var MSG_BOTH =
  "This meeting needs a job number and a purpose. Open the 'Meeting details' " +
  "button on the ribbon, fill both in, then send again.";

/* =====================================================================
 *  Logic below — no need to edit.
 * ===================================================================== */

function val(r) {
  return r.status === Office.AsyncResultStatus.Succeeded && r.value ? r.value : "";
}

function validateAndComplete(event) {
  var item = Office.context.mailbox.item;

  item.subject.getAsync(function (sr) {
    var subject = val(sr);

    item.body.getAsync(Office.CoercionType.Text, function (br) {
      var body = val(br);

      var jobOk =
        JOB_NUMBER_REGEX.test(subject) ||
        (ALSO_CHECK_BODY && JOB_NUMBER_REGEX.test(body));
      var purposeOk = !REQUIRE_PURPOSE || PURPOSE_REGEX.test(body);

      if (jobOk && purposeOk) {
        event.completed({ allowEvent: true });
        return;
      }

      var msg = !jobOk && !purposeOk ? MSG_BOTH : !jobOk ? MSG_JOB : MSG_PURPOSE;
      event.completed({ allowEvent: false, errorMessage: msg });
    });
  });
}

function onMessageSendHandler(event) {
  validateAndComplete(event);
}

function onAppointmentSendHandler(event) {
  validateAndComplete(event);
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
}
