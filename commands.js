/*
 * Job Number Enforcer for Outlook
 * ---------------------------------
 * Blocks a meeting/appointment from being sent unless a valid job number
 * is present. Runs on the OnMessageSend and OnAppointmentSend "Smart Alert"
 * events. Paired with SendMode="Block" in the manifest, this is a HARD block:
 * the item cannot leave Drafts until a job number is supplied.
 *
 * =====================================================================
 *  CONFIG — change these to match your job-number format, then re-host.
 * =====================================================================
 */

// Pattern a valid job number must match.
// Default: the letter "J" followed by 4 to 6 digits (e.g. J1234, J123456),
// case-insensitive, as a whole word. Edit this one line to change the rule.
//   Examples:
//     /\bJ\d{4,6}\b/i          -> J1234
//     /\bJOB[- ]?\d{4,6}\b/i   -> JOB-1234 or JOB 1234
//     /\b\d{6}\b/              -> any 6-digit number
var JOB_NUMBER_REGEX = /\bJ\d{4,6}\b/i;

// Also accept the job number in the meeting body/notes (not just the subject)?
var ALSO_CHECK_BODY = true;

// Message shown to the user when the job number is missing.
var ERROR_MESSAGE =
  "This meeting needs a job number in the subject before it can be sent " +
  "(for example: J1234). Add it to the subject line and send again.";

/* =====================================================================
 *  Logic below — no need to edit.
 * ===================================================================== */

function block(event) {
  event.completed({
    allowEvent: false,
    errorMessage: ERROR_MESSAGE
  });
}

function allow(event) {
  event.completed({ allowEvent: true });
}

function validateAndComplete(event) {
  var item = Office.context.mailbox.item;

  item.subject.getAsync(function (subjectResult) {
    var subject =
      subjectResult.status === Office.AsyncResultStatus.Succeeded && subjectResult.value
        ? subjectResult.value
        : "";

    if (JOB_NUMBER_REGEX.test(subject)) {
      allow(event);
      return;
    }

    if (!ALSO_CHECK_BODY) {
      block(event);
      return;
    }

    item.body.getAsync(Office.CoercionType.Text, function (bodyResult) {
      var body =
        bodyResult.status === Office.AsyncResultStatus.Succeeded && bodyResult.value
          ? bodyResult.value
          : "";
      if (JOB_NUMBER_REGEX.test(body)) {
        allow(event);
      } else {
        block(event);
      }
    });
  });
}

// Meetings sent to attendees fire OnMessageSend; personal appointments fire
// OnAppointmentSend. Both run the same check.
function onMessageSendHandler(event) {
  validateAndComplete(event);
}

function onAppointmentSendHandler(event) {
  validateAndComplete(event);
}

// Required so new Outlook (Windows) and Outlook on the web can find the
// handlers by name. Harmless on classic Outlook.
if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
}
