/*
 * Meeting Details task pane
 * -------------------------
 * Side panel that lets the organiser set a job number and pick what the
 * meeting is for. Writes both onto the meeting so the on-send check passes.
 *
 * =====================================================================
 *  CONFIG — edit the dropdown names here. Add/remove lines freely.
 * =====================================================================
 */
var PURPOSE_OPTIONS = [
  "Client Billable",
  "Internal",
  "Sales",
  "Support",
  "Admin",
  "Training"
];

// Must match the pattern used in commands.js.
var JOB_NUMBER_REGEX = /\bJ\d{4,6}\b/i;
// Marker line written into the meeting body to record the purpose.
var PURPOSE_PREFIX = "Meeting purpose:";
/* =====================================================================
 *  Logic below — no need to edit.
 * ===================================================================== */

Office.onReady(function () {
  var sel = document.getElementById("purpose");
  PURPOSE_OPTIONS.forEach(function (p) {
    var o = document.createElement("option");
    o.value = p;
    o.textContent = p;
    sel.appendChild(o);
  });
  document.getElementById("apply").addEventListener("click", apply);
  prefill();
});

function val(r) {
  return r.status === Office.AsyncResultStatus.Succeeded && r.value ? r.value : "";
}

// Pre-fill the fields if the meeting already has a job number / purpose.
function prefill() {
  var item = Office.context.mailbox.item;
  item.subject.getAsync(function (sr) {
    var m = val(sr).match(JOB_NUMBER_REGEX);
    if (m) document.getElementById("job").value = m[0];
  });
  item.body.getAsync(Office.CoercionType.Text, function (br) {
    var line = new RegExp(escapeRe(PURPOSE_PREFIX) + "\\s*(.+)", "i").exec(val(br));
    if (line && line[1]) {
      var existing = line[1].trim();
      var sel = document.getElementById("purpose");
      for (var i = 0; i < sel.options.length; i++) {
        if (sel.options[i].value === existing) { sel.selectedIndex = i; break; }
      }
    }
  });
}

function escapeRe(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function setStatus(msg, cls) {
  var s = document.getElementById("status");
  s.textContent = msg;
  s.className = cls || "";
}

function apply() {
  var job = document.getElementById("job").value.trim();
  var purpose = document.getElementById("purpose").value;

  if (!JOB_NUMBER_REGEX.test(job)) {
    setStatus("Enter a valid job number, e.g. J1234.", "err");
    return;
  }
  if (!purpose) {
    setStatus("Choose what this meeting is for.", "err");
    return;
  }

  var item = Office.context.mailbox.item;

  item.subject.getAsync(function (sr) {
    var subject = val(sr);
    var newSubject = JOB_NUMBER_REGEX.test(subject)
      ? subject
      : ("[" + job + "] " + subject).trim();

    item.subject.setAsync(newSubject, function () {
      item.body.getAsync(Office.CoercionType.Text, function (br) {
        var body = val(br);
        var line = PURPOSE_PREFIX + " " + purpose;
        var re = new RegExp("^.*" + escapeRe(PURPOSE_PREFIX) + ".*$", "im");
        var newBody = re.test(body) ? body.replace(re, line) : line + "\n" + body;

        item.body.setAsync(newBody, { coercionType: Office.CoercionType.Text }, function () {
          setStatus("Saved. You can send the meeting now.", "ok");
        });
      });
    });
  });
}
