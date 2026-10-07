/* global Office, EventParse */

var state = { bodyText: "", from: "" };

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Outlook) return;
  load();
  document.getElementById("create").addEventListener("click", create);
  document.getElementById("duration").addEventListener("change", refresh);
});

function localInputValue(d) {
  var pad = function (n) { return String(n).padStart(2, "0"); };
  return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()) +
         "T" + pad(d.getHours()) + ":" + pad(d.getMinutes());
}

function load() {
  var item = Office.context.mailbox.item;
  state.from = item.from ? (item.from.displayName + " <" + item.from.emailAddress + ">") : "";

  item.body.getAsync(Office.CoercionType.Text, function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Could not read the email body: " + result.error.message);
      state.bodyText = "";
    } else {
      state.bodyText = result.value;
    }
    var ev = EventParse.buildEvent(
      { subject: item.subject, bodyText: state.bodyText, from: state.from },
      { durationMinutes: duration() }
    );
    document.getElementById("subject").value = ev.subject;
    document.getElementById("location").value = ev.location;
    document.getElementById("start").value = localInputValue(ev.start);
    refresh();
    document.getElementById("create").disabled = false;
  });
}

function duration() {
  return parseInt(document.getElementById("duration").value, 10);
}

// Rebuild the preview/badge from current body text (subject + location stay user-edited).
function refresh() {
  var ev = EventParse.buildEvent(
    { subject: document.getElementById("subject").value, bodyText: state.bodyText, from: state.from },
    { durationMinutes: duration() }
  );
  var badge = document.getElementById("meeting");
  if (ev.meeting) {
    badge.className = "badge found";
    badge.textContent = ev.meeting.provider + " link found";
  } else {
    badge.className = "badge none";
    badge.textContent = "No meeting link found";
  }
  document.getElementById("preview").textContent =
    EventParse.truncate(state.bodyText, 2000).slice(0, 600);
  return ev;
}

function create() {
  var ev = refresh();
  var start = new Date(document.getElementById("start").value);
  if (isNaN(start.getTime())) return setStatus("Pick a valid start time.");
  var end = new Date(start.getTime() + duration() * 60000);

  setStatus("");
  Office.context.mailbox.displayNewAppointmentForm({
    subject: document.getElementById("subject").value,
    location: document.getElementById("location").value,
    body: ev.bodyHtml,
    start: start,
    end: end,
    requiredAttendees: [] // ponytail: you add attendees in the form, per the spec
  });
}

function setStatus(msg) {
  document.getElementById("status").textContent = msg;
}
