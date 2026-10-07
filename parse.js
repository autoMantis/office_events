/*
 * Pure logic for turning a read-mode email into appointment fields.
 * No Office.js in here so it can be unit-tested with plain node.
 * Loaded as a <script> in the task pane and via require() in test-parse.js.
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.EventParse = factory();
})(typeof self !== "undefined" ? self : this, function () {
  // Teams first: if a Teams join link exists it wins over everything else.
  var TEAMS = [
    /https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s"'<>)\]]+/i,
    /https:\/\/teams\.live\.com\/meet\/[^\s"'<>)\]]+/i,
    /https:\/\/teams\.microsoft\.com\/meet\/[^\s"'<>)\]]+/i
  ];

  var OTHERS = [
    { provider: "Zoom", re: /https:\/\/[\w.-]*zoom\.(?:us|com)\/[jw]\/[^\s"'<>)\]]+/i },
    { provider: "Google Meet", re: /https:\/\/meet\.google\.com\/[a-z-]{10,}[^\s"'<>)\]]*/i },
    { provider: "Webex", re: /https:\/\/[\w.-]*webex\.com\/[^\s"'<>)\]]+/i },
    { provider: "GoTo", re: /https:\/\/[\w.-]*goto(?:meeting)?\.com\/(?:join\/)?[^\s"'<>)\]]+/i },
    { provider: "Whereby", re: /https:\/\/whereby\.com\/[^\s"'<>)\]]+/i },
    { provider: "Meet", re: /https:\/\/[\w.-]*\/(?:meeting|meet)\/[^\s"'<>)\]]+/i }
  ];

  // Trailing punctuation swept up by the URL regexes.
  function tidy(url) {
    return url.replace(/[.,;:!?]+$/, "");
  }

  function findMeeting(text) {
    if (!text) return null;
    for (var i = 0; i < TEAMS.length; i++) {
      var t = TEAMS[i].exec(text);
      if (t) return { provider: "Teams", url: tidy(t[0]) };
    }
    for (var j = 0; j < OTHERS.length; j++) {
      var m = OTHERS[j].re.exec(text);
      if (m) return { provider: OTHERS[j].provider, url: tidy(m[0]) };
    }
    return null;
  }

  // ponytail: no natural-language date parsing. Too many false positives
  // ("call me at 5") for the payoff — you pick the slot in the form anyway.
  // Default: next :00 or :30, 30 minutes long.
  function defaultSlot(now, minutes) {
    var start = new Date(now.getTime());
    start.setSeconds(0, 0);
    start.setMinutes(start.getMinutes() > 30 ? 60 : 30);
    var end = new Date(start.getTime() + (minutes || 30) * 60000);
    return { start: start, end: end };
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function truncate(text, max) {
    var t = String(text || "").replace(/\r\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
    return t.length > max ? t.slice(0, max).trimEnd() + "\n\n[…email truncated]" : t;
  }

  /*
   * email:  { subject, bodyText, from }   from = "Name <addr>" or ""
   * opts:   { now, maxBodyChars, durationMinutes }
   * returns { subject, location, bodyHtml, start, end, meeting }
   */
  function buildEvent(email, opts) {
    opts = opts || {};
    var meeting = findMeeting(email.bodyText) || findMeeting(email.subject);
    var slot = defaultSlot(opts.now || new Date(), opts.durationMinutes);
    var body = truncate(email.bodyText, opts.maxBodyChars || 2000);

    var location = "";
    var header = "";
    if (meeting && meeting.provider === "Teams") {
      location = "Microsoft Teams Meeting";
      header = "Join Microsoft Teams: " + meeting.url;
    } else if (meeting) {
      location = meeting.url;
      header = "Join " + meeting.provider + ": " + meeting.url;
    }

    var parts = [];
    if (header) {
      parts.push(
        '<p><a href="' + escapeHtml(meeting.url) + '">' + escapeHtml(header) + "</a></p><hr>"
      );
    }
    if (email.from) parts.push("<p><b>From:</b> " + escapeHtml(email.from) + "</p>");
    parts.push("<pre style=\"white-space:pre-wrap;font-family:inherit\">" + escapeHtml(body) + "</pre>");

    return {
      subject: email.subject || "(no subject)",
      location: location,
      bodyHtml: parts.join("\n"),
      start: slot.start,
      end: slot.end,
      meeting: meeting
    };
  }

  return {
    findMeeting: findMeeting,
    buildEvent: buildEvent,
    defaultSlot: defaultSlot,
    truncate: truncate
  };
});
