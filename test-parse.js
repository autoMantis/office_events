// node test-parse.js  — fails loud if the link detection / body build breaks.
const assert = require("assert");
const P = require("./parse.js");

// Teams wins even when another link is present.
const mixed = `Hi, join here https://zoom.us/j/123456789?pwd=abc
or Microsoft Teams: https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=%7b%22Tid%22%3a%22x%22%7d`;
assert.strictEqual(P.findMeeting(mixed).provider, "Teams");

assert.strictEqual(P.findMeeting("see https://zoom.us/j/98765432100").provider, "Zoom");
assert.strictEqual(P.findMeeting("https://meet.google.com/abc-defg-hij").provider, "Google Meet");
assert.strictEqual(P.findMeeting("no links at all here"), null);

// Trailing punctuation is stripped.
assert.strictEqual(P.findMeeting("join https://zoom.us/j/111222333.").url, "https://zoom.us/j/111222333");

// Teams => location is the Teams label, link lives in the body.
const teams = P.buildEvent({
  subject: "Q3 review",
  bodyText: "Agenda attached.\nhttps://teams.microsoft.com/l/meetup-join/19%3ameeting_x%40thread.v2/0",
  from: "Ana <ana@example.com>"
});
assert.strictEqual(teams.subject, "Q3 review");
assert.strictEqual(teams.location, "Microsoft Teams Meeting");
assert.ok(teams.bodyHtml.includes("teams.microsoft.com"));
assert.ok(teams.bodyHtml.includes("Ana"));

// Non-Teams => link goes in the location field.
const zoom = P.buildEvent({ subject: "Sync", bodyText: "https://zoom.us/j/555", from: "" });
assert.strictEqual(zoom.location, "https://zoom.us/j/555");

// No link => empty location, body still carried over.
const plain = P.buildEvent({ subject: "Coffee", bodyText: "Cafe Central, 3rd floor", from: "" });
assert.strictEqual(plain.location, "");
assert.ok(plain.bodyHtml.includes("Cafe Central"));

// Long bodies are truncated.
assert.ok(P.truncate("x".repeat(5000), 2000).length < 2100);
assert.ok(P.truncate("x".repeat(5000), 2000).endsWith("[…email truncated]"));
assert.strictEqual(P.truncate("short", 2000), "short");

// HTML in the body cannot break out into the appointment body.
assert.ok(!P.buildEvent({ subject: "s", bodyText: "<script>bad()</script>", from: "" })
  .bodyHtml.includes("<script>"));

// Slot rounds up to the next :00 / :30.
const s1 = P.defaultSlot(new Date("2026-07-28T10:05:00"), 30);
assert.strictEqual(s1.start.getMinutes(), 30);
assert.strictEqual(s1.end.getHours(), 11);
const s2 = P.defaultSlot(new Date("2026-07-28T10:45:00"), 60);
assert.strictEqual(s2.start.getHours(), 11);
assert.strictEqual(s2.start.getMinutes(), 0);
assert.strictEqual(s2.end.getHours(), 12);

console.log("ok");
