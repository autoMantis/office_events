# Disclaimer — read this before you install anything

This is a personal side project, shared so people can learn from it. It is not a product,
it is not supported, and nobody is on call for it.

## No warranty, no liability

The software is provided **"as is"**, with no warranty of any kind, as stated in the
[LICENSE](LICENSE). The author is **not liable** for any damage, data loss, missed meeting,
wrong calendar entry, policy violation or security incident that follows from using,
modifying or deploying this code. You install it, you run it, you own the outcome.

## Where it is fine to use this

- Your own personal mailbox, on your own machine.
- A test or throwaway Microsoft 365 account.
- As a reference you read, learn from, and rewrite yourself.

## Where it is NOT fine to use this

- **Do not install it on your company mailbox** or on any account that holds work email,
  client data, patient data, legal or financial information.
- **Do not deploy it for your team, your employer or a client** on the strength of this
  README. Most organisations require review and approval before any add-in touches a
  mailbox, and installing one without that can breach your IT policy or your contract.
- Do not treat it as compliant with GDPR, HIPAA or any other framework. No such assessment
  has been done.

## Before it goes anywhere near real work

1. **Read the code first.** It is about 200 lines across `taskpane.js` and `parse.js`.
   If you cannot read it, do not run it on anything that matters.
2. **Test on your own account** and confirm it does what you expect, on your own emails.
3. **Have someone qualified review it** — your IT or security team, or a developer you
   trust. An outside pair of eyes catches what you and an AI assistant both missed. This
   is always faster than cleaning up afterwards.
4. **Get approval through the normal process** in your organisation. Do not skip it because
   the tool is small.

## The specific risk worth understanding

An Outlook add-in loads its code from a URL at the moment you open it. This one loads from
`events.jorgemenas.com`, a domain the author controls. That means the person who controls
that domain can change the code that runs inside your mailbox, with read access to the
message you have open, at any time, without telling you.

Nothing is being done with that access — the add-in sends nothing anywhere, and you can
verify that in the source. But the *capability* is inherent to every hosted Outlook add-in,
including this one. If that is not acceptable for your situation, host your own copy from a
domain you control, and read the code before you do.

## If you build your own version with AI

The same rules apply, and one more: AI assistants produce code that looks finished and
reads well, and that is not the same as code that is correct or safe. The confidence of the
output is not evidence. Review it, test it, and get a human who knows the area to look at
it before it touches anything real.
