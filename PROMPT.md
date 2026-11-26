# The prompt that built this

This add-in was built in one session with Claude Code, from three messages. The first one
did all the work. The other two were "continue now" and "it worked".

Read [DISCLAIMER.md](DISCLAIMER.md) before you run anything you build this way.

## The actual first prompt

Lightly tidied from voice dictation — the original had the usual spoken stumbles, and it
still worked. It does not need to be polished. It needs to be specific.

> I want you to act as an expert on add-ins for Microsoft Outlook and create an add-in, or
> the template for an add-in, that I could use to transform emails into events in my
> calendar.
>
> Say I have an email with an appointment, and that appointment is on Zoom, and Outlook
> doesn't pick it up automatically. What I would like is: the event name is the subject of
> the email, I add the attendees manually, and if there is a Teams link it goes to Teams —
> if not, it goes in the location and in the description. The description should also carry
> the body of the email, unless it's a very long one.
>
> I copied the documentation from Microsoft directly so you can check it. Read the
> documentation, and start with a plan to see what elements you need, then work from it.
>
> Don't go overboard, it's really simple: an email turns into an event that I can add to my
> calendar. That is it.
>
> Create everything in a specific folder called `event_addin_outlook`.
>
> https://learn.microsoft.com/en-us/office/dev/add-ins/develop/event-based-activation
> https://learn.microsoft.com/en-us/office/dev/add-ins/testing/test-debug-office-add-ins
> https://learn.microsoft.com/en-us/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets
> https://learn.microsoft.com/en-us/office/dev/add-ins/outlook/apis

## Why that prompt worked

Four things did the heavy lifting. Copy these, not the wording.

| What | Why it matters |
|---|---|
| **The official documentation links** | Office add-ins have a specific manifest format and a versioned API. Without the real docs, an assistant writes plausible-looking code against an API that does not exist. Paste the vendor's own pages. |
| **A worked example** ("an email with a Zoom appointment") | One concrete case pins down the behaviour far better than an abstract description. |
| **The rules spelled out** (Teams link → Teams, otherwise → location and description) | These became the actual branching logic in `parse.js`. Decisions you do not state get invented for you. |
| **"Don't go overboard, it's really simple"** | The single most useful sentence. Without a cap, assistants add build tools, frameworks, config and abstraction you did not ask for and now have to maintain. |

One more thing worth copying: it asked for **a plan first**, before any code. That makes the
assumptions visible while they are still cheap to correct.

## A template for your own small tool

Fill in the five brackets.

> I want you to act as an expert on **[the platform — e.g. Outlook add-ins, Google Apps
> Script, a Chrome extension]**, and build **[the smallest version of the thing]**.
>
> Here is the situation: **[one concrete example of the problem, with real details]**.
>
> What it should do: **[the rules, including what happens in the edge case you care about]**.
>
> Here is the official documentation: **[2-4 links from the vendor's own site]**.
>
> Start with a plan showing what pieces are needed, and wait for me to confirm before you
> write code. Don't go overboard — keep it as simple as it can be and still work. Put
> everything in a folder called **[name]**.

Then, when the plan looks right: `continue`.

## What the prompt did not cover, and cost time later

Worth adding to yours up front:

- **Where it will run when my laptop is off.** This add-in ran on `localhost` for about two
  and a half months, which meant it only worked while the laptop was on and a terminal was
  open. Hosting it as static files took about ten minutes and costs nothing. Ask for that
  on day one: *"it has to run without my machine being on — what are the free options?"*
- **What permissions it asks for.** This one requests `ReadItem`, the open message only.
  Ask the assistant to use the narrowest scope that works, and to tell you what it chose.
- **What could go wrong.** Ask directly: *"what are the security risks of this, and what
  would you warn a non-developer about?"* Then get a human to check the answer.
