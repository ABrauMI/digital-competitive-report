# Running the Digital Competitive Report Slack app

The Slack side has to be created through Slack's own dashboard — that part
can't be done for you. Everything below maps directly to what
`slack_app/app.py` expects, so following it exactly will make the bot work
the first time.

## 1. Create the app

1. Go to https://api.slack.com/apps -> **Create New App** -> **From scratch**.
2. Name it (e.g. "Digital Comp") and pick your workspace.

## 2. Turn on Socket Mode

1. Left sidebar -> **Socket Mode** -> toggle it on.
2. It'll prompt you to create an App-Level Token -- name it anything (e.g.
   `socket-token`), scope `connections:write` is added automatically.
3. Copy the generated token (starts with `xapp-`) -- this is `SLACK_APP_TOKEN`.

## 3. Add bot permissions

Left sidebar -> **OAuth & Permissions** -> **Scopes** -> **Bot Token Scopes**.
Add all of these:

- `commands`
- `chat:write`
- `files:read`
- `files:write`
- `channels:history`
- `groups:history`
- `im:history`
- `mpim:history`

## 4. Create the slash command

Left sidebar -> **Slash Commands** -> **Create New Command**:

- Command: `/digital-comp`
  Short description: `Build the digital competitive report from AdImpact exports`

Request URL isn't used with Socket Mode -- put anything, e.g.
`https://example.com`.

## 5. Turn on Interactivity

Left sidebar -> **Interactivity & Shortcuts** -> toggle on. Request URL isn't
used here either with Socket Mode.

## 6. Subscribe to message events

Left sidebar -> **Event Subscriptions** -> toggle on -> under **Subscribe to
bot events**, add:

- `message.channels`
- `message.groups`
- `message.im`
- `message.mpim`

(This is what lets the bot notice the AdImpact export files you reply with
in a thread.)

## 7. Install the app

**OAuth & Permissions** -> **Install to Workspace** -> approve. Copy the
**Bot User OAuth Token** (starts with `xoxb-`) -- this is `SLACK_BOT_TOKEN`.

## 8. Invite the bot

In whichever Slack channel you want to use it, run `/invite @Digital Comp`
(or whatever you named the app).

## 9. Run it

```bash
pip install -r requirements.txt
export SLACK_BOT_TOKEN=xoxb-...
export SLACK_APP_TOKEN=xapp-...
python3 -m slack_app.app
```

Leave that process running (Socket Mode keeps an open connection to Slack --
no public URL or hosting needed to try it out, though for regular team use
you'll eventually want it running somewhere persistent, like a small always-on
VM or Railway worker — see `RAILWAY_DEPLOY.md` — rather than a laptop).

## Using it

1. In the channel, type `/digital-comp`.
2. Optionally fill in a report title override in the modal — leave it blank
   to derive one from the export's Race, same as the CLI default.
3. The bot posts a message with a **Build Report** button. Reply to that
   message (in the thread) with your AdImpact exports:
   - **Spending Chart** export (`.xlsx`, required)
   - **Topline Creatives** export (`.xlsx`, optional — adds the Creative
     Timeline tab)

   Upload order doesn't matter and both can go in one message or separate
   ones — the bot tells the two apart by peeking at each file's header row.
   It confirms each one it recognizes.
4. Click **Build Report**. The bot runs the same pipeline as the CLI and
   uploads the resulting workbook into the thread, with a short summary
   (total spend, advertiser count, weeks covered, This Week status, and
   creative count if a Topline Creatives export was included).

If you reply with a file the bot doesn't recognize (wrong export type, or
an already-reformatted file), it says so rather than silently ignoring it —
reply again with the correct export and click Build once it's confirmed.
