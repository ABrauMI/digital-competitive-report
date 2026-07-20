"""GPS Impact Digital Competitive Report Slack app.

One flow, one slash command:

/digital-comp
  1. Opens a modal to pick the data source (AdImpact or AdHawk) and an
     optional report title override.
  2. Submitting posts a message with a "Build Report" button and opens a
     session keyed by that message's ts.
  3. Reply to that message with your export(s):
       - AdImpact: a Spending Chart (.xlsx, required) and optionally a
         Topline Creatives export (.xlsx) — any order, one message or
         several. The bot tells the two apart by peeking at each file's
         header row, so upload order doesn't matter.
       - AdHawk: a single .csv export.
  4. Click "Build Report" and the bot uploads the rendered workbook.
     AdImpact gets Competitive Digital Report / Market Summary / This Week,
     plus a Creative Timeline tab if a creative export was supplied. AdHawk
     gets Competitive Digital Report / This Week only — it has no
     market/DMA data, so there's no Market Summary tab, and no Creative
     Timeline yet.

Run with:
  SLACK_BOT_TOKEN=xoxb-... SLACK_APP_TOKEN=xapp-... python3 -m slack_app.app
"""

import json
import logging
import os
import tempfile

import requests
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

from report import parse, parse_adhawk
from report.pipeline import build_adhawk_competitive_report, build_digital_competitive_report

from .session import sessions

BOT_TOKEN = os.environ["SLACK_BOT_TOKEN"]

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("digital_comp_slack_app")

app = App(token=BOT_TOKEN)


def _modal_view(channel_id):
    return {
        "type": "modal",
        "callback_id": "digital_comp_modal",
        "private_metadata": json.dumps({"channel": channel_id}),
        "title": {"type": "plain_text", "text": "Digital Comp Report"},
        "submit": {"type": "plain_text", "text": "Next: upload exports"},
        "close": {"type": "plain_text", "text": "Cancel"},
        "blocks": [
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": "Builds the Digital Competitive Report workbook from your export(s).",
                },
            },
            {
                "type": "input",
                "block_id": "source",
                "label": {"type": "plain_text", "text": "Data source"},
                "element": {
                    "type": "radio_buttons",
                    "action_id": "value",
                    "initial_option": {
                        "text": {"type": "plain_text", "text": "AdImpact"},
                        "value": "adimpact",
                    },
                    "options": [
                        {"text": {"type": "plain_text", "text": "AdImpact"}, "value": "adimpact"},
                        {"text": {"type": "plain_text", "text": "AdHawk"}, "value": "adhawk"},
                    ],
                },
            },
            {
                "type": "input",
                "block_id": "title_override",
                "optional": True,
                "label": {"type": "plain_text", "text": "Report title"},
                "element": {"type": "plain_text_input", "action_id": "value"},
                "hint": {"type": "plain_text", "text": "Leave blank to derive it from the export's Race."},
            },
        ],
    }


@app.command("/digital-comp")
def handle_digital_comp_command(ack, body, client):
    ack()
    client.views_open(trigger_id=body["trigger_id"], view=_modal_view(body["channel_id"]))


SOURCE_UPLOAD_PROMPT = {
    "adimpact": (
        "Reply in this thread with your AdImpact **Spending Chart** export (required), and "
        "optionally a **Topline Creatives** export for the Creative Timeline tab. "
        "Any order, one message or several — the bot tells them apart by their header row."
    ),
    "adhawk": "Reply in this thread with your **AdHawk** export (.csv).",
}


@app.view("digital_comp_modal")
def handle_modal_submission(ack, body, client, view):
    ack()
    metadata = json.loads(view["private_metadata"])
    channel = metadata["channel"]
    values = view["state"]["values"]
    source = values["source"]["value"]["selected_option"]["value"]
    title_override = (values["title_override"]["value"]["value"] or "").strip() or None

    user_id = body["user"]["id"]
    summary = (
        f"*Digital competitive report requested by <@{user_id}>* ({source})\n"
        + (f"Title override: *{title_override}*\n" if title_override else "")
        + "\n" + SOURCE_UPLOAD_PROMPT[source]
        + " Click *Build Report* below once it's in."
    )
    posted = client.chat_postMessage(
        channel=channel,
        text=summary,
        blocks=[
            {"type": "section", "text": {"type": "mrkdwn", "text": summary}},
            {
                "type": "actions",
                "elements": [
                    {
                        "type": "button",
                        "text": {"type": "plain_text", "text": "Build Report \U0001F4CA"},
                        "action_id": "build_digital_comp",
                        "style": "primary",
                    }
                ],
            },
        ],
    )
    sessions.create(posted["ts"], channel, source=source, title_override=title_override)


@app.event("message")
def handle_message_with_files(event, client, say):
    thread_ts = event.get("thread_ts")
    files = event.get("files")
    if not thread_ts or not files:
        return
    session = sessions.get(thread_ts)
    if session is None:
        return

    for f in files:
        name = f.get("name", "export")
        ext = ".csv" if session.source == "adhawk" else ".xlsx"
        if not name.lower().endswith(ext):
            continue
        url = f.get("url_private_download") or client.files_info(file=f["id"])["file"]["url_private_download"]
        r = requests.get(url, headers={"Authorization": f"Bearer {BOT_TOKEN}"})
        r.raise_for_status()
        tmp_path = os.path.join(tempfile.mkdtemp(prefix="digital_comp_"), name)
        with open(tmp_path, "wb") as out:
            out.write(r.content)

        if session.source == "adhawk":
            if parse_adhawk.classify_export(tmp_path) == "adhawk":
                session.set_spending(tmp_path, name)
                say(channel=session.channel, thread_ts=thread_ts, text=f"Got {name} — AdHawk export.")
            else:
                say(
                    channel=session.channel, thread_ts=thread_ts,
                    text=f"{name} doesn't look like an unmodified AdHawk export — expected Advertiser/"
                         "Election Name/Spend Platform/Spender Type columns.",
                )
            continue

        kind = parse.classify_export(tmp_path)
        if kind == "spending":
            session.set_spending(tmp_path, name)
            say(channel=session.channel, thread_ts=thread_ts, text=f"Got {name} — Spending Chart export.")
        elif kind == "creative":
            session.set_creative(tmp_path, name)
            say(channel=session.channel, thread_ts=thread_ts, text=f"Got {name} — Topline Creatives export.")
        else:
            say(
                channel=session.channel,
                thread_ts=thread_ts,
                text=f"Couldn't tell what {name} is — is it an unmodified AdImpact Spending Chart or "
                     "Topline Creatives export?",
            )


@app.action("build_digital_comp")
def handle_build_button(ack, body, client):
    ack()
    thread_ts = body["message"]["ts"]
    channel = body["channel"]["id"]
    session = sessions.get(thread_ts)

    if session is None or not session.spending_path:
        client.chat_postMessage(
            channel=channel,
            thread_ts=thread_ts,
            text="I don't have an export for this thread yet — reply with it first, then click Build again.",
        )
        return

    client.chat_postMessage(channel=channel, thread_ts=thread_ts, text="Building the report...")

    output_dir = tempfile.mkdtemp(prefix="digital_comp_out_")
    output_path = os.path.join(output_dir, "digital_competitive_report.xlsx")

    try:
        if session.source == "adhawk":
            result = build_adhawk_competitive_report(
                session.spending_path, output_path, title=session.title_override,
            )
        else:
            result = build_digital_competitive_report(
                session.spending_path,
                output_path,
                creative_path=session.creative_path,
                title=session.title_override,
            )
    except ValueError as e:
        client.chat_postMessage(
            channel=channel, thread_ts=thread_ts,
            text=f"Couldn't read that export: {e}",
        )
        return
    except Exception:
        logger.exception("Digital competitive report build failed")
        client.chat_postMessage(
            channel=channel,
            thread_ts=thread_ts,
            text="Something went wrong building the report — check the uploaded file is an unmodified export.",
        )
        return

    week_note = "" if result["this_week_in_export"] else " (no data yet — placeholder tab shown)"
    summary = (
        f"*{result['title']}*\n"
        f"Total tracked spend: *${result['grand_total']:,.0f}*  |  "
        f"Advertisers: *{result['advertiser_count']}*  |  Weeks: *{result['n_weeks']}*\n"
        f"This Week tab: *{result['this_week_iso']}*{week_note}"
    )
    if result["creative_count"] is not None:
        summary += f"\nCreative Timeline: *{result['creative_count']}* creatives"

    client.files_upload_v2(
        channel=channel,
        thread_ts=thread_ts,
        file=output_path,
        filename="digital_competitive_report.xlsx",
        initial_comment=summary,
    )
    sessions.discard(thread_ts)


if __name__ == "__main__":
    handler = SocketModeHandler(app, os.environ["SLACK_APP_TOKEN"])
    handler.start()
