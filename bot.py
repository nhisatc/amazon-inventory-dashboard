"""
bot.py  —  US+ Health Inventory Slack Bot
------------------------------------------
Interactive Slack bot that answers inventory questions using live data
from the Google Sheet + Claude AI.

Run locally:   python bot.py
Deploy:        Railway / Render (set env vars, run this file)

Required env vars (add to .env or Railway dashboard):
  SLACK_BOT_TOKEN      xoxb-...   (Bot User OAuth Token)
  SLACK_APP_TOKEN      xapp-...   (App-Level Token, for Socket Mode)
  ANTHROPIC_API_KEY               (Claude API key)
  FORECAST_SHEET_ID               (Google Sheet ID)
  GOOGLE_CREDENTIALS_PATH         (path to JSON file, local)
  GOOGLE_CREDENTIALS_JSON         (JSON string, for Railway/cloud)
"""

import os, re, json, datetime
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(dotenv_path=Path(__file__).parent / ".env")

import gspread
from google.oauth2.service_account import Credentials
import anthropic
from slack_bolt import App
from slack_bolt.adapter.socket_mode import SocketModeHandler

# ── Clients ────────────────────────────────────────────────────────────────────

slack_app = App(token=os.environ["SLACK_BOT_TOKEN"])
claude    = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
SHEET_ID  = os.environ["FORECAST_SHEET_ID"]

def _gc():
    scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
    creds_path = os.environ.get("GOOGLE_CREDENTIALS_PATH")
    if creds_path and Path(creds_path).exists():
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    else:
        creds = Credentials.from_service_account_info(
            json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"]), scopes=scopes
        )
    return gspread.authorize(creds)

# ── Inventory data reader ──────────────────────────────────────────────────────

def get_inventory_context() -> str:
    """Read the Forecast & Reorder tab and format it for Claude."""
    try:
        gc = _gc()
        ss = gc.open_by_key(SHEET_ID)
        ws = ss.worksheet("Forecast & Reorder")
        rows = ws.get_all_values()
    except Exception as e:
        return f"Could not read inventory data: {e}"

    if len(rows) < 3:
        return "No inventory data available."

    headers = rows[1]
    data    = rows[2:]

    lines = []
    for row in data:
        if not any(row):
            continue
        d = dict(zip(headers, row + [""] * max(0, len(headers) - len(row))))
        lines.append(
            f"• {d.get('Product (Name + ASIN)', '?')} | "
            f"Available: {d.get('Available Stock', '?')} | "
            f"Inbound: {d.get('Inbound Stock', '?')} | "
            f"Reserved: {d.get('Reserved Stock', '?')} | "
            f"Forecast/mo: {d.get('Forecasted Demand', '?')} | "
            f"Reorder Point: {d.get('Reorder Point', '?')} | "
            f"Order Qty: {d.get('Order Qty', '?')} | "
            f"Status: {d.get('Status', '?')} | "
            f"Days of Stock: {d.get('Days of Stock', '?')}"
        )

    return "\n".join(lines) if lines else "No active SKUs found."

# ── System prompt ──────────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You are an inventory assistant for US+ Health, a health products company selling on Amazon FBA.
You have access to their live inventory dashboard, which updates every day at 7 AM.

You help the team with:
- Stock levels and days of stock remaining per product
- Which products need reordering and how many units
- Inbound shipments and pipeline coverage
- Forecasted demand and trends
- Products on hold (e.g. aged inventory clearance)

Guidelines:
- Be concise and direct — this is Slack, not an essay
- Lead with the number or answer, then explain if needed
- If something is urgent (Reorder Now), say so clearly
- "Hold" status means reordering is paused intentionally (business decision)
- Available + Inbound + Reserved = total pipeline stock
- Lead time is 60 days, so the model targets 2 months of forward stock
- Today: {today}

Live inventory data:
{inventory_data}
"""

# ── Event handlers ─────────────────────────────────────────────────────────────

@slack_app.event("app_mention")
def handle_mention(event, say):
    thread_ts = event.get("thread_ts", event["ts"])

    # Strip the @mention tag from the message
    raw = event.get("text", "")
    question = re.sub(r"<@[A-Z0-9]+>", "", raw).strip()

    if not question:
        say(
            text="Hi! Ask me anything about inventory — stock levels, what to reorder, days of stock, forecasts, and more.",
            thread_ts=thread_ts
        )
        return

    # Acknowledge immediately so the user knows it's working
    say(text="_Looking at the inventory data..._", thread_ts=thread_ts)

    inventory_data = get_inventory_context()
    today = datetime.date.today().strftime("%B %d, %Y")

    try:
        response = claude.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=1024,
            system=SYSTEM_PROMPT.format(today=today, inventory_data=inventory_data),
            messages=[{"role": "user", "content": question}]
        )
        answer = response.content[0].text
    except Exception as e:
        answer = f"Sorry, I ran into an error: {e}"

    say(text=answer, thread_ts=thread_ts)


@slack_app.event("message")
def handle_message(event, logger):
    # Suppress unhandled message subtype warnings
    pass


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("Starting US+ Health Inventory Bot (Socket Mode)...")
    handler = SocketModeHandler(slack_app, os.environ["SLACK_APP_TOKEN"])
    handler.start()
