"""
app.py — HD-Email-Summary Flask Service
Receives raw NMT emails from Gmail Apps Script, parses them with Gemini AI,
and returns structured trade data. Google Sheets writing is handled entirely
by the Apps Script (SpreadsheetApp) — no Google credentials needed here.

Endpoints:
  POST /ingest/raw_email   — Receives email, returns parsed trades as JSON
  GET  /health             — Basic uptime check
  GET  /test               — Fires the sample email through the parser (dev only)
"""

import os
import logging
from flask import Flask, request, jsonify
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Health
# ---------------------------------------------------------------------------

@app.route("/health", methods=["GET"])
def health():
    return {"ok": True, "service": "HD-Email-Summary"}, 200


# ---------------------------------------------------------------------------
# Core ingestion endpoint
# ---------------------------------------------------------------------------

@app.route("/ingest/raw_email", methods=["POST"])
def ingest_raw_email():
    """
    Receives a raw NMT email from Gmail Apps Script and parses it with
    Gemini AI. Returns the structured trade list as JSON — the Apps Script
    then writes the trades directly to Google Sheets via SpreadsheetApp.

    Expected JSON body:
      {
        "subject": "BERSERKER - MULTI TRADE ALERT",
        "body":    "...",
        "date":    "2026-05-14"   <- optional ISO date from Gmail
      }

    Response:
      {
        "ok":     true,
        "trades": [{"trade_date": "05/14/2026", "portfolio": "Berserker", ...}]
      }
    """
    from ai_parser import extract_trade_rows

    data = request.get_json(force=True, silent=True)
    if not data:
        return {"error": "Request body must be valid JSON."}, 400

    email_subject = data.get("subject", "")
    email_body    = data.get("body", "")
    email_date    = data.get("date", "")

    if not email_body:
        return {"error": "email_body is required."}, 400

    logger.info(f"Received email — Subject: '{email_subject}'")

    try:
        trades = extract_trade_rows(email_subject, email_body, email_date)
    except Exception as e:
        logger.error(f"AI extraction failed: {e}")
        return {"error": f"AI parsing failed: {e}"}, 500

    if not trades:
        logger.warning("AI returned no trades from this email.")
        return {"ok": False, "trades": [], "message": "No trades extracted from email."}, 200

    logger.info(f"Parsed {len(trades)} trade(s) — returning to Apps Script for Sheets write.")
    return {"ok": True, "trades": trades}, 200


# ---------------------------------------------------------------------------
# Dev test endpoint
# ---------------------------------------------------------------------------

@app.route("/test", methods=["GET"])
def test_parse():
    """
    Fires the canonical example email through the AI parser.
    Returns the parsed trades — use the Apps Script testWithSampleEmail()
    function to also verify the end-to-end Sheets write.
    """
    sample_subject = "BERSERKER - MULTI TRADE ALERT"
    sample_body = """READ IN APP

CORZ 1% Added to 230k Port at 1:27 PM EST

Average Fill: 24.67

USAC 1% Added to 230k Port at 1:29 PM EST

Average Fill: 28.97
"""
    sample_date = "2026-05-14"

    from ai_parser import extract_trade_rows

    try:
        trades = extract_trade_rows(sample_subject, sample_body, sample_date)
    except Exception as e:
        return {"error": f"AI parsing failed: {e}"}, 500

    return {"ok": True, "trades": trades}, 200


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port, debug=True)
