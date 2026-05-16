"""
ai_parser.py — NMT Trade Email Parser
HD-Email-Summary Project

Extracts structured trade log rows from raw NMT newsletter emails.
Mirrors the orbital-blazar ai_parser.py pattern:
  - Primary: Gemini 2.5 Flash (with exponential backoff for 503/429)
  - Fallback: Gemini 2.5 Flash-Lite
  - Last resort: Claude claude-haiku-3-5
"""

import os
import json
import logging
import time
from datetime import date
from google import genai
import anthropic
from pydantic import BaseModel, Field

logger = logging.getLogger(__name__)


class TradeRow(BaseModel):
    """One row in the NMT Trade Log spreadsheet."""
    trade_date: str = Field(
        description="The date of the trade in MM/DD/YYYY format. "
                    "Infer from the email date header (e.g. 'May 14') using the current year. "
                    "If completely unavailable, use today's date."
    )
    portfolio: str = Field(
        description="The portfolio name: 'Berserker' or 'Huginn'. "
                    "Detect from subject line or body keywords: "
                    "'BERSERKER' / '200k port' / '200K port' → Berserker. "
                    "'HUGINN' / '50k port' / '50K port' → Huginn."
    )
    symbol: str = Field(description="The stock or ETF ticker symbol (e.g. 'CORZ', 'USAC').")
    action: str = Field(description="BUY or SELL")
    allocation_pct: float = Field(
        description="The percentage of the portfolio allocated to this trade "
                    "(e.g. '1% Added' → 1.0, '2% position' → 2.0). "
                    "Use 0.0 for SELL signals where no new allocation is stated."
    )
    avg_fill: float = Field(
        description="The average fill price from the 'Average Fill: X.XX' line. "
                    "Use 0.0 if not present."
    )


class MultiTradeExtract(BaseModel):
    trades: list[TradeRow] = Field(
        description="All independent trade rows found in this email."
    )


# ---------------------------------------------------------------------------
# Private helpers — Gemini and Claude callers
# ---------------------------------------------------------------------------

def _call_gemini(client, model: str, prompt: str, max_retries: int = 5) -> dict:
    """Call a Gemini model with exponential backoff for 503/429 errors."""
    base_delay = 5
    last_exc = None

    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model=model,
                contents=prompt,
                config={
                    'response_mime_type': 'application/json',
                    'response_schema': MultiTradeExtract,
                    'temperature': 0.1,
                },
            )
            return json.loads(response.text)
        except Exception as e:
            last_exc = e
            error_str = str(e).upper()
            is_retryable = (
                ("503" in error_str and "UNAVAILABLE" in error_str)
                or "429" in error_str
            )
            if is_retryable and attempt < max_retries - 1:
                delay = base_delay * (2 ** attempt)
                logger.warning(
                    f"[{model}] Transient error ({error_str[:60]}). "
                    f"Retrying in {delay}s... (Attempt {attempt + 1}/{max_retries})"
                )
                time.sleep(delay)
            else:
                logger.error(f"[{model}] Non-retryable error or retries exhausted: {e}")
                raise

    raise last_exc


def _call_claude(prompt: str, max_retries: int = 3) -> dict:
    """Claude claude-haiku-3-5 cross-provider fallback."""
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise ValueError("ANTHROPIC_API_KEY not set — cannot use Claude fallback.")

    client = anthropic.Anthropic(api_key=api_key)
    base_delay = 5
    last_exc = None

    structured_prompt = (
        prompt
        + "\n\nRespond with ONLY a valid JSON object in this exact format, no commentary:\n"
        + '{"trades": [{"trade_date": "MM/DD/YYYY", "portfolio": "Berserker|Huginn", '
        + '"symbol": "TICKER", "action": "BUY|SELL", "allocation_pct": 0.0, "avg_fill": 0.0}]}'
    )

    for attempt in range(max_retries):
        try:
            message = client.messages.create(
                model="claude-haiku-3-5",
                max_tokens=1024,
                messages=[{"role": "user", "content": structured_prompt}],
            )
            raw = message.content[0].text.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
            return json.loads(raw.strip())
        except (anthropic.RateLimitError, anthropic.APIStatusError) as e:
            last_exc = e
            if attempt < max_retries - 1:
                delay = base_delay * (2 ** attempt)
                logger.warning(
                    f"[claude-haiku-3-5] API error: {e}. "
                    f"Retrying in {delay}s... (Attempt {attempt + 1}/{max_retries})"
                )
                time.sleep(delay)
            else:
                logger.error(f"[claude-haiku-3-5] Exhausted all retries: {e}")
                raise
        except Exception as e:
            logger.error(f"[claude-haiku-3-5] Non-retryable error: {e}")
            raise

    raise last_exc


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def extract_trade_rows(email_subject: str, email_body: str, email_date: str = "") -> list[dict]:
    """
    Parse a raw NMT email and return a list of TradeRow dicts ready to
    append to the NMT Trade Log Google Sheet.

    Args:
        email_subject: The email subject line.
        email_body:    The plain-text email body.
        email_date:    Optional ISO date string (e.g. '2026-05-14') from the
                       Gmail message date header. Used as the trade date hint.

    Returns:
        List of dicts with keys: trade_date, portfolio, symbol, action,
        allocation_pct, avg_fill.
    """
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        raise ValueError("GEMINI_API_KEY not found in environment.")

    client = genai.Client(api_key=api_key)

    today_str = date.today().strftime("%m/%d/%Y")
    date_hint = f"The email was received on {email_date}." if email_date else f"Today's date is {today_str}."

    prompt = f"""
You are an expert financial trade parser for Norseman Market Timing (NMT) newsletters.
Extract ALL independent trades from the email below.

{date_hint}

Rules for portfolio:
- If the subject or body contains "BERSERKER", "200k port", or "200K port" → portfolio = "Berserker"
- If the subject or body contains "HUGINN", "50k port", or "50K port" → portfolio = "Huginn"
- A MULTI TRADE email may contain trades from BOTH portfolios — assign each trade the correct portfolio.
- Default to "Huginn" if unclear.

Rules for action:
- "Added", "bought", "added to" → BUY
- "Trimmed", "sold", "removed" → SELL

Rules for allocation_pct:
- Extract the percentage from phrases like "CORZ 1% Added" → allocation_pct = 1.0
- For BUY signals, this is the new portfolio allocation %.
- For SELL signals with no explicit % stated, use 0.0.

Rules for avg_fill:
- Extract from "Average Fill: 24.67" → avg_fill = 24.67
- If not present, use 0.0.

Rules for trade_date:
- Use the date from the email header/body (e.g. "May 14" → "05/14/2026").
- Format as MM/DD/YYYY.

EMAIL SUBJECT:
{email_subject}

EMAIL BODY:
{email_body}
"""

    logger.info("Sending NMT email to Gemini AI for trade extraction...")

    gemini_models = [
        ("gemini-2.5-flash", 5),
        ("gemini-2.5-flash-lite", 3),
    ]

    last_exc = None

    for model_name, retries in gemini_models:
        try:
            result = _call_gemini(client, model_name, prompt, max_retries=retries)
            if model_name != "gemini-2.5-flash":
                logger.warning(f"Used fallback model '{model_name}'.")
            trades = result.get("trades", [])
            logger.info(f"[{model_name}] Extracted {len(trades)} trade(s): {trades}")
            return trades
        except Exception as e:
            error_str = str(e).upper()
            if "503" in error_str or "429" in error_str:
                logger.error(
                    f"[{model_name}] Capacity error — trying next model..."
                )
                last_exc = e
                continue
            else:
                raise

    # Cross-provider fallback
    logger.warning("All Gemini models unavailable. Escalating to Claude claude-haiku-3-5.")
    try:
        result = _call_claude(prompt, max_retries=3)
        trades = result.get("trades", [])
        logger.warning(f"[claude-haiku-3-5] FALLBACK SUCCESS — Extracted {len(trades)} trade(s): {trades}")
        return trades
    except Exception as e:
        logger.error(f"[claude-haiku-3-5] Final fallback also failed: {e}")
        last_exc = e

    raise last_exc
