/**
 * NMT Trade Log - Google Apps Script
 * HD-Email-Summary Project
 *
 * Polls Gmail every 1 minute for unread NMT trade alert emails.
 * POSTs raw email to Railway for AI parsing, then writes parsed
 * trades directly to Google Sheets — no credentials needed.
 *
 * Columns written: Date | Time | Portfolio | Ticker | Action | Direction | Capital % | Price
 *
 * Action logic:
 *   - BUY   = first entry for this ticker (no prior open position in sheet)
 *   - ADD   = subsequent buy into an existing open position
 *   - TRIM  = partial exit (AI detected "Trimmed")
 *   - CLOSE = full exit (AI detected "Closed" / "Sold")
 *
 * Direction logic:
 *   - HEDGE = known inverse/short ETFs (QID, SDS, TWM, UVIX, SPXS, DXD, + others)
 *   - LONG  = everything else
 *
 * Setup:
 *   1. Paste this code into script.google.com → New Project
 *   2. Set SHEET_ID to your "NMT Trade Log" Google Sheet ID
 *   3. Run installTrigger() once — grant Gmail + Sheets permissions
 */

// ─── CONFIG ──────────────────────────────────────────────────────────────────

var TARGET_URL = "https://hd-email-summary-production.up.railway.app/ingest/raw_email";

// Sheet ID — the long string between /d/ and /edit in your Sheet URL
var SHEET_ID = "12Ksh7ISHq3hVT5KjaAWVQwsl49xlXnfplCxiKcRiQiE";

// Tab name inside the spreadsheet
var SHEET_TAB = "Trade Log";

// Gmail query — same NMT senders as orbital-blazar
var GMAIL_QUERY = [
  '(from:norseman@substack.com OR from:norsemanmarkettiming@substack.com)',
  'is:unread',
  'subject:("Trade Alert" OR "TRADE ALERT" OR "MULTI TRADE")'
].join(' ');

// Gmail label applied after successful processing
var PROCESSED_LABEL_NAME = "NMT/Logged";

// Column headers — matches the friend's trade log format
var HEADERS = ["Date", "Time", "Portfolio", "Ticker", "Action", "Direction", "Capital %", "Price"];

// ─── HEDGE ETF LIST ───────────────────────────────────────────────────────────
// Confirmed NMT hedges: QID, SDS, TWM, UVIX, SPXS, DXD
// Extended list covers other common inverse ETFs that may appear

var HEDGE_ETFS = [
  // Confirmed NMT hedges
  "QID",   // 2x inverse Nasdaq 100
  "SDS",   // 2x inverse S&P 500
  "TWM",   // 2x inverse Russell 2000
  "UVIX",  // 2x long VIX (inverse market)
  "SPXS",  // 3x inverse S&P 500
  "DXD",   // 2x inverse Dow Jones

  // Common inverse ETFs (extended coverage)
  "SH",    // 1x inverse S&P 500
  "PSQ",   // 1x inverse Nasdaq 100
  "SQQQ",  // 3x inverse Nasdaq 100
  "DOG",   // 1x inverse Dow Jones
  "SDOW",  // 3x inverse Dow Jones
  "RWM",   // 1x inverse Russell 2000
  "SRTY",  // 3x inverse Russell 2000
  "TZA",   // 3x inverse Russell 2000
  "TECS",  // 3x inverse Technology
  "SOXS",  // 3x inverse Semiconductors
  "LABD",  // 3x inverse Biotech
  "FAZ",   // 3x inverse Financials
  "SKF",   // 2x inverse Financials
  "DUST",  // 2x inverse Gold Miners
  "VIXY",  // 1x long VIX
  "UVXY",  // 1.5x long VIX
];


// ─── HELPERS ──────────────────────────────────────────────────────────────────

/**
 * Safely parse a JSON string. Returns null if the response is not valid JSON
 * (e.g. Railway returns an HTML error page on cold start / 502).
 */
function safeParseJSON(text) {
  try {
    return JSON.parse(text);
  } catch(e) {
    Logger.log("Could not parse response as JSON (Railway may be cold-starting): " + text.substring(0, 120));
    return null;
  }
}


function processNMTEmails() {
  var threads = GmailApp.search(GMAIL_QUERY);

  if (threads.length === 0) {
    Logger.log("No unread NMT trade emails found.");
    return;
  }

  var processedLabel = getOrCreateLabel(PROCESSED_LABEL_NAME);

  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      if (!msg.isUnread()) return;

      var subject = msg.getSubject();
      var body    = msg.getPlainBody();
      var msgDate = msg.getDate();
      var isoDate = Utilities.formatDate(msgDate, "UTC", "yyyy-MM-dd");
      // Fallback time from email receive timestamp (Eastern)
      var emailTime = Utilities.formatDate(msgDate, "America/New_York", "h:mm a");

      Logger.log("Processing: " + subject + " (" + isoDate + ")");

      var payload = JSON.stringify({ subject: subject, body: body, date: isoDate });

      try {
        var response = UrlFetchApp.fetch(TARGET_URL, {
          method:             "post",
          contentType:        "application/json",
          payload:            payload,
          muteHttpExceptions: true
        });

        var code   = response.getResponseCode();
        var result = safeParseJSON(response.getContentText());

        if (!result) {
          Logger.log("Invalid response from Railway (HTML/empty) — leaving email unread for retry.");
          return;
        }

        var trades = result.trades || [];
        if (trades.length === 0) {
          Logger.log("No trades extracted — marking read to avoid reprocessing.");
          msg.markRead();
          if (processedLabel) thread.addLabel(processedLabel);
          return;
        }

        // Write trades to Google Sheet natively (no credentials needed)
        appendTradesToSheet(trades, emailTime);

        msg.markRead();
        if (processedLabel) thread.addLabel(processedLabel);
        Logger.log("Logged " + trades.length + " trade(s) from: " + subject);

      } catch(e) {
        Logger.log("Error processing email: " + e);
        // Do NOT mark as read — will retry next minute
      }
    });
  });
}


// ─── SHEETS WRITER ────────────────────────────────────────────────────────────

/**
 * Appends parsed trade rows to the NMT Trade Log Google Sheet.
 * Resolves BUY vs ADD based on existing position history in the sheet.
 *
 * @param {Array}  trades    - Array of trade objects from Railway AI parser
 * @param {string} emailTime - Fallback time string if AI didn't extract one
 */
function appendTradesToSheet(trades, emailTime) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = getOrCreateTab(ss);

  // Read all existing data once (for BUY vs ADD detection)
  var existingData = ws.getDataRange().getValues();

  var rows = trades.map(function(t) {
    var symbol    = (t.symbol || "").toUpperCase();
    var portfolio = t.portfolio || "Huginn";
    var rawAction = (t.action || "BUY").toUpperCase();
    var time      = (t.trade_time && t.trade_time.trim() !== "") ? t.trade_time : emailTime;

    var action    = resolveAction(existingData, symbol, rawAction);
    var direction = getDirection(symbol);

    // Update existingData in-memory so subsequent trades in same email
    // correctly detect each other (e.g. two new BUYs in one multi-trade email)
    // NOTE: column layout is now 8 wide — Date, Time, Portfolio, Ticker, Action, Direction, Capital%, Price
    if (action === "BUY" || action === "ADD") {
      existingData.push([t.trade_date, time, portfolio, symbol, action, direction, t.allocation_pct, t.avg_fill]);
    }

    return [
      t.trade_date     || "",   // A: Date
      time,                      // B: Time
      portfolio,                 // C: Portfolio (Berserker / Huginn)
      symbol,                    // D: Ticker
      action,                    // E: Action (BUY / ADD / TRIM / CLOSE)
      direction,                 // F: Direction (LONG / HEDGE)
      t.allocation_pct || 0,    // G: Capital %
      t.avg_fill       || 0     // H: Price
    ];
  });

  ws.getRange(ws.getLastRow() + 1, 1, rows.length, HEADERS.length)
    .setValues(rows);

  Logger.log("Appended " + rows.length + " row(s) to sheet.");
}


// ─── HELPERS ──────────────────────────────────────────────────────────────────

/**
 * Determines whether a BUY-type signal is a fresh BUY or an ADD to an existing
 * position by scanning all prior rows in the sheet for the same ticker.
 *
 * Logic: If the ticker has a prior BUY or ADD with no subsequent CLOSE,
 *        the position is still open → label new signal as ADD.
 *
 * @param {Array}  existingData - 2D array of all current sheet values
 * @param {string} symbol       - Ticker to look up
 * @param {string} rawAction    - AI-returned action: BUY | TRIM | CLOSE
 * @returns {string}            - BUY | ADD | TRIM | CLOSE
 */
function resolveAction(existingData, symbol, rawAction) {
  if (rawAction === "TRIM" || rawAction === "CLOSE") return rawAction;

  // Scan existing rows (skip header row at index 0)
  var hasOpenPosition = false;
  for (var i = 1; i < existingData.length; i++) {
    var rowTicker = (existingData[i][3] || "").toString().toUpperCase();  // col D (index 3) after Portfolio inserted at C
    var rowAction = (existingData[i][4] || "").toString().toUpperCase();   // col E (index 4)

    if (rowTicker === symbol) {
      if (rowAction === "BUY" || rowAction === "ADD") {
        hasOpenPosition = true;
      } else if (rowAction === "CLOSE") {
        hasOpenPosition = false; // Position was fully closed — reset
      }
    }
  }

  return hasOpenPosition ? "ADD" : "BUY";
}


/**
 * Returns HEDGE for known inverse/short ETFs, LONG for everything else.
 *
 * @param {string} symbol - Ticker symbol
 * @returns {string}      - "HEDGE" or "LONG"
 */
function getDirection(symbol) {
  return HEDGE_ETFS.indexOf(symbol.toUpperCase()) !== -1 ? "HEDGE" : "LONG";
}


/**
 * Returns the target worksheet tab, creating it with headers if it doesn't exist.
 */
function getOrCreateTab(ss) {
  var ws = ss.getSheetByName(SHEET_TAB);

  if (!ws) {
    ws = ss.insertSheet(SHEET_TAB);
    ws.appendRow(HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    Logger.log("Created tab: " + SHEET_TAB);
  } else if (ws.getLastRow() === 0 || ws.getRange(1, 1).getValue() !== HEADERS[0]) {
    ws.insertRowBefore(1);
    ws.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
    ws.setFrozenRows(1);
    Logger.log("Wrote headers to existing tab.");
  }

  return ws;
}


/**
 * Gets or creates a Gmail label by name.
 */
function getOrCreateLabel(name) {
  try {
    var label = GmailApp.getUserLabelByName(name);
    if (!label) {
      label = GmailApp.createLabel(name);
      Logger.log("Created Gmail label: " + name);
    }
    return label;
  } catch(e) {
    Logger.log("Could not get/create label: " + e);
    return null;
  }
}


// ─── TRIGGER MANAGEMENT ──────────────────────────────────────────────────────

function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "processNMTEmails") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("processNMTEmails")
    .timeBased()
    .everyMinutes(1)
    .create();

  Logger.log("Trigger installed: processNMTEmails runs every 1 minute.");
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "processNMTEmails") {
      ScriptApp.deleteTrigger(t);
      Logger.log("Trigger removed.");
    }
  });
}


// ─── TEST ─────────────────────────────────────────────────────────────────────

/**
 * End-to-end test covering BUY, ADD, TRIM, CLOSE, and HEDGE direction.
 * Run from Apps Script editor to verify the full pipeline.
 *
 * Expected sheet output:
 *   CORZ  → BUY  / LONG  (first entry)
 *   CORZ  → ADD  / LONG  (second entry — same ticker, already in sheet)
 *   QID   → BUY  / HEDGE (first hedge position)
 *   USAC  → TRIM / LONG  (partial exit)
 */
function testWithSampleEmail() {
  var payload = JSON.stringify({
    subject: "BERSERKER - MULTI TRADE ALERT",
    body: [
      "READ IN APP",
      "",
      "CORZ 1% Added to 230k Port at 1:27 PM EST",
      "",
      "Average Fill: 24.67",
      "",
      "USAC Trimmed at 2:15 PM EST",
      "",
      "Average Fill: 28.97",
      "",
      "QID 2% Added to 230k Port at 2:30 PM EST",
      "",
      "Average Fill: 19.50"
    ].join("\n"),
    date: "2026-05-14"
  });

  var response = UrlFetchApp.fetch(TARGET_URL, {
    method:             "post",
    contentType:        "application/json",
    payload:            payload,
    muteHttpExceptions: true
  });

  var code   = response.getResponseCode();
  var result = safeParseJSON(response.getContentText());
  Logger.log("Test response (" + code + "): " + JSON.stringify(result));

  if (code === 200 && result.trades && result.trades.length > 0) {
    var emailTime = Utilities.formatDate(new Date(), "America/New_York", "h:mm a");
    appendTradesToSheet(result.trades, emailTime);
    Logger.log("Test: wrote " + result.trades.length + " row(s) to sheet.");
  } else {
    Logger.log("Test: no trades returned or error.");
  }
}


// ─── BACKFILL ─────────────────────────────────────────────────────────────────

/**
 * ONE-TIME backfill: processes all NMT trade emails from the last N days.
 * Handles both read AND unread emails.
 * Skips threads already tagged NMT/Logged (no duplicates).
 *
 * Run ONCE from the Apps Script editor to seed historical trades.
 * The live trigger (processNMTEmails) handles new emails going forward.
 */
function backfillLast5Days() {
  backfillLastNDays(5);
}

function backfillLast10Days() {
  backfillLastNDays(10);
}

function backfillLastNDays(n) {
  var days = n || 5;

  var query = [
    '(from:norseman@substack.com OR from:norsemanmarkettiming@substack.com)',
    'subject:("Trade Alert" OR "TRADE ALERT" OR "MULTI TRADE")',
    'newer_than:' + days + 'd'
  ].join(' ');

  Logger.log("Backfill query: " + query);
  var threads = GmailApp.search(query);
  Logger.log("Found " + threads.length + " thread(s) to backfill.");

  if (threads.length === 0) {
    Logger.log("No NMT emails found in the last " + days + " days.");
    return;
  }

  var processedLabel = getOrCreateLabel(PROCESSED_LABEL_NAME);
  var totalLogged = 0;
  var totalSkipped = 0;

  threads.forEach(function(thread) {
    // Skip threads already labeled NMT/Logged
    var threadLabels = thread.getLabels().map(function(l) { return l.getName(); });
    if (threadLabels.indexOf(PROCESSED_LABEL_NAME) !== -1) {
      Logger.log("Skipping already-processed: " + thread.getFirstMessageSubject());
      totalSkipped++;
      return;
    }

    thread.getMessages().forEach(function(msg) {
      var subject   = msg.getSubject();
      var body      = msg.getPlainBody();
      var msgDate   = msg.getDate();
      var isoDate   = Utilities.formatDate(msgDate, "UTC", "yyyy-MM-dd");
      var emailTime = Utilities.formatDate(msgDate, "America/New_York", "h:mm a");

      Logger.log("Backfilling: " + subject + " (" + isoDate + ")");

      var payload = JSON.stringify({ subject: subject, body: body, date: isoDate });

      try {
        var response = UrlFetchApp.fetch(TARGET_URL, {
          method:             "post",
          contentType:        "application/json",
          payload:            payload,
          muteHttpExceptions: true
        });

        var code   = response.getResponseCode();
        var result = safeParseJSON(response.getContentText());

        if (!result) {
          Logger.log("Invalid response from Railway — skipping: " + subject);
          return;
        }

        var trades = result.trades || [];
        if (trades.length > 0) {
          appendTradesToSheet(trades, emailTime);
          totalLogged += trades.length;
          Logger.log("Logged " + trades.length + " trade(s) from: " + subject);
        } else {
          Logger.log("No trades found in: " + subject);
        }

        // Mark processed so live trigger won't re-process
        if (processedLabel) thread.addLabel(processedLabel);

      } catch(e) {
        Logger.log("Error processing: " + subject + " — " + e);
      }
    });
  });

  Logger.log("=== Backfill complete: " + totalLogged + " trade(s) logged, " + totalSkipped + " skipped. ===");
}
