/**
 * NMT Trade Log - Google Apps Script
 * HD-Email-Summary Project
 *
 * Polls Gmail every 1 minute for unread NMT trade alert emails.
 * POSTs raw email to Railway for AI parsing, then writes the parsed
 * trades directly to Google Sheets using SpreadsheetApp — no credentials
 * or service account keys required on the Railway/Python side.
 *
 * This mirrors the orbital-blazar pattern: Apps Script handles all Google
 * API calls natively; Railway only does AI work.
 *
 * Setup:
 *   1. Open script.google.com → New Project → paste this code
 *   2. Set TARGET_URL to your Railway deployment URL
 *   3. Set SHEET_ID to your "NMT Trade Log" Google Sheet ID
 *      (the long ID in the sheet URL: /spreadsheets/d/SHEET_ID/edit)
 *   4. Run installTrigger() once to activate 1-min polling
 *   5. Grant Gmail + Sheets permissions when prompted
 */

// ─── CONFIG ──────────────────────────────────────────────────────────────────

// Railway deployment URL — AI parsing only, no Sheets auth needed on Python side
var TARGET_URL = "https://hd-email-summary-production.up.railway.app/ingest/raw_email";

// Google Sheet ID for "NMT Trade Log"
// Get this from the sheet URL: docs.google.com/spreadsheets/d/THIS_PART/edit
var SHEET_ID = "YOUR_SHEET_ID_HERE";

// Tab name inside the spreadsheet
var SHEET_TAB = "Trade Log";

// Email query — same senders as orbital-blazar
var GMAIL_QUERY = [
  '(from:norseman@substack.com OR from:norsemanmarkettiming@substack.com)',
  'is:unread',
  'subject:("Trade Alert" OR "TRADE ALERT" OR "MULTI TRADE")'
].join(' ');

// Gmail label applied after successful processing
var PROCESSED_LABEL_NAME = "NMT/Logged";

// Sheet column headers (must match order in appendTradesToSheet)
var HEADERS = ["Date", "Portfolio", "Symbol", "Action", "% Portfolio", "Avg Fill", "Logged At"];

// ─── MAIN ─────────────────────────────────────────────────────────────────────

function processNMTEmails() {
  var threads = GmailApp.search(GMAIL_QUERY);

  if (threads.length === 0) {
    Logger.log("No unread NMT trade emails found.");
    return;
  }

  // Get or create the processed label
  var processedLabel = null;
  try {
    processedLabel = GmailApp.getUserLabelByName(PROCESSED_LABEL_NAME);
    if (!processedLabel) {
      processedLabel = GmailApp.createLabel(PROCESSED_LABEL_NAME);
      Logger.log("Created Gmail label: " + PROCESSED_LABEL_NAME);
    }
  } catch (e) {
    Logger.log("Could not get/create label: " + e);
  }

  threads.forEach(function(thread) {
    var messages = thread.getMessages();

    messages.forEach(function(msg) {
      if (!msg.isUnread()) return;

      var subject = msg.getSubject();
      var body    = msg.getPlainBody();
      var msgDate = msg.getDate();
      var isoDate = Utilities.formatDate(msgDate, "UTC", "yyyy-MM-dd");

      Logger.log("Processing email: " + subject + " (" + isoDate + ")");

      var payload = JSON.stringify({
        subject: subject,
        body:    body,
        date:    isoDate
      });

      try {
        // Step 1: Send to Railway for AI parsing (Railway returns parsed trades)
        var response = UrlFetchApp.fetch(TARGET_URL, {
          method:             "post",
          contentType:        "application/json",
          payload:            payload,
          muteHttpExceptions: true
        });

        var code = response.getResponseCode();
        var text = response.getContentText();
        Logger.log("Railway response (" + code + "): " + text);

        if (code !== 200) {
          Logger.log("Non-200 response — email left unread for retry: " + subject);
          return;
        }

        // Step 2: Parse the trades from the Railway response
        var result = JSON.parse(text);
        var trades = result.trades || [];

        if (!result.ok || trades.length === 0) {
          Logger.log("No trades extracted from: " + subject);
          // Still mark as read so we don't reprocess a non-trade email forever
          msg.markRead();
          if (processedLabel) thread.addLabel(processedLabel);
          return;
        }

        // Step 3: Write trades to Google Sheets natively (no credentials needed)
        appendTradesToSheet(trades);

        // Step 4: Mark email as processed
        msg.markRead();
        if (processedLabel) thread.addLabel(processedLabel);
        Logger.log("Logged " + trades.length + " trade(s) from: " + subject);

      } catch(e) {
        Logger.log("Error processing email: " + e);
        // Do NOT mark as read on error — will retry next minute
      }
    });
  });
}


// ─── SHEETS WRITER ────────────────────────────────────────────────────────────

/**
 * Appends parsed trade rows directly to the NMT Trade Log Google Sheet.
 * Uses SpreadsheetApp — runs natively in Apps Script, no auth required.
 *
 * @param {Array} trades - Array of trade objects from Railway AI parser
 */
function appendTradesToSheet(trades) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var ws = getOrCreateTab(ss);

  var loggedAt = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd HH:mm") + " UTC";

  var rows = trades.map(function(t) {
    return [
      t.trade_date     || "",
      t.portfolio      || "",
      (t.symbol        || "").toUpperCase(),
      t.action         || "",
      t.allocation_pct || 0,
      t.avg_fill       || 0,
      loggedAt
    ];
  });

  ws.getRange(ws.getLastRow() + 1, 1, rows.length, HEADERS.length)
    .setValues(rows);

  Logger.log("Appended " + rows.length + " row(s) to sheet.");
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
    // Bold the header row
    ws.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    Logger.log("Created tab: " + SHEET_TAB);
  } else if (ws.getLastRow() === 0) {
    // Sheet exists but is empty — write headers
    ws.appendRow(HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }

  return ws;
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
 * End-to-end test: sends the canonical example email through the full pipeline.
 * Run this from the Apps Script editor to verify Railway parsing + Sheet writing.
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
      "USAC 1% Added to 230k Port at 1:29 PM EST",
      "",
      "Average Fill: 28.97"
    ].join("\n"),
    date: "2026-05-14"
  });

  var response = UrlFetchApp.fetch(TARGET_URL, {
    method:             "post",
    contentType:        "application/json",
    payload:            payload,
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var text = response.getContentText();
  Logger.log("Test response (" + code + "): " + text);

  if (code === 200) {
    var result = JSON.parse(text);
    if (result.trades && result.trades.length > 0) {
      appendTradesToSheet(result.trades);
      Logger.log("Test: wrote " + result.trades.length + " row(s) to sheet.");
    } else {
      Logger.log("Test: no trades returned.");
    }
  }
}
