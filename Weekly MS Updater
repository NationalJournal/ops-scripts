/**
 * Run weekly on Monday 10 AM
 * 1) Given the import of the TEAM2026Annual sheet, save a Drive copy of the current sheet to Weekly MS Team Pacing folder
 * 2) Search the Drive folder for the previous week's copy
 * 3) Compare previous to current using the opportunity's name as a stable key column and pull changes
 * 4) Email a summary of week to week changes to relevant parties
 */


const DRIVE_FOLDER_ID = "11cNDhwkaRgAWYf2LW-rqvGTz71n9PR1V";
const LIVE_IMPORT_SHEET = "TEAM2026Annual";
const KEY_COL_ID = "See in SFDC";

const REPORT_SHEET_NAME = "TEAM2026Annual Updates";
const METRICS_REPORT_NAME = "Annual Metrics Updates";
//const EMAIL_RECIPIENTS = ["ltuider@nationaljournal.com"];
const EMAIL_RECIPIENTS = ["ltuider@nationaljournal.com",
                          "abari@nationaljournal.com",
                          "ailic@nationaljournal.com",
                          "eakhtarzandi@nationaljournal.com",
                          "eorem@nationaljournal.com",
                          "pward@nationaljournal.com",
                          "rkatona@nationaljournal.com",
                          "kturpin@nationaljournal.com"];
const IGNORE_COLS = ["Opportunity Owner Name", "Account Primary ChurnScore Value"];

const STATUS_COL_NAME = "Stage";  
const NAME_COL_NAME   = "Name";    
const PRICE_COL_NAME  = "Annualized Renewal Price (Won)";
const ACP_COL_NAME = "Annualized Current Price";
const OWNER_COL_NAME = "Opportunity Owner Name";
const DIAGNOSED_COL_NAME = "Diagnosed in Forecast?";

const METRICS = [
  {name: "Renewable Revenue Pool", namedRange: "renewable_revenue_pool", format: "currency"},
  {name: "Renewal Goal", namedRange: "renewal_goal", format: "currency"},
  {name: "Renewed Revenue", namedRange: "renewed_revenue", format: "currency"},
  {name: "Revenue Retention", namedRange: "revenue_retention", format: "percentage"},
  {name: "Unit Retention", namedRange: "unit_retention", format: "percentage"},
  {name: "Gap to Goal", namedRange: "gap_to_goal", format: "currency"},
  {name: "Outstanding Revenue Pool", namedRange: "outstanding_revenue_pool", format: "currency"},
  {name: "Gap to Goal with Outstanding Revenue Pool", namedRange: "gap_to_goal_w_outstanding", format: "currency"},
  {name: "Outstanding Forecast (Worst Case)", namedRange: "outstanding_worst_case", format: "currency"},
  {name: "Outstanding Forecast (Most Likely)", namedRange: "outstanding_most_likely", format: "currency"},
  {name: "Outstanding Forecast (Best Case)", namedRange: "outstanding_best_case", format: "currency"},
  {name: "Gap to Goal w/ Forecast (Worst Case)", namedRange: "gap_forecast_worst_case", format: "currency"},
  {name: "Gap to Goal w/ Forecast (Most Likely)", namedRange: "gap_forecast_most_likely", format: "currency"},
  {name: "Gap to Goal w/ Forecast (Best Case)", namedRange: "gap_forecast_best_case", format: "currency"},
  {name: "Forecasted Revenue Retention (Worst Case)", namedRange: "forecasted_revenue_worst_case", format: "percentage"},
  {name: "Forecasted Revenue Retention (Most Likely)", namedRange: "forecasted_revenue_most_likely", format: "percentage"},
  {name: "Forecasted Revenue Retention (Best Case)", namedRange: "forecasted_revenue_best_case", format: "percentage"},
  {name: "Forecasted Unit Retention (Worst Case)", namedRange: "forecasted_unit_worst_case", format: "percentage"},
  {name: "Forecasted Unit Retention (Most Likely)", namedRange: "forecasted_unit_most_likely", format: "percentage"},
  {name: "Forecasted Unit Retention (Best Case)", namedRange: "forecasted_unit_best_case", format: "percentage"},
  {name: "Closed Won", namedRange: "closed_won", format: "number"},
  {name: "Closed Lost", namedRange: "closed_lost", format: "number"},
  {name: "Outstanding", namedRange: "outstanding", format: "number"},
  {name: "Q1 Revenue Retention", namedRange: "q1_revenue_retention", format: "percentage"},
  {name: "Q2 Revenue Retention", namedRange: "q2_revenue_retention", format: "percentage"},
  {name: "Q3 Revenue Retention", namedRange: "q3_revenue_retention", format: "percentage"},
  {name: "Q4 Revenue Retention", namedRange: "q4_revenue_retention", format: "percentage"},
  {name: "Q1 Revenue Retention (Most Likely)", namedRange: "q1_most_likely", format: "percentage"},
  {name: "Q2 Revenue Retention (Most Likely)", namedRange: "q2_most_likely", format: "percentage"},
  {name: "Q3 Revenue Retention (Most Likely)", namedRange: "q3_most_likely", format: "percentage"},
  {name: "Q4 Revenue Retention (Most Likely)", namedRange: "q4_most_likely", format: "percentage"},
  {name: "Q1 Revenue Retention (Best Case)", namedRange: "q1_best_case", format: "percentage"},
  {name: "Q2 Revenue Retention (Best Case)", namedRange: "q2_best_case", format: "percentage"},
  {name: "Q3 Revenue Retention (Best Case)", namedRange: "q3_best_case", format: "percentage"},
  {name: "Q4 Revenue Retention (Best Case)", namedRange: "q4_best_case", format: "percentage"},
];

const EMAIL_METRIC_WHITELIST = [
  "Renewed Revenue",
  "Revenue Retention",
  "Unit Retention",
  "Forecasted Revenue Retention (Most Likely)",
  "Forecasted Revenue Retention (Best Case)",
];


// ---------------- UI MENU ----------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔄 Refresh")
    .addItem("Refresh reports (will not send an email)", "refreshData")
    .addToUi();
}


function refreshData() {
  const curr_sheet = SpreadsheetApp.getActive();
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  // Find most recent existing snapshots only
  const [latestSnapshot] = find_two_most_recent_snapshots(folder, LIVE_IMPORT_SHEET);
  if (!latestSnapshot) {
    throw new Error("No CSV snapshot found in the folder yet.");
  }

  const [latestMetricSnapshot] = find_two_most_recent_metrics(folder, LIVE_IMPORT_SHEET);
  if (!latestMetricSnapshot) {
    throw new Error("No metrics snapshot found in the folder yet.");
  }

  // CURRENT live sheet values
  const liveSheet = curr_sheet.getSheetByName(LIVE_IMPORT_SHEET);
  if (!liveSheet) throw new Error(`Missing live sheet tab: ${LIVE_IMPORT_SHEET}`);

  const currentRows = liveSheet.getDataRange().getDisplayValues();

  // MOST RECENT saved CSV snapshot
  const snapshotRows = Utilities.parseCsv(latestSnapshot.getBlob().getDataAsString());

  // Compare current live sheet vs latest snapshot
  const diff = find_diff_rows(snapshotRows, currentRows, KEY_COL_ID, IGNORE_COLS);
  write_diff_report(curr_sheet, latestSnapshot.getName(), `${LIVE_IMPORT_SHEET} (current live)`, diff);

  // CURRENT live metrics
  const currentMetrics = read_metrics_from_named_ranges(curr_sheet);

  // MOST RECENT saved metrics snapshot
  const snapshotMetrics = JSON.parse(latestMetricSnapshot.getBlob().getDataAsString());

  // Compare current live metrics vs latest metrics snapshot
  const diff2 = diffMetrics(snapshotMetrics, currentMetrics);
  write_diff_metrics_to_report(
    curr_sheet,
    latestMetricSnapshot.getName(),
    `${LIVE_IMPORT_SHEET} metrics (current live)`,
    diff2
  );

  Logger.log("Manual refresh complete.");
  Logger.log("Compared live sheet to snapshot: %s", latestSnapshot.getName());
  Logger.log("Compared live metrics to snapshot: %s", latestMetricSnapshot.getName());
}


/** MAIN FUNCTION (schedule to run Monday mornings) */
function monday_snapshot_diff_and_email() {
  const curr_sheet = SpreadsheetApp.getActive();
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

  //Step 1
  const todaysFile = save_csv_copy(curr_sheet, folder, LIVE_IMPORT_SHEET);
  const todaysMetrics = read_metrics_from_named_ranges(curr_sheet);
  const todaysMetricsFile = save_metrics_csv_copy(folder, LIVE_IMPORT_SHEET, todaysMetrics);

  //Step 2
  const [curr, prev] = find_two_most_recent_snapshots(folder, LIVE_IMPORT_SHEET);
  if (!curr) throw new Error("No CSV snapshots found in the folder yet.");
  if (!prev) throw new Error(`Found only ONE snapshot:\n${curr.getName()}\n\nCreate at least one more snapshot for comparison.`);

  const [curr_m, prev_m] = find_two_most_recent_metrics(folder, LIVE_IMPORT_SHEET);
  if (!curr_m) throw new Error("No CSV metrics snapshots found in the folder yet.");
  if (!prev_m) throw new Error(`Found only ONE metrics snapshot:\n${curr_m.getName()}\n\nCreate at least one more snapshot for comparison.`);

  //Step 3
  const currRows = Utilities.parseCsv(curr.getBlob().getDataAsString());
  const prevRows = Utilities.parseCsv(prev.getBlob().getDataAsString());
  const diff = find_diff_rows(prevRows, currRows, KEY_COL_ID, IGNORE_COLS);
  
  write_diff_report(curr_sheet, prev.getName(), curr.getName(), diff);

  const currMetrics = JSON.parse(curr_m.getBlob().getDataAsString());
  const prevMetrics = JSON.parse(prev_m.getBlob().getDataAsString());
  const diff2 = diffMetrics(prevMetrics, currMetrics);

  write_diff_metrics_to_report(curr_sheet, prev_m.getName(), curr_m.getName(), diff2)

  //Step 4 
  sendChangesEmail_(curr_sheet, diff.closedWon, diff.closedLost, diff2, prev.getName(), curr.getName(), curr_sheet.getUrl(), EMAIL_RECIPIENTS);

  //Logging
  Logger.log("Compared: %s -> %s", prev.getName(), curr.getName());
  Logger.log("Compared: %s -> %s", prev_m.getName(), curr_m.getName());
  Logger.log("COMPLETE!");
}


/** STEP 1: SNAPSHOT CSV */
function save_csv_copy(curr_sheet, folder, sheetName) {
  const sheet = curr_sheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Missing sheet/tab: ${sheetName}`);

  // Use DISPLAY values so dates remain mm/dd/yyyy, etc.
  const values = sheet.getDataRange().getDisplayValues();
  const csv = toCsv_(values);

  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const fileName = `${sheetName}_snapshot_${stamp}.csv`;

  // If today's file already exists, replace it (prevents duplicates if you rerun)
  const existing = findFileByExactName_(folder, fileName);
  if (existing) existing.setTrashed(true);

  return folder.createFile(fileName, csv, MimeType.CSV);
}

function toCsv_(grid) {
  return grid.map(row => row.map(csvEscape_).join(",")).join("\r\n");
}

function csvEscape_(value) {
  if (value === null || value === undefined) return "";
  let s = String(value);
  if (s.includes('"')) s = s.replace(/"/g, '""');
  if (/[",\r\n]/.test(s)) s = `"${s}"`;
  return s;
}

function findFileByExactName_(folder, fileName) {
  const it = folder.getFilesByName(fileName);
  return it.hasNext() ? it.next() : null;
}

/** SNAPSHOT METRICS */
function save_metrics_csv_copy(folder, prefix, metricsObj) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const fileName = `${prefix}_metrics_snapshot_${stamp}.json`;

  // replace if rerun same day
  const it = folder.getFilesByName(fileName);
  if (it.hasNext()) it.next().setTrashed(true);

  return folder.createFile(fileName, JSON.stringify(metricsObj, null, 2), MimeType.PLAIN_TEXT);
}

function read_metrics_from_named_ranges(curr_sheet) {
  const out = {};
  METRICS.forEach(m => {
    const r = curr_sheet.getRangeByName(m.namedRange);
    if (!r) throw new Error(`Missing named range: ${m.namedRange}`);
    out[m.name] = r.getDisplayValue(); // keeps "57%" "$1,234" etc.
  });
  return out;
}


/** STEP 2: FIND SNAPSHOTS */
function find_two_most_recent_snapshots(folder, importSheet) {
  const it = folder.getFiles();
  const matches = [];

  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();

    if (!name.startsWith(`${importSheet}_snapshot_`)) continue;
    if (!name.endsWith(".csv")) continue;

    matches.push({file: f, t: f.getLastUpdated().getTime() });
  }

  matches.sort((a,b) => b.t - a.t);
  
  const curr = matches[0]?.file || null;
  const prev = matches[1]?.file || null;
  
  return [curr, prev];
}

function find_two_most_recent_metrics(folder, importSheet) {
  const it = folder.getFiles();
  const matches = [];

  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();
    if (!name.startsWith(`${importSheet}_metrics_snapshot_`)) continue;
    if (!name.endsWith(".json")) continue;
    matches.push({ file: f, t: f.getLastUpdated().getTime() });
  }

  matches.sort((a, b) => b.t - a.t);
  return [matches[0]?.file || null, matches[1]?.file || null];
}


/** STEP 3: DIFF WITH BEFORE / AFTER */
function find_diff_rows(prevRows, currRows, keyColName, ignoreCols) {
  const prevHeaders = prevRows[0];
  const currHeaders = currRows[0];

  const prevKeyIdx = prevHeaders.indexOf(keyColName);
  const currKeyIdx = currHeaders.indexOf(keyColName);
  if (prevKeyIdx === -1 || currKeyIdx === -1) { throw new Error(`Key column "${keyColName}" not found in one of the CSVs`); }

  const commonHeaders = currHeaders.filter(h => prevHeaders.includes(h));
  const prevIdx = Object.fromEntries(commonHeaders.map(h => [h, prevHeaders.indexOf(h)]));
  const currIdx = Object.fromEntries(commonHeaders.map(h => [h, currHeaders.indexOf(h)]));

  const ignoreSet = new Set(ignoreCols || []);
  ignoreSet.add(keyColName); //never treat the key itself as a "changed field"

  const prevMap = build_csv_map(prevRows, prevKeyIdx);
  const currMap = build_csv_map(currRows, currKeyIdx);

  const added = [];
  const removed = [];
  const modified = [];
  const closedWon = [];
  const closedLost = [];

  const statusIdx = commonHeaders.indexOf(STATUS_COL_NAME);
  const nameIdx   = commonHeaders.indexOf(NAME_COL_NAME);
  const priceIdx  = commonHeaders.indexOf(PRICE_COL_NAME);
  const acpIdx = commonHeaders.indexOf(ACP_COL_NAME);
  const ownerIdx = commonHeaders.indexOf(OWNER_COL_NAME);
  const diagnosedIdx = commonHeaders.indexOf(DIAGNOSED_COL_NAME);

  //added or modified
  for (const [key, currRow] of currMap.entries()) {
    if(!prevMap.has(key)) {
      added.push({ key, row: project_row(currRow, commonHeaders, currIdx) });
      continue;
    }

    const prevRow = prevMap.get(key);
    const changedFields = [];

    for (const h of commonHeaders) {
      if (ignoreSet.has(h)) continue;

      const a = norm(prevRow[prevIdx[h]]);
      const b = norm(currRow[currIdx[h]]);
      if (a !== b) changedFields.push(h);
    }

    if (changedFields.length) {
      const beforeRow = project_row(prevRow, commonHeaders, prevIdx);
      const afterRow  = project_row(currRow, commonHeaders, currIdx);

      modified.push({
        key,
        changedFields,
        beforeRow,
        afterRow,
      });

      if (
        statusIdx !== -1 &&
        changedFields.includes(STATUS_COL_NAME) &&
        norm(afterRow[statusIdx]) === "Closed Won"
      ) {
        closedWon.push({
          key,
          owner: ownerIdx !== -1 ? afterRow[ownerIdx] : "",
          name: nameIdx !== -1 ? afterRow[nameIdx] : "",
          prevPrice: priceIdx !== -1 ? beforeRow[priceIdx] : "",
          currPrice: priceIdx !== -1 ? afterRow[priceIdx] : "",
          prevStatus: beforeRow[statusIdx],
          currStatus: afterRow[statusIdx],
          diagnosedInForecast: beforeRow[diagnosedIdx],

        });
      }

      if (
        statusIdx !== -1 &&
        changedFields.includes(STATUS_COL_NAME) &&
        (norm(afterRow[statusIdx]) === "Closed Lost" || norm(afterRow[statusIdx]) === "Closed Won - Cancelled")
      ) {
        closedLost.push({
          key,
          owner: ownerIdx !== -1 ? afterRow[ownerIdx] : "",
          name: nameIdx !== -1 ? afterRow[nameIdx] : "",
          prevPrice: acpIdx !== -1 ? beforeRow[acpIdx] : "",
          currPrice: acpIdx !== -1 ? afterRow[acpIdx] : "",
          prevStatus: beforeRow[statusIdx],
          currStatus: afterRow[statusIdx],
          diagnosedInForecast: beforeRow[diagnosedIdx],
        });
      }
    }
  }

  //removed
  for (const [key, prevRow] of prevMap.entries()) {
    if (!currMap.has(key)) {
      removed.push({ key, row: project_row(prevRow, commonHeaders, prevIdx) });
    }
  }

  return { commonHeaders, added, removed, modified, closedWon, closedLost };
}

function build_csv_map(rows, keyIdx) {
  const map = new Map();

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const key = String(row[keyIdx] || "").trim();
    if (!key) continue;
    map.set(key, row);
  }

  return map;
}

function project_row(row, headers, idxMap) {
  return headers.map(h => row[idxMap[h]]);
}

function norm(v) {
  if (v === null || v === undefined) return "";
  return String(v).trim();
}

/** COMPARE METRICS */
function diffMetrics(prevMetrics, currMetrics) {
  const rows = [];

  METRICS.forEach(m => {
    const prevDisplay = prevMetrics[m.name];
    const currDisplay = currMetrics[m.name];

    if (prevDisplay === undefined || currDisplay === undefined) return;

    if (norm_(prevDisplay) === norm_(currDisplay)) return;

    let changedBy = "";
    if (m.format === "percentage") {
      const p = parsePercent_(prevDisplay);
      const c = parsePercent_(currDisplay);

      if (p !== null && c !== null) {
        const deltaPP = c - p;
        const sign = deltaPP > 0 ? "+" : "";
        changedBy = `${sign}${round2_(deltaPP)} pp`;
      }
    } else {
      const pNum = parseNumber_(prevDisplay);
      const cNum = parseNumber_(currDisplay);
      if (pNum !== null && cNum !== null) {
        const delta = cNum - pNum;
        const sign = delta > 0 ? "+" : "";
        changedBy = `${sign}${round2_(delta)}`;
      }
    }

    rows.push([
      m.name,
      changedBy,
      prevDisplay,
      currDisplay
    ]);
  });

  return rows;
}

function norm_(s) {
  return String(s || "").replace(/\s+/g, "").replace(/,/g, "");
}

function formatMetricChange_(m, prevDisplay, currDisplay) {
  if (m.format === "percent") {
    const p = parsePercent_(prevDisplay);
    const c = parsePercent_(currDisplay);
    if (p === null || c === null) return `${m.name} changed (${prevDisplay} → ${currDisplay})`;

    const deltaPP = c - p; // percentage points
    const sign = deltaPP > 0 ? "+" : "";
    return `${m.name} changed by ${sign}${round2_(deltaPP)} pp (${prevDisplay} → ${currDisplay})`;
  }

  const pNum = parseNumber_(prevDisplay);
  const cNum = parseNumber_(currDisplay);
  if (pNum !== null && cNum !== null) {
    const delta = cNum - pNum;
    const sign = delta > 0 ? "+" : "";
    return `${m.name} changed by ${sign}${round2_(delta)} (${prevDisplay} → ${currDisplay})`;
  }

  return `${m.name} changed (${prevDisplay} → ${currDisplay})`;
}

function parsePercent_(s) {
  const t = String(s || "").trim();
  if (!t) return null;

  if (t.endsWith("%")) {
    const n = Number(t.replace("%", "").trim());
    return Number.isFinite(n) ? n : null;
  }

  const n = Number(t);
  if (!Number.isFinite(n)) return null;
  if (n >= 0 && n <= 1) return n * 100; // treat 0.57 as 57%
  return n;
}

function parseNumber_(s) {
  const t = String(s || "").replace(/[$,]/g, "").trim();
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

function round2_(n) {
  return Math.round(n * 100) / 100;
}


/** CLEAN AND WRITE REPORT OUTPUT */
function write_diff_report(curr_sheet, prevName, currName, diff) {
  const sh = curr_sheet.getSheetByName(REPORT_SHEET_NAME) || curr_sheet.insertSheet(REPORT_SHEET_NAME);
  sh.clear();

  const headerRow = ["ChangeType", "ChangedFields", ...diff.commonHeaders];

  const meta = [
    ["Generated", new Date().toString()],
    ["Compared", `${prevName}  →  ${currName}`],
    ["", ""],
    ["Summary", ""],
    ["Modified", diff.modified.length],
    ["", ""],
  ];

  sh.getRange(1, 1, meta.length, 2).setValues(meta);

  let row = meta.length + 1;

  const writeSection = (title, headers, rows) => {
    sh.getRange(row, 1).setValue(title);
    row += 1;
    sh.getRange(row, 1, 1, headers.length).setValues([headers]);
    row += 1;

    if (rows.length) {
      sh.getRange(row, 1, rows.length, headers.length).setValues(rows);
      row += rows.length;
    } else {
      sh.getRange(row, 1, 1, headers.length).setValues([headers.map(_ => "")]);
      row += 1;
    }
    row += 2; 
  };

  /** // Added
  writeSection(
    "ADDED (current version)",
    diff.commonHeaders,
    diff.added.map(x => x.row)
  );

  // Removed
  writeSection(
    "REMOVED (previous version)",
    diff.commonHeaders,
    diff.removed.map(x => x.row)
  ); */

  // Modified (BEFORE/AFTER rows)
  const modHeaders = ["ChangedFields", "Version", ...diff.commonHeaders];
  const modRows = diff.modified.flatMap(m => ([
    [m.changedFields.join(", "), "BEFORE", ...m.beforeRow],
    [m.changedFields.join(", "), "AFTER",  ...m.afterRow],
    ["", "", ...diff.commonHeaders.map(_ => "")], // spacer line between records
  ]));

  writeSection(
    "MODIFIED (before/after)",
    modHeaders,
    modRows
  );

  // Make it readable
  //sh.setFrozenRows(meta.length); // keep meta visible (optional)
  //sh.autoResizeColumns(1, Math.min(15, modHeaders.length));

  return curr_sheet.getUrl();
}

function write_diff_metrics_to_report(curr_sheet, prevName, currName, metricRows) {
  const sh = curr_sheet.getSheetByName(METRICS_REPORT_NAME) || curr_sheet.insertSheet(METRICS_REPORT_NAME);
  sh.clear();

  const meta = [
    ["Metric changes"],
    ["Compared", `${prevName} → ${currName}`],
    ["Generated", new Date().toString()],
    [""],
  ];

  sh.getRange(1, 1, meta.length, 2).setValues(
    meta.map(r => [r[0], r[1] || ""])
  );

  const header = ["Metric", "Changed By", "Prev Value", "Curr Value"];
  const startRow = meta.length + 1;

  sh.getRange(startRow, 1, 1, header.length).setValues([header]);

  if (metricRows.length) {
    sh.getRange(startRow + 1, 1, metricRows.length, header.length)
      .setValues(metricRows);
  } else {
    sh.getRange(startRow + 1, 1).setValue("(no metric changes)");
  }

}


/** STEP 4: SEND EMAIL */
function sendChangesEmail_(curr_sheet, closedWon, closedLost, metricRows, prevName, currName, reportUrl, recipients) {
  const subject = `2026 MS Weekly Updates`;
  const emailMetricRows = filterMetricRowsForEmail_(metricRows);

  const now = new Date();
  const qi = getQuarterInfo_(now);
  const currentQ = qi.q;
  const includeNextQ = now >= qi.midpoint;
  const nextQ = qi.nextQuarter;

  const currRR = [curr_sheet.getRangeByName(`q${currentQ}_revenue_retention`).getDisplayValue()];
  const currML = [curr_sheet.getRangeByName(`q${currentQ}_most_likely`).getDisplayValue()];
  const currBC = [curr_sheet.getRangeByName(`q${currentQ}_best_case`).getDisplayValue()];
  const nextRR = [curr_sheet.getRangeByName(`q${nextQ}_revenue_retention`).getDisplayValue()];
  const nextML = [curr_sheet.getRangeByName(`q${nextQ}_most_likely`).getDisplayValue()];
  const nextBC = [curr_sheet.getRangeByName(`q${nextQ}_best_case`).getDisplayValue()];
 
  metricRows.forEach(r => {
    if (r[0] === (`Q${currentQ} Revenue Retention`)) { currRR.push(r[1], r[2])}
    if (r[0] === (`Q${currentQ} Revenue Retention (Most Likely)`)) { currML.push(r[1], r[2])}
    if (r[0] === (`Q${currentQ} Revenue Retention (Best Case)`)) { currBC.push(r[1], r[2])}
    if (r[0] === (`Q${nextQ} Revenue Retention`)) { nextRR.push(r[1], r[2])}
    if (r[0] === (`Q${nextQ} Revenue Retention (Most Likely)`)) { nextML.push(r[1], r[2])}
    if (r[0] === (`Q${nextQ} Revenue Retention (Best Case)`)) { nextBC.push(r[1], r[2])}
  });
  // Logger.log("currRR: %s\ncurrML: %s\ncurrBC: %s\nnextRR: %s\nnextML: %s\nnextBC: %s", currRR, currML, currBC, nextRR, nextML, nextBC);

  const closedWonHtml = (closedWon && closedWon.length)
    ? `
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">
        <tr>
          <th align="left">Account Owner</th>
          <th align="left">Opportunity Name</th>
          <th align="left">Renewed Price</th>
          <th align="left">Diagnosed in Forecast?</th>
        </tr>
        ${closedWon.map(x => `
          <tr>
            <td>${escapeHtml_(cleanForEmailStrong_(x.owner))}</td>
            <td>${escapeHtml_(cleanForEmailStrong_(x.name))}</td>
            <td>${escapeHtml_(cleanForEmailStrong_(x.currPrice))}</td>
            <td>${escapeHtml_(cleanForEmailStrong_(x.diagnosedInForecast))}</td>
          </tr>
        `).join("")}
      </table>
    `
    : `<p>No changes from last week.</p>`;

  const closedLostHtml = (closedLost && closedLost.length)
    ? `
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">
        <tr>
          <th align="left">Account Owner</th>
          <th align="left">Opportunity Name</th>
          <th align="left">Lost Revenue</th>
          <th align="left">Diagnosed in Forecast?</th>
        </tr>
        ${closedLost.map(x => `
          <tr>
            <td>${escapeHtml_(cleanForEmailStrong_(x.owner))}</td>
            <td>${escapeHtml_(cleanForEmailStrong_(x.name))}</td>
            <td>${escapeHtml_(cleanForEmailStrong_(x.currPrice))}</td>
            <td>${escapeHtml_(cleanForEmailStrong_(x.diagnosedInForecast))}</td>
          </tr>
        `).join("")}
      </table>
    `
    : `<p>No changes from last week.</p>`;

  const kpiHtml = (emailMetricRows.length)
    ? `<ul>${emailMetricRows.map(r => {
        const metric = escapeHtml_(String(r[0] ?? ""));
        const changedBy = String(r[1] ?? "").trim();
        const prevVal = escapeHtml_(String(r[2] ?? ""));
        const currVal = escapeHtml_(String(r[3] ?? ""));
        const tail = changedBy ? `${escapeHtml_(changedBy)} (${prevVal} → <strong>${currVal}</strong>)` : `(${prevVal} → <strong>${currVal}</strong>)`;
        return `<li>${metric}: ${tail}</li>`;
      }).join("")}</ul>`
    : `<p>No changes from last week.</p>`;

  const qtrHtml = `
    <h3>Q${currentQ} Revenue Retention</h3>
    <ul>
      <li>Actual: <strong>${currRR[0]}</strong> ${currRR.length > 1 ? `(${currRR[1]} from ${currRR[2]})` : ``}</li>
      <li>Most Likely: <strong>${currML[0]}</strong> ${currML.length > 1 ? `(${currML[1]} from ${currML[2]})` : ``}</li>
      <li>Best Case: <strong>${currBC[0]}</strong> ${currBC.length > 1 ? `(${currBC[1]} from ${currBC[2]})` : ``}</li>
    </ul>
  `;

  /**
  ${includeNextQ ? `
      <h3>Upcoming: Q${nextQ} Revenue Retention</h3>
      <ul>
        <li>Actual: <strong>${nextRR[0]}</strong> ${nextRR.length > 1 ? `(${nextRR[1]} from ${nextRR[2]})` : ``}</li>
        <li>Most Likely: <strong>${nextML[0]}</strong> ${nextML.length > 1 ? `(${nextML[1]} from ${nextML[2]})` : ``}</li>
        <li>Best Case: <strong>${nextBC[0]}</strong> ${nextBC.length > 1 ? `(${nextBC[1]} from ${nextBC[2]})` : ``}</li>
      </ul>
      ` : ``}
   */

  const htmlBody = `
    <p>Weekly changes report is ready.</p> <br>

    <h3>Closed Won Updates</h3>
    ${closedWonHtml} <br>

    <h3>Closed Lost Updates</h3>
    ${closedLostHtml} <br>

    ${qtrHtml} <br>

    <h3>Annual Revenue Retention</h3>
    ${kpiHtml} <br>

    <p><a href="${escapeHtml_(reportUrl)}">Open the report</a></p>
  `;

  MailApp.sendEmail({
    to: recipients.join(","),
    subject,
    htmlBody
  });
}

// Strong cleaner: removes all breaks + control chars (keeps normal spaces)
function cleanForEmailStrong_(v) {
  return String(v ?? "")
    .replace(/[\r\n\u2028\u2029\u0085]/g, " ")     // common hard line separators
    .replace(/[\x00-\x1F\x7F]/g, " ")             // control chars
    .replace(/[\s\u00A0]+/g, " ")                 // collapse whitespace incl NBSP
    .trim();
}

function escapeHtml_(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function filterMetricRowsForEmail_(metricRows) {
  if (!EMAIL_METRIC_WHITELIST || EMAIL_METRIC_WHITELIST.length === 0) return metricRows;
  const allowed = new Set(EMAIL_METRIC_WHITELIST);
  return metricRows.filter(r => allowed.has(r[0])); // r[0] is Metric name
}

function formatMetricRowsAsText_(metricRows) {
  // Plain text (email-safe)
  // Example line: Revenue retention: +7 pp (50% → 57%)
  return metricRows.map(r => `${r[0]}: ${r[1] || "(changed)"} (${r[2]} → ${r[3]})`).join("\n");
}

function getQuarterInfo_(d) {
  const month = d.getMonth(); // 0-11
  const q = Math.floor(month / 3) + 1;

  const startMonth = (q - 1) * 3;
  const start = new Date(d.getFullYear(), startMonth, 1);
  const end = new Date(d.getFullYear(), startMonth + 3, 1); // quarter end (exclusive)

  const midpoint = new Date(start.getTime() + (end.getTime() - start.getTime()) / 2);

  const nextQuarter = q === 4 ? 1 : q + 1;
  const nextQuarterYear = q === 4 ? d.getFullYear() + 1 : d.getFullYear();

  return { q, start, end, midpoint, nextQuarter, nextQuarterYear };
}
