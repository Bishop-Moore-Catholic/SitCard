/**
 * Yearbook Picture Day — QR Scanner Portal (Scans logging)
 * Bound to your Google Sheet.
 */

const SHEET_SCANS = "Scans";

/**
 * Column order:
 * Timestamp, SessionID, Club, Operator, RowNumber, RowPostion, StudentID,
 * First Name, Last Name, Grade, GradYear, QRRawPayload, Status, Notes
 */
const SCANS_HEADERS = [
  "Timestamp",
  "SessionID",
  "Club",
  "Operator",
  "RowNumber",
  "RowPostion",
  "StudentID",
  "First Name",
  "Last Name",
  "Grade",
  "GradYear",
  "QRRawPayload",
  "Status",
  "Notes"
];

// Adjust later if you want stricter 26–29 constraint, etc.
const STUDENT_ID_REGEX = /^20\d{5}$/;

// Prevent spam duplicates server-side (per session+studentId) within window
const DUP_WINDOW_SECONDS = 20;

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Scanner")
    .setTitle("QR Scanner Portal");
}

/**
 * Run once (or anytime) to create/repair the Scans sheet and headers.
 * Does NOT delete existing scan data.
 */
function setupYearbookScannerSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_SCANS);
  if (!sh) sh = ss.insertSheet(SHEET_SCANS);

  // Ensure the sheet has enough columns for the headers
  const neededCols = SCANS_HEADERS.length;
  const currentCols = sh.getMaxColumns();
  if (currentCols < neededCols) {
    sh.insertColumnsAfter(currentCols, neededCols - currentCols);
  }

  // Write headers (row 1 only)
  sh.getRange(1, 1, 1, neededCols)
    .setValues([SCANS_HEADERS])
    .setFontWeight("bold")
    .setBackground("#f1f3f4");

  sh.setFrozenRows(1);

  // Light formatting (optional)
  sh.setColumnWidths(1, neededCols, 140);
  sh.setColumnWidth(12, 360); // QRRawPayload
  sh.setColumnWidth(14, 260); // Notes
  sh.getRange(1, 1, sh.getMaxRows(), neededCols).setWrap(true);
  sh.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");

  SpreadsheetApp.flush();
  return { ok: true, sheet: SHEET_SCANS, columns: neededCols };
}

function getClientConfig() {
  return {
    scansSheetName: SHEET_SCANS,
    studentIdRegex: STUDENT_ID_REGEX.source,
    dupWindowSeconds: DUP_WINDOW_SECONDS
  };
}

/**
 * BATCHED + NO DUP DETECTION
 *
 * Client should call:
 *   recordScansBatch(scansArray)
 *
 * where scansArray is an array of objects like:
 * {
 *   payloadRaw: "...", club: "...", notes: "...",
 *   rowNumber: 1, rowPosition: 12,
 *   sessionId: "SESSION-...", gradYearParam: 2026, gradeParam: 12,
 *   clientId: "optional-id-to-match-results"
 * }
 *
 * This function:
 *  - does NOT check duplicates (scanner handles it)
 *  - appends rows in ONE write (setValues) for speed/quota safety
 *  - still rejects invalid JSON payloads (won't write them)
 *  - still writes INVALID_ID rows (so you can audit)
 */
function recordScansBatch(scansArray) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_SCANS);
    if (!sh) throw new Error("Scans sheet missing. Run setupYearbookScannerSheets().");

    const operator = Session.getActiveUser().getEmail() || "Unknown";
    const now = new Date();

    const scans = Array.isArray(scansArray) ? scansArray : [];
    if (scans.length === 0) {
      return { ok: true, status: "OK", operator, appended: 0, results: [] };
    }

    // Safety guard: keep batches reasonable (client should also limit)
    const MAX_BATCH = 50;
    const batch = scans.slice(0, MAX_BATCH);

    const rowsToAppend = [];
    const results = [];

    for (let i = 0; i < batch.length; i++) {
      const item = batch[i] || {};
      const t = new Date(now.getTime()); // one timestamp per row (same ms is fine)
      const raw = String(item.payloadRaw || "").trim();

      const sid =
        String(item.sessionId || "").trim() ||
        ("SESSION-" + Utilities.getUuid().slice(0, 8).toUpperCase());

      const clubSafe  = String(item.club || "").trim();
      const notesSafe = String(item.notes || "").trim();

      const rowNum = Number(item.rowNumber || 1);
      const rowPos = Number(item.rowPosition || 0);

      const clientId = item.clientId != null ? String(item.clientId) : "";

      // Parse payload (JSON or ID-only)
      let obj;
      let payloadMode = "JSON";
      if (STUDENT_ID_REGEX.test(raw)) {
        payloadMode = "ID_ONLY";
        obj = { id: raw, first: "", last: "", grade: "", gradYear: "" };
      } else {
        try {
          obj = JSON.parse(raw);
        } catch (e) {
          // DO NOT write invalid JSON rows (per your original behavior)
          results.push({
            clientId,
            ok: false,
            status: "INVALID_JSON",
            message: "Invalid QR payload (ignored)."
          });
          continue;
        }
      }

      const studentId = String(obj.id || obj.studentId || obj.sid || "").trim();
      const firstName = String(obj.first || obj.firstName || obj.f || "").trim();
      const lastName  = String(obj.last  || obj.lastName  || obj.l || "").trim();

      const gradeVal    = (item.gradeParam ?? obj.grade ?? obj.Grade ?? "");
      const gradYearVal = (item.gradYearParam ?? obj.gradYear ?? obj.graduationYear ?? obj.grad_year ?? "");

      let grade = String(gradeVal).trim();
      let gradYear = String(gradYearVal).trim();

      // If missing, infer from StudentID prefix (first 4 digits = gradYear)
      if (!gradYear && /^\d{4}/.test(studentId)) {
        gradYear = studentId.slice(0, 4);
      }

      // If grade still missing and gradYear is valid, infer grade for current school year
      if (!grade && /^\d{4}$/.test(gradYear)) {
        const gy = Number(gradYear);
        const y = now.getFullYear();
        const mo = now.getMonth(); // 0=Jan
        const schoolYearEnds = (mo >= 6) ? (y + 1) : y; // July rollover
        const g = 12 - (gy - schoolYearEnds);
        if (g >= 9 && g <= 12) grade = String(g);
      }

      // Validate StudentID (we still WRITE the row with INVALID_ID so you can audit)
      if (!STUDENT_ID_REGEX.test(studentId)) {
        rowsToAppend.push([
          t, sid, clubSafe, operator, rowNum, rowPos,
          studentId, firstName, lastName, grade, gradYear,
          raw, "INVALID_ID", notesSafe
        ]);
        results.push({
          clientId,
          ok: false,
          status: "INVALID_ID",
          message: "StudentID failed validation.",
          student: { studentId, firstName, lastName, grade, gradYear },
          payloadMode
        });
        continue;
      }

      // OK row (NO DUP CHECK)
      rowsToAppend.push([
        t, sid, clubSafe, operator, rowNum, rowPos,
        studentId, firstName, lastName, grade, gradYear,
        raw, "OK", notesSafe
      ]);

      results.push({
        clientId,
        ok: true,
        status: "OK",
        student: { studentId, firstName, lastName, grade, gradYear },
        payloadMode
      });
    }

    // One append for the whole batch
    if (rowsToAppend.length) {
      const startRow = sh.getLastRow() + 1;
      sh.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }

    return {
      ok: true,
      status: "OK",
      operator,
      received: scans.length,
      processed: batch.length,
      appended: rowsToAppend.length,
      results
    };
  } finally {
    lock.releaseLock();
  }
}


/**
 * Backward-compatible single-scan entry point.
 * Uses the batched path (still no dup detection).
 */
function recordScan(payloadRaw, club, notes, rowNumber, rowPosition, sessionId, gradYearParam, gradeParam) {
  const res = recordScansBatch([{
    payloadRaw, club, notes,
    rowNumber, rowPosition,
    sessionId, gradYearParam, gradeParam,
    clientId: "" // optional
  }]);

  // Mirror your old return shape as closely as possible:
  const r0 = (res && res.results && res.results[0]) ? res.results[0] : null;

  if (!r0) return { ok: true, status: "OK", operator: res.operator, message: "No-op." };

  if (r0.status === "OK") {
    return {
      ok: true,
      status: "OK",
      operator: res.operator,
      student: r0.student,
      payloadMode: r0.payloadMode
    };
  }

  // INVALID_JSON is not written; INVALID_ID is written
  return {
    ok: false,
    status: r0.status,
    message: r0.message || r0.status,
    operator: res.operator,
    student: r0.student || null,
    payloadMode: r0.payloadMode
  };
}
