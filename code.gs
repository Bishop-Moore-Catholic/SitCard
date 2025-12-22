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
 * Record a scan with LOCAL counters from the client.
 *
 * payloadRaw can be:
 *  - JSON string {"sid":"20xxxxx","f":"...","l":"..."} (lean QR) OR {"id":"20xxxxx","first":"...","last":"..."}  (from QR)
 *  - ID-only "20xxxxx" (manual)
 *
 * rowNumber/rowPosition/sessionId/club come from the client.
 */
function recordScan(payloadRaw, club, notes, rowNumber, rowPosition, sessionId, gradYearParam, gradeParam) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_SCANS);
    if (!sh) throw new Error("Scans sheet missing. Run setupYearbookScannerSheets().");

    const operator = Session.getActiveUser().getEmail() || "Unknown";
    const ts = new Date();

    const raw = String(payloadRaw || "").trim();
    const sid = String(sessionId || "").trim() || ("SESSION-" + Utilities.getUuid().slice(0, 8).toUpperCase());
    const clubSafe = String(club || "").trim();
    const notesSafe = String(notes || "").trim();
    const rowNum = Number(rowNumber || 1);
    const rowPos = Number(rowPosition || 0);

    // Parse payload (JSON or ID-only)
    let obj;
    let payloadIsIdOnly = false;

    if (STUDENT_ID_REGEX.test(raw)) {
      payloadIsIdOnly = true;
      obj = { id: raw, first: "", last: "", grade: "", gradYear: "" };
    } else {
      try {
        obj = JSON.parse(raw);
      } catch (e) {
        // Do NOT write invalid JSON to the sheet
        // Client will handle position rollback
        return { ok: false, status: "INVALID_JSON", message: "Invalid QR payload (ignored)." };
      }
    }

    const studentId = String(obj.id || obj.studentId || obj.sid || "").trim();
    const firstName = String(obj.first || obj.firstName || obj.f || "").trim();
    const lastName  = String(obj.last  || obj.lastName  || obj.l || "").trim();

    // Accept numeric or string values; store as clean strings
    const gradeVal = (gradeParam ?? obj.grade ?? obj.Grade ?? "");
    const gradYearVal = (gradYearParam ?? obj.gradYear ?? obj.graduationYear ?? obj.grad_year ?? "");

    let grade = String(gradeVal).trim();
    let gradYear = String(gradYearVal).trim();

    // If missing, infer from StudentID prefix (first 4 digits = gradYear)
    if (!gradYear && /^\d{4}/.test(studentId)) {
      gradYear = studentId.slice(0, 4);
    }

    // If grade still missing and gradYear is valid, infer grade for current school year
    if (!grade && /^\d{4}$/.test(gradYear)) {
      const gy = Number(gradYear);
      const now = new Date();
      const y = now.getFullYear();
      const mo = now.getMonth(); // 0=Jan
      const schoolYearEnds = (mo >= 6) ? (y + 1) : y; // July rollover
      const g = 12 - (gy - schoolYearEnds);
      if (g >= 9 && g <= 12) grade = String(g);
    }

    if (!STUDENT_ID_REGEX.test(studentId)) {
      sh.appendRow([ts, sid, clubSafe, operator, rowNum, rowPos, studentId, firstName, lastName, grade, gradYear, raw, "INVALID_ID", notesSafe]);
      return { ok: false, status: "INVALID_ID", message: "StudentID failed validation." };
    }

    // Server-side duplicate protection (session + studentId)
    const cache = CacheService.getScriptCache();
    const dupKey = `dup:${sid}:${studentId}`;
    if (cache.get(dupKey)) {
      sh.appendRow([ts, sid, clubSafe, operator, rowNum, rowPos, studentId, firstName, lastName, grade, gradYear, raw, "DUPLICATE", notesSafe]);
      return { ok: true, status: "DUPLICATE", message: "Duplicate scan (recent).", operator };
    }
    cache.put(dupKey, "1", DUP_WINDOW_SECONDS);

    // Write OK row
    sh.appendRow([ts, sid, clubSafe, operator, rowNum, rowPos, studentId, firstName, lastName, grade, gradYear, raw, "OK", notesSafe]);

    return {
      ok: true,
      status: "OK",
      operator,
      session: { SessionID: sid, RowNumber: rowNum, RowPostion: rowPos },
      student: { studentId, firstName, lastName, grade, gradYear },
      payloadMode: payloadIsIdOnly ? "ID_ONLY" : "JSON"
    };
  } finally {
    lock.releaseLock();
  }
}
