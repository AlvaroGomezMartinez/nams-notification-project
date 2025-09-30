/**
 * NAMS NOTIFICATION PROJECT (Restroom Log)
 * 
 * This project logs student restroom usage and manages the associated data.
 * It provides a user interface for logging restroom visits and a backend for data storage and retrieval.
 * The system aims to streamline the process of tracking student movements and ensuring their safety.
 * 
 * Key Features:
 * - Log student restroom usage with time stamps.
 * - Differentiate between AM and PM logs.
 * - Archive logs into a central database for record-keeping.
 * - Provide a user-friendly sidebar interface for teachers to log entries.
 * - Ensure data integrity by validating student IDs against a daily tracking list.
 * - Notify teachers when a student has exceeded a predefined number of restroom visits (2 times in AM or PM).
 * - Customizable teacher identification based on email.
 * - Error handling and logging for debugging and maintenance.
 * - Menu integration within Google Sheets for easy access to functionalities (log students and archive data).
 * - Diagnostic tools to inspect and verify sheet configurations.
 * - Designed for use within Google Sheets using Google Apps Script.
 * - Timezone-aware date and time handling.
 * 
 * @author Alvaro Gomez, Academic Technology Coach, 210-397-9408
 * @version 1.0, 2025-09-30
 */


/**
 * Moves all rows from the AM and PM log sheets into the `Database` sheet and clears the logs.
 * If the `Database` sheet does not exist it will be created with a header row.
 * This function preserves header rows in the AM/PM sheets and only transfers non-empty rows.
 *
 * No parameters.
 * @return {void}
 */
function archiveAndClearLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dbSheet = ss.getSheetByName("Database");
  if (!dbSheet) {
    dbSheet = ss.insertSheet("Database");
    dbSheet.appendRow([
      "Date",
      "Student",
      "ID",
      "Gender",
      "Teacher",
      "Time Out",
      "Time Back",
      "Period",
      "Notes",
    ]);
  }
  var logSheetNames = ["AM", "PM"];
  logSheetNames.forEach(function (sheetName) {
    var logSheet = ss.getSheetByName(sheetName);
    if (logSheet) {
      var lastRow = logSheet.getLastRow();
      var lastCol = logSheet.getLastColumn();
      if (lastRow > 1) {
        var data = logSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
        // Only append non-empty rows
        var nonEmptyRows = data.filter(function (row) {
          return row.join("").trim() !== "";
        });
        if (nonEmptyRows.length > 0) {
          dbSheet
            .getRange(dbSheet.getLastRow() + 1, 1, nonEmptyRows.length, lastCol)
            .setValues(nonEmptyRows);
        }
        // Clear all rows except header
        logSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
      }
    }
  });
}

/**
 * Diagnostic helper you can run from the Apps Script editor.
 * It logs sheet names, visibility, detected day sheet, and a sample of student IDs.
 */
function diagnoseSheets() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var sheetInfo = sheets.map(function (s) {
      return { name: s.getName(), hidden: s.isSheetHidden() };
    });
    Logger.log("Sheets in spreadsheet: %s", JSON.stringify(sheetInfo));

    // Find candidate day sheet the same way logRestroomUsage does
    var daySheet = null;
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (
        name !== "AM" &&
        name !== "PM" &&
        sheets[i].isSheetHidden() === false
      ) {
        daySheet = sheets[i];
        break;
      }
    }
    if (!daySheet) {
      Logger.log("No day sheet detected (no visible sheet other than AM/PM)");
      return { sheets: sheetInfo, daySheet: null };
    }
    Logger.log("Detected day sheet: %s", daySheet.getName());

    var lastRow = daySheet.getLastRow();
    var lastCol = daySheet.getLastColumn();
    Logger.log("Day sheet lastRow=%s lastCol=%s", lastRow, lastCol);

    if (lastRow >= 3) {
      var idRange = daySheet
        .getRange(3, 5, Math.max(0, lastRow - 2), 1)
        .getValues();
      var sampleIds = idRange.slice(0, 10).map(function (r) {
        return r[0];
      });
      Logger.log(
        "Sample IDs (first 10) from column E starting at row 3: %s",
        JSON.stringify(sampleIds)
      );
      return {
        sheets: sheetInfo,
        daySheet: daySheet.getName(),
        sampleIds: sampleIds,
      };
    } else {
      Logger.log("Day sheet has no rows starting at row 3");
      return { sheets: sheetInfo, daySheet: daySheet.getName(), sampleIds: [] };
    }
  } catch (e) {
    Logger.log("Error in diagnoseSheets: %s", e.toString());
    return { error: e.toString() };
  }
}

/**
 * Inspect the spreadsheet and return metadata useful for debugging.
 * It finds visible sheets, detects a candidate "day sheet" (the first visible sheet that is not AM/PM),
 * and returns a short sample of student IDs from column E starting at row 3 when available.
 *
 * @return {{sheets: Array<{name:string,hidden:boolean}>, daySheet: ?string, sampleIds?: Array, error?: string}}
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Restroom Log")
    .addItem("Log Students", "showSidebar")
    .addItem("Move current AM/PM logs to Database", "archiveAndClearLogs")
    .addToUi();
}

/**
 * Adds a custom menu to the Google Sheets UI when the spreadsheet is opened.
 * This enables the sidebar and a manual archive action.
 * @return {void}
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Northside Alternative Middle School")
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Log restroom usage for a student.
 * Expects an object with the shape: { studentId: string, period?: string, gb: 'G'|'B', action: 'Out'|'Back', forceLog?: boolean }
 *
 * On success returns an object: {
 *   confirmationNeeded: boolean,
 *   countBefore: number,
 *   countAfter: number,
 *   studentName: string,
 *   logSheetName: string,
 *   appended: boolean
 * }
 * On error returns { error: string }.
 *
 * @param {{studentId: string, period?: string, gb: string, action: string, forceLog?: boolean}} data
 * @return {{confirmationNeeded?: boolean, countBefore?: number, countAfter?: number, studentName?: string, logSheetName?: string, appended?: boolean, error?: string}}
 */
function logRestroomUsage(data) {
  try {
    Logger.log("logRestroomUsage called with: %s", JSON.stringify(data));
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var nowDate = new Date();
    var today = Utilities.formatDate(
      nowDate,
      Session.getScriptTimeZone(),
      "MM/dd/yy"
    );
    var nowTime = Utilities.formatDate(
      nowDate,
      Session.getScriptTimeZone(),
      "HH:mm:ss"
    );
    var hour = nowDate.getHours();
    var studentId = (data.studentId || "").toString().trim();
    var gb = (data.gb || "").toString().trim(); // 'G' for girl or 'B' for boy
    var action = data.action; // 'Out' or 'Back'
    var forceLog = !!data.forceLog;
    Logger.log(
      "Parsed inputs - studentId: %s, gb: %s, action: %s, forceLog: %s",
      studentId,
      gb,
      action,
      forceLog
    );

    // Teacher email and name mapping
    var email = Session.getActiveUser().getEmail();
    var teacherEmailList = {
      Aguilar: { Email: "russell.aguilar@nisd.net", Salutation: "Mr. " },
      Atoui: { Email: "atlanta.atoui@nisd.net", Salutation: "Mrs." },
      Bowery: { Email: "melissa.bowery@nisd.net", Salutation: "Mrs. " },
      Cantu: { Email: "sandy.cantu@nisd.net", Salutation: "Mrs. " },
      Casanova: { Email: "henry.casanova@nisd.net", Salutation: "Mr. " },
      Coyle: { Email: "deborah.coyle@nisd.net", Salutation: "Mrs. " },
      "De Leon": { Email: "ulices.deleon@nisd.net", Salutation: "Mr. " },
      Farias: { Email: "michelle.farias@nisd.net", Salutation: "Mrs. " },
      Franco: { Email: "george.franco01@nisd.net", Salutation: "Mr." },
      Garcia: { Email: "danny.garcia@nisd.net", Salutation: "Mr. " },
      Goff: { Email: "steven.goff@nisd.net", Salutation: "Mr. " },
      Gomez: { Email: "alvaro.gomez@nisd.net", Salutation: "Mr." },
      Gonzales: { Email: "zina.gonzales@nisd.net", Salutation: "Dr." },
      Hernandez: { Email: "david.hernandez@nisd.net", Salutation: "Mr. " },
      Hutton: { Email: "rebekah.hutton@nisd.net", Salutation: "Mrs. " },
      Idrogo: { Email: "valerie.idrogo@nisd.net", Salutation: "Mrs. " },
      Jasso: { Email: "nadia.jasso@nisd.net", Salutation: "Mrs. " },
      Marquez: { Email: "monica.marquez@nisd.net", Salutation: "Mrs. " },
      Ollendieck: { Email: "reggie.ollendieck@nisd.net", Salutation: "Mr. " },
      Paez: { Email: "john.paez@nisd.net", Salutation: "Mr. " },
      Ramon: { Email: "israel.ramon@nisd.net", Salutation: "Mr. " },
      Tellez: { Email: "lisa.tellez@nisd.net", Salutation: "Mrs. " },
      Trevino: { Email: "marcos.trevino@nisd.net", Salutation: "Mr. " },
      Wine: { Email: "stephanie.wine@nisd.net", Salutation: "Mrs. " },
      Yeager: { Email: "sheila.yeager@nisd.net", Salutation: "Mrs. " },
    };
    var teacherName = email;
    for (var key in teacherEmailList) {
      if (teacherEmailList[key].Email.toLowerCase() === email.toLowerCase()) {
        teacherName = teacherEmailList[key].Salutation + key;
        break;
      }
    }

    // Find the current day's sheet (skips hidden sheets, AM/PM)
    var sheets = ss.getSheets();
    var daySheet = null;
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (
        name !== "AM" &&
        name !== "PM" &&
        sheets[i].isSheetHidden() === false
      ) {
        daySheet = sheets[i];
        break;
      }
    }
    if (!daySheet) {
      Logger.log(
        "No current day sheet found among visible sheets. Existing sheets: %s",
        ss
          .getSheets()
          .map(function (s) {
            return s.getName();
          })
          .join(", ")
      );
      throw new Error("No current day sheet found.");
    }
    Logger.log("Using day sheet: %s", daySheet.getName());

    // Get student IDs from the current day's sheet, column E starting at E3
    var idRange = daySheet
      .getRange(3, 5, daySheet.getLastRow() - 2, 1)
      .getValues();
    var studentIds = idRange.map(function (row) {
      return String(row[0]).trim();
    });
    var studentRow = studentIds.indexOf(studentId);
    if (studentRow === -1) {
      // If the student ID is not found, return special value to alert user to enter ID
      Logger.log(
        "Student ID not found in day sheet. studentId: %s, availableIds: %s",
        studentId,
        JSON.stringify(studentIds.slice(0, 10))
      );
      return { error: "Student ID not found. Please enter a valid ID." };
    }

    // Get student name from the current day's sheet, column A (from the same row it found the ID in)
    var studentName = daySheet.getRange(3 + studentRow, 1).getValue();
    Logger.log("Found studentName: %s at row %s", studentName, 3 + studentRow);

    // Determine whether to add the information in the AM or PM sheet
    var logSheetName = hour < 12 ? "AM" : "PM";
    var logSheet = ss.getSheetByName(logSheetName);
    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName);
      logSheet.appendRow([
        "Date",
        "Student",
        "ID",
        "Gender",
        "Teacher",
        "Time Out",
        "Time Back",
        "Period",
        "Notes",
      ]);
    }
    // Ensure header has Date column
    var logHeaders = logSheet
      .getRange(1, 1, 1, Math.max(7, logSheet.getLastColumn()))
      .getValues()[0];
    if (logHeaders[0] !== "Date") {
      logSheet.insertColumnBefore(1);
      logSheet.getRange(1, 1).setValue("Date");
    }
    // Ensure Period header exists at column 8 (H) and Notes is at column 9 (I)
    if (logSheet.getLastColumn() < 9 || logSheet.getRange(1, 8).getValue() !== "Period" || logSheet.getRange(1, 9).getValue() !== "Notes") {
      // Ensure we have at least 9 columns
      while (logSheet.getLastColumn() < 9) {
        logSheet.insertColumnAfter(logSheet.getLastColumn());
      }
      logSheet.getRange(1, 8).setValue("Period");
      logSheet.getRange(1, 9).setValue("Notes");
    }

    // Get current rows in log sheet and initialize lastRow
    var rows = logSheet.getDataRange().getValues();
    var lastRow = null;
    for (var i = rows.length - 1; i > 0; i--) {
      // Indices: 0=Date, 1=Student, 2=ID, 3=Gender, 4=Teacher, 5=Time Out, 6=Time Back
      if (
        String(rows[i][1]).trim() === studentName &&
        String(rows[i][2]).trim() === studentId &&
        String(rows[i][3]).trim() === gb &&
        String(rows[i][4]).trim() === teacherName
      ) {
        lastRow = i + 1;
        break;
      }
    }
    // Count restroom usage for student today (based on Time Out present)
    var count = 0;
    for (var i = 1; i < rows.length; i++) {
      if (
        String(rows[i][1]).trim() === studentName &&
        String(rows[i][2]).trim() === studentId &&
        rows[i][5]
      ) {
        count++;
      }
    }
    Logger.log(
      "logSheet: %s, existingRows: %s, count so far: %s, lastRow: %s",
      logSheet.getName(),
      rows.length,
      count,
      lastRow
    );

    // We'll return the count before appending so the client can decide whether to warn.
    var appended = false;
    var confirmationNeeded = false;
    var countBefore = count;
    var countAfter = count;
    if (action === "Out") {
      // If this would be the 3rd Out (countBefore >= 2) and the client didn't force it,
      // require confirmation and DO NOT append on the server side.
      if (countBefore >= 2 && !forceLog) {
        confirmationNeeded = true;
        appended = false;
        Logger.log(
          "Confirmation needed for Out: student=%s countBefore=%s",
          studentName,
          countBefore
        );
      } else {
        // Only append if under limit or forced
        // Period may be provided in data.period; default to empty string
        var periodValue = (data.period || "").toString();
        var notesValue = (data.notes || "").toString();
        // Columns now: Date, Student, ID, Gender, Teacher, Time Out, Time Back, Period (H), Notes (I)
        var newRow = [
          today,
          studentName,
          studentId,
          gb,
          teacherName,
          nowTime,
          "",
          periodValue,
          notesValue,
        ];
        logSheet.appendRow(newRow);
        appended = true;
        countAfter = countBefore + 1;
        Logger.log(
          "Appended Out row to %s: %s",
          logSheet.getName(),
          JSON.stringify(newRow)
        );
      }
    } else if (action === "Back") {
      // Attempt to update Back time in the current AM/PM log (column 7: Back)
      if (lastRow) {
        logSheet.getRange(lastRow, 7).setValue(nowTime);
        Logger.log(
          "Updated Back time in %s at row %s with %s",
          logSheet.getName(),
          lastRow,
          nowTime
        );
        // If a Period was provided, set it in column 8 if empty
        try {
          var periodValueBack = (data.period || "").toString();
          if (periodValueBack !== "") {
            var existingPeriod = (logSheet.getRange(lastRow, 8).getValue() || "").toString();
            if (!existingPeriod) {
              logSheet.getRange(lastRow, 8).setValue(periodValueBack);
            }
          }
        } catch (pe) {
          // ignore period write errors
        }
        // If notes were provided, append them to column 9 (Notes)
        try {
          var notesToAdd = (data.notes || "").toString();
          if (notesToAdd !== "") {
            var existingNotes = (logSheet.getRange(lastRow, 9).getValue() || "").toString();
            var newNotes = existingNotes ? existingNotes + ' | ' + notesToAdd : notesToAdd;
            logSheet.getRange(lastRow, 9).setValue(newNotes);
          }
        } catch (ne) {
          // ignore notes write errors
        }
      } else {
        // Fallback: if no matching row in the current sheet, search the other
        // AM/PM sheet for the most recent matching Out (Time Out present and Time Back empty)
        var otherSheetName = logSheetName === "AM" ? "PM" : "AM";
        var otherSheet = ss.getSheetByName(otherSheetName);
        if (otherSheet) {
          var otherRows = otherSheet.getDataRange().getValues();
          var foundRow = null;
          // iterate from bottom to top to find the most recent matching Out with empty Back
          for (var j = otherRows.length - 1; j > 0; j--) {
            if (
              String(otherRows[j][1]).trim() === studentName &&
              String(otherRows[j][2]).trim() === studentId &&
              String(otherRows[j][3]).trim() === gb &&
              String(otherRows[j][4]).trim() === teacherName &&
              otherRows[j][5] && // Time Out present
              !otherRows[j][6] // Time Back empty
            ) {
              foundRow = j + 1; // sheet rows are 1-indexed
              break;
            }
          }
          if (foundRow) {
              otherSheet.getRange(foundRow, 7).setValue(nowTime);
              // If Period/Notes provided, write them (append Notes)
              try {
                var periodValueFallback = (data.period || "").toString();
                var notesValueFallback = (data.notes || "").toString();
                if (periodValueFallback !== "") {
                  var existingPeriodOther = (otherSheet.getRange(foundRow, 8).getValue() || "").toString();
                  if (!existingPeriodOther) {
                    otherSheet.getRange(foundRow, 8).setValue(periodValueFallback);
                  }
                }
                if (notesValueFallback !== "") {
                  var existingNotesOther = (otherSheet.getRange(foundRow, 9).getValue() || "").toString();
                  var newNotesOther = existingNotesOther ? existingNotesOther + ' | ' + notesValueFallback : notesValueFallback;
                  otherSheet.getRange(foundRow, 9).setValue(newNotesOther);
                }
              } catch (pe) {
                // ignore write errors
              }
            Logger.log(
              "Fallback: Updated Back time in %s at row %s with %s",
              otherSheetName,
              foundRow,
              nowTime
            );
            // adjust logSheetName/logSheet to reflect where we updated for the response
            logSheetName = otherSheetName;
            logSheet = otherSheet;
            appended = false;
          } else {
            Logger.log(
              "No matching Out row found in either %s or %s for Back action",
              logSheetName,
              otherSheetName
            );
          }
        } else {
          Logger.log("Other sheet %s does not exist", otherSheetName);
        }
      }
    }

    return {
      confirmationNeeded: confirmationNeeded,
      countBefore: countBefore,
      countAfter: countAfter,
      studentName: studentName,
      logSheetName: logSheetName,
      appended: appended,
    };
  } catch (e) {
    // Log the full error to the Apps Script execution log and return an error message to the client
    Logger.log("Error in logRestroomUsage: " + e.toString());
    try {
      Logger.log(e.stack);
    } catch (ee) {}
    return { error: "Server error: " + (e.message || e.toString()) };
  }
}
