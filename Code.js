/**
 * Moves all data from AM and PM sheets (starting from row 2) to the Database sheet, then clears AM and PM logs except for headers.
 */
function archiveAndClearLogs() {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var dbSheet = ss.getSheetByName('Database');
	if (!dbSheet) {
		dbSheet = ss.insertSheet('Database');
		dbSheet.appendRow(['Date', 'Student', 'ID', 'Gender', 'Teacher', 'Time Out', 'Time Back']);
	}
	var logSheetNames = ['AM', 'PM'];
	logSheetNames.forEach(function(sheetName) {
		var logSheet = ss.getSheetByName(sheetName);
		if (logSheet) {
			var lastRow = logSheet.getLastRow();
			var lastCol = logSheet.getLastColumn();
			if (lastRow > 1) {
				var data = logSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
				// Only append non-empty rows
				var nonEmptyRows = data.filter(function(row) {
					return row.join('').trim() !== '';
				});
				if (nonEmptyRows.length > 0) {
					dbSheet.getRange(dbSheet.getLastRow() + 1, 1, nonEmptyRows.length, lastCol).setValues(nonEmptyRows);
				}
				// Clear all rows except header
				logSheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
			}
		}
	});
}

function onOpen() {
	SpreadsheetApp.getUi()
		.createMenu('Restroom Log')
		.addItem('Open Sidebar', 'showSidebar')
		.addItem('Clear AM/PM Logs and move to Database', 'archiveAndClearLogs')
		.addToUi();
	showSidebar();
}

function showSidebar() {
	var html = HtmlService.createHtmlOutputFromFile('Sidebar')
		.setTitle('NAMS')
		.setWidth(350);
	SpreadsheetApp.getUi().showSidebar(html);
}

function logRestroomUsage(data) {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var nowDate = new Date();
	var today = Utilities.formatDate(nowDate, Session.getScriptTimeZone(), 'MM/dd/yy');
	var nowTime = Utilities.formatDate(nowDate, Session.getScriptTimeZone(), 'HH:mm:ss');
	var hour = nowDate.getHours();
	var studentId = (data.studentId || '').trim();
	var gb = (data.gb || '').trim(); // 'G' for girl or 'B' for boy
	var action = data.action; // 'Out' or 'Back'
	var forceLog = data.forceLog || false;

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
	  Yeager: { Email: "sheila.yeager@nisd.net", Salutation: "Mrs. " }
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
		if (name !== 'AM' && name !== 'PM' && sheets[i].isSheetHidden() === false) {
			daySheet = sheets[i];
			break;
		}
	}
	if (!daySheet) {
		throw new Error('No current day sheet found.');
	}

	// Get student IDs from the current day's sheet, column E starting at E3
	var idRange = daySheet.getRange(3, 5, daySheet.getLastRow() - 2, 1).getValues();
	var studentIds = idRange.map(function(row) { return String(row[0]).trim(); });
	var studentRow = studentIds.indexOf(studentId);
	if (studentRow === -1) {
		// If the student ID is not found, return special value to alert user to enter ID
		return { error: 'Student ID not found. Please enter a valid ID.' };
	}

	// Get student name from the current day's sheet, column A (from the same row it found the ID in)
	var studentName = daySheet.getRange(3 + studentRow, 1).getValue();

		// Determine whether to add the information in the AM or PM sheet
		var logSheetName = hour < 12 ? 'AM' : 'PM';
		var logSheet = ss.getSheetByName(logSheetName);
		if (!logSheet) {
				logSheet = ss.insertSheet(logSheetName);
				logSheet.appendRow(['Date', 'Student', 'ID', 'Gender', 'Teacher', 'Time Out', 'Time Back']);
		}
		// Ensure header has Date column
		var logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
		if (logHeaders[0] !== 'Date') {
			logSheet.insertColumnBefore(1);
			logSheet.getRange(1, 1).setValue('Date');
		}

	// Find last entry for student today
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
		if (action === 'Out') {
			// Count restroom usage for student today (already calculated below as count)
			if (count < 2 || forceLog) {
				var newRow = [today, studentName, studentId, gb, teacherName, nowTime, ''];
				logSheet.appendRow(newRow);
				// Also add to Database sheet
				var dbSheet = ss.getSheetByName('Database');
				if (!dbSheet) {
					dbSheet = ss.insertSheet('Database');
					dbSheet.appendRow(['Date', 'Student', 'ID', 'Gender', 'Teacher', 'Time Out', 'Time Back']);
				}
				dbSheet.appendRow(newRow);
			}
		} else if (action === 'Back' && lastRow) {
			// Update Back time in AM/PM log (column 7: Back)
			logSheet.getRange(lastRow, 7).setValue(nowTime);
			// Also update Back time in Database for this entry (find by date, name, id, teacher, Out time)
			var dbSheet = ss.getSheetByName('Database');
			if (dbSheet) {
				var dbRows = dbSheet.getDataRange().getValues();
				for (var d = dbRows.length - 1; d > 0; d--) {
					if (
						dbRows[d][0] === today &&
						dbRows[d][1] === studentName &&
						dbRows[d][2] === studentId &&
						dbRows[d][3] === gb &&
						dbRows[d][4] === teacherName &&
						dbRows[d][5] === rows[lastRow-1][5] // Out time matches (column 6)
					) {
						dbSheet.getRange(d+1, 7).setValue(nowTime); // Column 7: Back
						break;
					}
				}
			}
		}

	// Count restroom usage for student today
	var count = 0;
	for (var i = 1; i < rows.length; i++) {
		if (String(rows[i][1]).trim() === studentName && String(rows[i][2]).trim() === studentId && rows[i][5]) {
			count++;
		}
	}
	return {
	  count: count,
	  studentName: studentName,
	  logSheetName: logSheetName
	};
}
