
function onOpen() {
	SpreadsheetApp.getUi()
		.createMenu('Restroom Log')
		.addItem('Open Sidebar', 'showSidebar')
		.addToUi();
	showSidebar(); // Automatically show sidebar when spreadsheet opens
}


function showSidebar() {
	var html = HtmlService.createHtmlOutputFromFile('Sidebar')
		.setTitle('Restroom Log')
		.setWidth(350);
	SpreadsheetApp.getUi().showSidebar(html);
}


function logRestroomUsage(data) {
	var ss = SpreadsheetApp.getActiveSpreadsheet();
	var nowDate = new Date();
	var today = Utilities.formatDate(nowDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
	var nowTime = Utilities.formatDate(nowDate, Session.getScriptTimeZone(), 'HH:mm:ss');
	var hour = nowDate.getHours();
	var studentId = (data.studentId || '').trim();
	var gb = (data.gb || '').trim();
	var action = data.action; // 'Out' or 'Back'

	// Teacher email and name mapping
	var email = Session.getActiveUser().getEmail();
	var teacherEmailList = {
	  "Aguilar, R": { Email: "russell.aguilar@nisd.net", Salutation: "Mr. " },
	  Atoui: { Email: "atlanta.atoui@nisd.net", Salutation: "Mrs." },
	  Bowery: { Email: "melissa.bowery@nisd.net", Salutation: "Mrs. " },
	  "Cantu, S": { Email: "sandy.cantu@nisd.net", Salutation: "Mrs. " },
	  Casanova: { Email: "henry.casanova@nisd.net", Salutation: "Mr. " },
	  Coyle: { Email: "deborah.coyle@nisd.net", Salutation: "Mrs. " },
	  "De Leon, U": { Email: "ulices.deleon@nisd.net", Salutation: "Mr. " },
	  "Deleon, R": { Email: "rebeca.deleon@nisd.net", Salutation: "Mrs. " },
	  Farias: { Email: "michelle.farias@nisd.net", Salutation: "Mrs. " },
	  "Franco, G": { Email: "george.franco01@nisd.net", Salutation: "Mr." },
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

	// Find the current day's sheet (visible, not AM/PM)
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

	// Get student IDs from column E starting at E3
	var idRange = daySheet.getRange(3, 5, daySheet.getLastRow() - 2, 1).getValues();
	var studentIds = idRange.map(function(row) { return String(row[0]).trim(); });
	var studentRow = studentIds.indexOf(studentId);
	if (studentRow === -1) {
		// Return special value to alert user to enter ID
		return { error: 'Student ID not found. Please enter a valid ID.' };
	}

	// Get student name from column A (same row as ID)
	var studentName = daySheet.getRange(3 + studentRow, 1).getValue();

	// Determine AM/PM sheet
	var logSheetName = hour < 12 ? 'AM' : 'PM';
	var logSheet = ss.getSheetByName(logSheetName);
	if (!logSheet) {
		logSheet = ss.insertSheet(logSheetName);
		logSheet.appendRow(['Student Name', 'ID', 'G or B', 'Teacher', 'Out', 'Back']);
	}

	// Find last entry for student today
	var rows = logSheet.getDataRange().getValues();
	var lastRow = null;
	for (var i = rows.length - 1; i > 0; i--) {
		if (String(rows[i][0]).trim() === studentName && String(rows[i][1]).trim() === studentId && String(rows[i][2]).trim() === gb && String(rows[i][3]).trim() === teacherName) {
			lastRow = i + 1;
			break;
		}
	}
	if (action === 'Out') {
		logSheet.appendRow([studentName, studentId, gb, teacherName, nowTime, '']);
	} else if (action === 'Back' && lastRow) {
		logSheet.getRange(lastRow, 6).setValue(nowTime);
	}

	// Count restroom usage for student today
	var count = 0;
	for (var i = 1; i < rows.length; i++) {
		if (String(rows[i][0]).trim() === studentName && String(rows[i][1]).trim() === studentId && rows[i][4]) {
			count++;
		}
	}
	return { count: count };
}
