function doGet(e) {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('menu')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return htmlOutput;
}

function processForm(formObject) {
    try {
        var ss = SpreadsheetApp.openById('1Gb-7YRGhWFvBm7j-vAETVG5R6xwWdByadmAHw19ofak');
        var sheet = ss.getSheetByName('Sheet1');
        
        var tasksString = Array.isArray(formObject.tasks) ? formObject.tasks.join(', ') : formObject.tasks;
        
        if (formObject.otherTasks) {
            tasksString += (tasksString ? ', ' : '') + formObject.otherTasks;
        }

        sheet.appendRow([
            formObject.name,
            formObject.userNameInput,
            tasksString,
            formObject.details,
            formObject.startDate,
            formObject.startTime,
            formObject.endTime
        ]);
        
        return HtmlService.createHtmlOutputFromFile('success').getContent();
    } catch (error) {
        Logger.log('Error processing form: ' + error.message);
        return HtmlService.createHtmlOutput('An error occurred while processing the form. Please try again.');
    }
}

function getFormPage() {
    return HtmlService.createHtmlOutputFromFile('schedule_form').getContent();
}

function findUsers(input) {
    try {
        var ss = SpreadsheetApp.openById('1Gb-7YRGhWFvBm7j-vAETVG5R6xwWdByadmAHw19ofak');
        var sheet = ss.getSheetByName('sheet2');
        var lastRow = sheet.getLastRow();
        var range = sheet.getRange(1, 1, lastRow, 1);
        var data = range.getValues();
        var filteredUsers = data.filter(function(row) {
            return row[0].toLowerCase().includes(input.toLowerCase());
        }).map(function(row) {
            return row[0];
        });

        return filteredUsers;
    } catch (error) {
        Logger.log('Error finding users: ' + error.message);
        return [];
    }
}

function getMenuPage() {
    return HtmlService.createHtmlOutputFromFile('menu').getContent();
}

function getDailyReportListPage() {
    return HtmlService.createHtmlOutputFromFile('daily_report_list').getContent();
}

function getDailyReports() {
    try {
        var ss = SpreadsheetApp.openById('1Gb-7YRGhWFvBm7j-vAETVG5R6xwWdByadmAHw19ofak');
        var sheet = ss.getSheetByName('Sheet1');
        var data = sheet.getDataRange().getValues();
        var currentMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
        var filteredData = data.filter(function(row) {
            var reportDate = Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), 'yyyy-MM');
            return reportDate === currentMonth;
        });
        return filteredData;
    } catch (error) {
        Logger.log('Error getting daily reports: ' + error.message);
        return [];
    }
}

function editReport(rowIndex, updatedReport) {
    try {
        var ss = SpreadsheetApp.openById('1Gb-7YRGhWFvBm7j-vAETVG5R6xwWdByadmAHw19ofak');
        var sheet = ss.getSheetByName('Sheet1');
        var range = sheet.getRange(rowIndex + 1, 1, 1, sheet.getLastColumn());
        range.setValues([updatedReport]);
    } catch (error) {
        Logger.log('Error editing report: ' + error.message);
    }
}

function deleteReport(rowIndex) {
    try {
        var ss = SpreadsheetApp.openById('1Gb-7YRGhWFvBm7j-vAETVG5R6xwWdByadmAHw19ofak');
        var sheet = ss.getSheetByName('Sheet1');
        sheet.deleteRow(rowIndex + 1);
    } catch (error) {
        Logger.log('Error deleting report: ' + error.message);
    }
}

function refreshDailyReports() {
    return getDailyReports();
}
