function doGet(e) {
    var htmlOutput = HtmlService.createHtmlOutputFromFile('schedule_form')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
    return htmlOutput;
}

function processForm(formObject) {
    var ss = SpreadsheetApp.openById('1Gb-7YRGhWFvBm7j-vAETVG5R6xwWdByadmAHw19ofak');
    var sheet = ss.getSheetByName('Sheet1');
    
    // tasksが配列である場合はカンマで結合し、文字列として保存
    var tasksString = Array.isArray(formObject.tasks) ? formObject.tasks.join(', ') : formObject.tasks;
    
    // otherTasksが提供されていればtasksStringに追加
    if (formObject.otherTasks) {
        tasksString += (tasksString ? ', ' : '') + formObject.otherTasks;
    }

    sheet.appendRow([
        formObject.name,
        formObject.userNameInput,
        tasksString, // 結合されたタスク文字列（tasksとotherTasks）
        formObject.details,
        formObject.startDate,
        formObject.startTime,
        formObject.endTime
    ]);
    
    return HtmlService.createHtmlOutputFromFile('success').getContent();
}



function getFormPage() {
    return HtmlService.createHtmlOutputFromFile('schedule_form').getContent();
}

function findUsers(input) {
    var ss = SpreadsheetApp.openById('1Gb-7YRGhWFvBm7j-vAETVG5R6xwWdByadmAHw19ofak');
    var sheet = ss.getSheetByName('sheet2'); // 利用者名が保存されているシート名を指定
    var lastRow = sheet.getLastRow(); // データのある最後の行を取得
    var range = sheet.getRange(1, 1, lastRow, 1); // A列のみ、最初の行から最後の行までの範囲を指定
    var data = range.getValues(); // 指定範囲のデータを二次元配列で取得
    var filteredUsers = data.filter(function(row) {
        return row[0].toLowerCase().includes(input.toLowerCase());
    }).map(function(row) {
        return row[0]; // 利用者名がA列にあると仮定
    });

    return filteredUsers;
}

