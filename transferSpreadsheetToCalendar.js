// スプレッドシートの内容をカレンダーに転記する
function transferSpreadsheetToCalendar() {

    // ジャガー用カレンダーID
    calendarId_jaguer = ScriptProperties.getProperty('CALENDAR_ID_JAGUER');

    // GoogleCloudEvent投稿用カレンダーID
    calendarId_post = ScriptProperties.getProperty('CALENDAR_ID_CCOE');

    // GoogleCloudEvent管理用ExcelファイルID
    spreadSheetId = ScriptProperties.getProperty('SPREAD_SHEET_ID_CCOE');

    // スプレッドシート名
    spreadSheetName = ScriptProperties.getProperty('SPREAD_SHEET_NAME_CCOE');

    // スプレッドシートの取得
    let sheet = SpreadsheetApp.openById(spreadSheetId).getSheetByName(sheetName);


    const lastRow = sheet.getLastRow() - 1;

    // スプレッドシートからカレンダーに転記するデータを取得
    //let data = sheet.getDataRange().getValues();
    let data = sheet.getRange(2, 1, lastRow, 4).getValues();

    // カレンダーの取得
    let calendar = CalendarApp.getCalendarById(calendarId);

    // スプレッドシートの各行のデータをカレンダーに転記
    data.slice(1).forEach(row => {
        const [title, startDate, endDate, description] = row;
        const start = new Date(startDate);
        const end = new Date(endDate);
        const event = {
            "description": description,
        }
        calendar.createEvent(title, { start, end }, event);
    });
}
