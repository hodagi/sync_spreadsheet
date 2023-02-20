function main() {

    // GoogleCloudEvent投稿用カレンダーID
    calendarId_post = ScriptProperties.getProperty('CALENDAR_ID_CCOE');

    // GoogleCloudEvent管理用ExcelファイルID
    spreadSheetId = ScriptProperties.getProperty('SPREAD_SHEET_ID_CCOE');

    // スプレッドシート名
    spreadSheetName = ScriptProperties.getProperty('SPREAD_SHEET_NAME_CCOE');

    transferSpreadsheetToCalendar(calendarId_post, spreadSheetId, spreadSheetName);
}

// スプレッドシートの内容をカレンダーに転記する
function transferSpreadsheetToCalendar(calendarId, spreadSheetId, spreadSheetName) {

    // スプレッドシートの取得
    let sheet = SpreadsheetApp.openById(spreadSheetId).getSheetByName(spreadSheetName);

    // 一番最後の行を取得する
    // 空白があるかもなのでとりあえず見えてる下を取得する
    const lastRow = sheet.getLastRow();
    //const range = sheet.getRange(lastRow, 1);
    //const nextRow = range.getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
    //const lastDataRow = nextRow - 1;

    const numRows = lastRow - 2 + 1;
    const numColumns = 5;

    // スプレッドシートからカレンダーに転記するデータを取得
    // 2行目の2列目から何行分取得するか...
    let data = sheet.getRange(2, 2, numRows, numColumns).getValues();

    // カレンダーの取得
    let calendar = CalendarApp.getCalendarById(calendarId);

    // スプレッドシートの各行のデータをカレンダーに転記
    data.forEach(row => {
        Logger.log(row)
        const [title, startDate, endDate, description] = row;
        const start = new Date(startDate);
        const end = new Date(endDate);
        const event = {
            "description": description,
        }
        calendar.createEvent(title, start, end, event);
    });
}
