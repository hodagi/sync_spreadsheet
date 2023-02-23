function main2() {

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

    // rangeの右下のセル
    let numRows = sheet.getLastRow() - 1;
    let numColumns = 6;

    if (numRows <= 0) {
        // エラー処理
        Logger.log("指定した範囲に値がありません。");
        return;
    }

    // スプレッドシートからカレンダーに転記するデータを取得
    let range = sheet.getRange(2, 2, numRows, numColumns);
    let data = range.getValues();
    if (data.length === 0 || data[0].length === 0) {
        // エラー処理
        Logger.log("指定した範囲に値がありません。");
        return;
    }

    // カレンダーの取得
    let calendar = CalendarApp.getCalendarById(calendarId);

    // スプレッドシートの各行のデータをカレンダーに転記
    data.forEach(row => {
        const [title, startDate, endDate, description, participant, event_id] = row;

        if (endDate === null || endDate === "" || startDate == endDate) {
            isAllDay = true;
        } else {
            isAllDay = false;
        }
        const start = new Date(startDate);
        const end = new Date(endDate);
        const option = {
            "description": description,
        }

        // イベントIDがすでにスプレッドシートにある場合は更新する
        if(event_id !== ""){
          action = "update"
        }else{
          action = "create"
        }

        let event;
        let eventArray = []
        switch (action) {
            case "update":
                // 登録済みの場合、カレンダーのイベントを更新する
                event = calendar.getEventById(event_id);
                if(event === null){
                  // イベントIDが不正なので予定を削除する
                  let rowIndex = data.indexOf(row) + 2;
                  sheet.deleteRow(rowIndex);
                  return;
                }
                event.setTitle(title);
                if(isAllDay){
                  event.setAllDayDate(start)
                }else{
                  event.setTime(start, end);
                }
                event.setDescription(description);
                break;
            case "create":
                // 未登録の場合、カレンダーに登録する
                if (isAllDay) {
                    event = calendar.createAllDayEvent(title, start, option);
                } else {
                    event = calendar.createEvent(title, start, end, option);
                }
                // イベントIDを取得する
                update_event_id = event.getId();
                // スプレッドシートにイベントIDを設定する
                const columnIndex = 7;
                // イベントIDが未設定の行を取得する
                const rowIndex = data.indexOf(row) + 2;
                sheet.getRange(rowIndex, columnIndex).setValue(update_event_id);
            default:
                break;
        }
    });

    numRows = sheet.getLastRow() - 1;
    if (numRows <= 0) {
        // エラー処理
        Logger.log("指定した範囲に値がありません。");
        return;
    }

    // 最後にカレンダーからイベント情報をまとめて取得する
    const events = calendar.getEvents(new Date(), new Date("2030"));
    const eventIdsInCalendar = events.map(event => event.getId());

    // 今更新したスプレッドシートからイベントIDのリストを取得する
    const eventIdsInSheet = sheet.getRange(2, 6, numRows, 6).getValues().flat();
    
    //  カレンダーにはあるが、シートには存在しないイベントIDを削除する
    const eventIdsToDelete = eventIdsInCalendar.filter(id => !eventIdsInSheet.includes(id));
    eventIdsToDelete.forEach(id => {
      const event = calendar.getEventById(id);
      event.deleteEvent();
    });
}
