function main1() {

    //ジャガーのカレンダーID
    calendarId_jaguer = ScriptProperties.getProperty('CALENDAR_ID_JAGUAR');

    // GoogleCloudEvent投稿用カレンダーID
    calendarId_post = ScriptProperties.getProperty('CALENDAR_ID_CCOE');

    // GoogleCloudEvent管理用ExcelファイルID
    spreadSheetId = ScriptProperties.getProperty('SPREAD_SHEET_ID_CCOE');

    // スプレッドシート名
    spreadSheetName = ScriptProperties.getProperty('SPREAD_SHEET_NAME_JAGUAR');

    // ジャガー用カレンダー用パラメーター
    // grepしたいタイトル名
    let grep = "";
    // 今日からいつまで検索するか？
    let toYear = 5;

    // GoogleCalendarからデータを取得する
    let start = new Date();
    let end = new Date(start.getFullYear() + toYear, start.getMonth(), start.getDate());
    let matchingEvents = loadEvents_part2(calendarId_jaguer, start, end, grep);

    // GoogleSpreadSheetにGooleCalendaのeventデータを書き込む
    writeCalendar(matchingEvents, spreadSheetId, spreadSheetName);
}

// GoogleCalendarからデータを取得する_part2
function loadEvents_part2(calendarId, start, end, grep) {
    var matchingEvents = Calendar.Events.list(calendarId, {
        timeMin: start.toISOString(),
        timeMax: end.toISOString(),
        singleEvents: true,
        orderBy: 'startTime'
    }).items
        .filter(function (event) {
            return (!grep || event.summary.includes(grep));
        })
        .map(function (event) {
            // リンク用ID
            let link = event.htmlLink;
            let event_id = link.split("eid=")[1];
            return {
                id: event.id,
                event_id: event_id,
                start: event.start.dateTime || event.start.date,
                end: event.end.dateTime || event.end.date,
                title: event.summary,
                description: event.description,
                location: event.location,
                guests: event.attendees,
                isAllDayEvent: !event.start.dateTime,
                isRecurringEvent: !!event.recurrence,
                recurrence: event.recurrence,
                color: event.colorId
            };
        });

    return matchingEvents;
}

// GoogleSpreadSheetにGooleCalendaのeventデータを書き込む
function writeCalendar(events, spreadSheetId, sheetName) {
    let ss = SpreadsheetApp.openById(spreadSheetId);
    let sheet = ss.getSheetByName(sheetName);
    let row = 2;

    // 既に登録済みのタイトルを取得する
    let titles = sheet.getRange("B2:B" + sheet.getLastRow()).getValues().flat();

    // eventごとに登録する
    events.forEach(function (event) {
        // プロパティ名
        let no = row - 1;
        let title = event.title;
        let startDate = parseDateTime(event.start);
        let endDate = parseDateTime(event.end);
        let event_id = event.event_id;
        let link = `https://www.google.com/calendar/event?eid=${event_id}`;
        let linkFormula = `=HYPERLINK("${link}","${title}")`;

        // 重複チェック
        if (titles.includes(title)) {
            // タイトルが一致する行がある場合はスキップする
            return;
        }

        // シートへの書き込み
        sheet.getRange(row, 1).setValue(no);
        sheet.getRange(row, 2).setFormula(linkFormula);
        sheet.getRange(row, 3).setValue(startDate);
        sheet.getRange(row, 4).setValue(endDate);
        row++;
    });
}

// 日付をフォーマットする
// 2020-12-31T15:00:00+09:00 -> 2020/12/31 15:00
function parseDateTime(datetimeString) {
    var date = new Date(datetimeString);
    var formattedDate = Utilities.formatDate(date, "JST", "yyyy/MM/dd HH:mm");
    return formattedDate;
}
