function syncCalendarToSheet() {
  var calendarId = '0a437281002a546d7e17233cefa484100bd212d88da1c792e9f162bdf1be23e3@group.calendar.google.com';
  var sheetId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
  var sheetName = 'CalendarEvents';
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  var header = ['calendarId', 'eventId', 'startDatetime', 'endDatetime', 'student', 'teacher', 'attendance'];
  sheet.clear();
  sheet.appendRow(header);

  var now = new Date();
  var timeMinDate = new Date(now.getFullYear(), now.getMonth() - 3, now.getDate());
  var timeMaxDate = new Date(now.getFullYear(), now.getMonth() + 6, now.getDate());
  var timeMin = timeMinDate.toISOString();
  var timeMax = timeMaxDate.toISOString();
  
  var allEvents = [];
  var pageToken = null;
  var args = {
    timeMin: timeMin,
    timeMax: timeMax,
    singleEvents: true,
    orderBy: 'startTime',
    maxResults: 2500
  };
  do {
    if (pageToken) args.pageToken = pageToken;
    var resp = Calendar.Events.list(calendarId, args);
    if (resp && resp.items && resp.items.length) allEvents = allEvents.concat(resp.items);
    pageToken = resp && resp.nextPageToken ? resp.nextPageToken : null;
  } while (pageToken);

  var rows = allEvents.map(function(ev) {
    var student = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.student) || '';
    var teacher = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.teacher) || '';
    var attendance = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.attendance) || '';
    return [
      calendarId,
      ev.id || '',
      ev.start && (ev.start.dateTime || ev.start.date) || '',
      ev.end && (ev.end.dateTime || ev.end.date) || '',
      student,
      teacher,
      attendance
    ];
  });
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}