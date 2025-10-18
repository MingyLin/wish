function onEventOpen(e) {
  var calendarId = e.calendar.calendarId;
  var eventId = e.calendar.id;
  if (!eventId) {
    return [createInfoCard('請選擇既有事件以使用此外掛。')];
  }
  var studentValue = '';
  var teacherValue = '';
  var subjectValue = '';
  var attendanceValue = '';
  var event = null;
  try {
    event = Calendar.Events.get(calendarId, eventId);
    studentValue = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.student) || '';
    teacherValue = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.teacher) || '';
    subjectValue = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.subject) || '';
    attendanceValue = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.attendance) || '';
  } catch (err) {
    event = null;
  }
  if (!event) {
    return [createInfoCard('找不到事件資料，請先選擇既有事件。')];
  }
  var sheetId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
  var studentOptions = fetchSheetOptions(sheetId, 'Students!A:C', true);
  var teacherOptions = fetchSheetOptions(sheetId, 'Teachers!A:B');
  var subjectOptions = [
    { id: '國文', name: '國文' },
    { id: '英文', name: '英文' },
    { id: '數學', name: '數學' },
    { id: '生物', name: '生物' }
  ];
  var studentDropdown = createDropdown('student', '學生', studentOptions, studentValue);
  var teacherDropdown = createDropdown('teacher', '老師', teacherOptions, teacherValue, 3);
  var subjectDropdown = createDropdown('subject', '科目', subjectOptions, subjectValue);
  var attendanceRadio = createAttendanceRadio(attendanceValue, calendarId, eventId);
  var studentBatchBtn = createUpdateButton('student', 'batch', calendarId, eventId, '更新所有活動');
  var teacherSingleBtn = createUpdateButton('teacher', 'single', calendarId, eventId, '更新這項活動');
  var teacherFutureBtn = createUpdateButton('teacher', 'future', calendarId, eventId, '更新這項活動和後續活動');
  var teacherBatchBtn = createUpdateButton('teacher', 'batch', calendarId, eventId, '更新所有活動');
  var subjectSingleBtn = createUpdateButton('subject', 'single', calendarId, eventId, '更新這項活動');
  var subjectFutureBtn = createUpdateButton('subject', 'future', calendarId, eventId, '更新這項活動和後續活動');
  var subjectBatchBtn = createUpdateButton('subject', 'batch', calendarId, eventId, '更新所有活動');
  var infoWidget = CardService.newKeyValue().setTopLabel('目前標題').setContent(event.summary);
  var creatorWidget = CardService.newKeyValue().setTopLabel('建立者').setContent(event.creator.email);
  var updatedWidget = CardService.newKeyValue().setTopLabel('異動時間').setContent(event.updated);
  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('點名工具'))
    .addSection(CardService.newCardSection()
      .addWidget(infoWidget)
      .addWidget(creatorWidget)
      .addWidget(updatedWidget)
      .addWidget(studentDropdown)
      .addWidget(studentBatchBtn)
      .addWidget(teacherDropdown)
      .addWidget(teacherSingleBtn)
      .addWidget(teacherFutureBtn)
      .addWidget(teacherBatchBtn)
      .addWidget(subjectDropdown)
      .addWidget(subjectSingleBtn)
      .addWidget(subjectFutureBtn)
      .addWidget(subjectBatchBtn)
      .addWidget(attendanceRadio)
    )
    .build();
  return [card];
}

function createInfoCard(message) {
  return CardService.newCardBuilder()
    .addSection(CardService.newCardSection().addWidget(
      CardService.newTextParagraph().setText(message)
    ))
    .build();
}

function fetchSheetOptions(sheetId, range, includeThirdColumn) {
  var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + sheetId + '/values/' + encodeURIComponent(range);
  var params = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, params);
  var result = JSON.parse(response.getContentText());
  var arr = [];
  if (result.values && result.values.length > 1) {
    for (var i = 1; i < result.values.length; i++) {
      var row = result.values[i];
      if (row[0] && row[1]) {
        var option = { id: row[0], name: row[1] };
        if (includeThirdColumn && row[2]) {
          option.entranceYear = row[2];
          option.name = row[0] + '.' + row[1] + '(' + (new Date().getFullYear() - 1903 - (new Date().getMonth() > 8 ? 1 : 0) - row[2]) + ')';
        } else if (includeThirdColumn) {
          option.name = row[0] + '.' + row[1];
        } else {
          option.name = row[0] + '.' + row[1];
        }
        arr.push(option);
      }
    }
    return arr;
  }
  return [{ id: '', name: '選項讀取失敗' }];
}

function createDropdown(fieldName, title, options, selected, maxSelect) {
  var dropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName(fieldName)
    .setTitle(title);
  if (maxSelect) dropdown.setMultiSelectMaxSelectedItems(maxSelect);
  for (var i = 0; i < options.length; i++) {
    dropdown.addItem(options[i].name, options[i].id, selected === options[i].id);
  }
  return dropdown;
}

function createAttendanceRadio(selected, calendarId, eventId) {
  var radio = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.RADIO_BUTTON)
    .setFieldName('attendance')
    .setTitle('點名狀態')
    .addItem('未點名', '未點名', selected === '未點名' || !selected)
    .addItem('出席', '出席', selected === '出席')
    .addItem('缺席', '缺席', selected === '缺席')
    .addItem('請假', '請假', selected === '請假')
    .addItem('補課', '補課', selected === '補課')
    .addItem('自習', '自習', selected === '自習')
    .addItem('試聽', '試聽', selected === '試聽')
    .setOnChangeAction(CardService.newAction().setFunctionName('saveAttendanceField').setParameters({ calendarId: calendarId, eventId: eventId }));
  return radio;
}

function createUpdateButton(field, updateType, calendarId, eventId, text) {
  return CardService.newTextButton()
    .setText(text)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('saveField')
      .setParameters({ field: field, updateType: updateType, calendarId: calendarId, eventId: eventId }));
}

function saveField(e) {
  var params = (e && e.parameters) || {};
  var calendarId = params.calendarId || 'primary';
  var eventId = params.eventId;
  var form = e && e.formInput ? e.formInput : {};
  var studentId = form.student || '';
  var teacherId = form.teacher || '';
  var subjectId = form.subject || '';
  var field = params.field;
  var updateType = params.updateType;
  var sheetId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
  var studentOptions = fetchSheetOptions(sheetId, 'Students!A:C', true);
  var teacherOptions = fetchSheetOptions(sheetId, 'Teachers!A:B');
  var studentName = '';
  var teacherName = '';
  var subjectName = subjectId;
  for (var i = 0; i < studentOptions.length; i++) {
    if (studentOptions[i].id === studentId) {
      studentName = studentOptions[i].name;
      break;
    }
  }
  for (var j = 0; j < teacherOptions.length; j++) {
    if (teacherOptions[j].id === teacherId) {
      teacherName = teacherOptions[j].name;
      break;
    }
  }
  if (!eventId) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('找不到事件 ID，無法儲存。'))).build();
  }
  try {
    var event = Calendar.Events.get(calendarId, eventId);
    var resource = { extendedProperties: { private: {} }, colorId: '1' };
    
    // 取得目前事件的學生、老師、科目資訊
    var currentStudentId = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.student) || '';
    var currentTeacherId = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.teacher) || '';
    var currentSubject = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.subject) || '';
    
    var currentStudentName = '';
    var currentTeacherName = '';
    
    // 取得學生姓名
    for (var k = 0; k < studentOptions.length; k++) {
      if (studentOptions[k].id === currentStudentId) {
        currentStudentName = studentOptions[k].name;
        break;
      }
    }
    // 取得老師姓名
    for (var k = 0; k < teacherOptions.length; k++) {
      if (teacherOptions[k].id === currentTeacherId) {
        currentTeacherName = teacherOptions[k].name;
        break;
      }
    }
    
    if (field === 'student') {
      resource.extendedProperties.private.student = studentId;
      resource.summary = studentName + '-' + currentSubject + '(' + currentTeacherName + ')';
    } else if (field === 'teacher') {
      resource.extendedProperties.private.teacher = teacherId;
      resource.summary = currentStudentName + '-' + currentSubject + '(' + teacherName + ')';
    } else if (field === 'subject') {
      resource.extendedProperties.private.subject = subjectId;
      resource.summary = currentStudentName + '-' + subjectName + '(' + currentTeacherName + ')';
    }
    if ((field === 'teacher' || field === 'student' || field === 'subject') && updateType === 'single') {
      Calendar.Events.patch(resource, calendarId, eventId);
      syncEventToSheet();
    } else if (updateType === 'future') {
      var masterId = event && event.recurringEventId ? event.recurringEventId : null;
      if (masterId) {
        try {
          var resourceRecurring = {
            summary: currentStudentName + '-' + currentSubject + '(' + currentTeacherName + ')',
            extendedProperties: { private: { student: currentStudentId, teacher: currentTeacherId, subject: currentSubject } },
            colorId: '1'
          };
          
          if (field === 'student') {
            resourceRecurring.summary = studentName + '-' + currentSubject + '(' + currentTeacherName + ')';
            resourceRecurring.extendedProperties.private.student = studentId;
          } else if (field === 'teacher') {
            resourceRecurring.summary = currentStudentName + '-' + currentSubject + '(' + teacherName + ')';
            resourceRecurring.extendedProperties.private.teacher = teacherId;
          } else if (field === 'subject') {
            resourceRecurring.summary = currentStudentName + '-' + subjectName + '(' + currentTeacherName + ')';
            resourceRecurring.extendedProperties.private.subject = subjectId;
          }
          var instancesResp = Calendar.Events.instances(calendarId, masterId, { maxResults: 2500 });
          if (instancesResp && instancesResp.items && instancesResp.items.length) {
            var eventsToWrite = [];
            var currentInst = null;
            for (var i = 0; i < instancesResp.items.length; i++) {
              var inst = instancesResp.items[i];
              if (inst.id === eventId) {
                currentInst = inst;
                break;
              }
            }
            var currentStart = currentInst && (currentInst.start && (currentInst.start.dateTime || currentInst.start.date));
            for (var i = 0; i < instancesResp.items.length; i++) {
              var inst = instancesResp.items[i];
              var instStart = inst.start && (inst.start.dateTime || inst.start.date);
              if (currentStart && instStart >= currentStart) {
                try {
                  Calendar.Events.patch(resourceRecurring, calendarId, inst.id);
                  var attendanceInst = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.attendance) || '未點名';
                  var finalStudentId = field === 'student' ? studentId : currentStudentId;
                  var finalTeacherId = field === 'teacher' ? teacherId : currentTeacherId;
                  var finalSubject = field === 'subject' ? subjectId : currentSubject;
                  eventsToWrite.push({ calendarId: calendarId, eventId: inst.id, startDatetime: instStart, endDatetime: inst.end && (inst.end.dateTime || inst.end.date), student: finalStudentId, teacher: finalTeacherId, subject: finalSubject, attendance: attendanceInst });
                } catch (errInst) {
                  Logger.log('Failed to patch instance %s: %s', inst.id, errInst.message);
                }
              }
            }
            if (eventsToWrite.length) {
              try { syncEventToSheet(); } catch (errBulk) { Logger.log('Bulk sync to sheet failed: %s', errBulk.message); }
            }
          }
        } catch (err2) {
          Logger.log('Error updating future instances: %s', err2.message);
        }
      } else {
        Calendar.Events.patch(resource, calendarId, eventId);
        syncEventToSheet();
      }
    } else {
      var masterId = event && event.recurringEventId ? event.recurringEventId : null;
      if (masterId) {
        try {
          var resourceRecurring = {
            summary: currentStudentName + '-' + currentSubject + '-(' + currentTeacherName + ')',
            extendedProperties: { private: { student: currentStudentId, teacher: currentTeacherId, subject: currentSubject } },
            colorId: '1'
          };
          
          if (field === 'student') {
            resourceRecurring.summary = studentName + '-' + currentSubject + '(' + currentTeacherName + ')';
            resourceRecurring.extendedProperties.private.student = studentId;
          } else if (field === 'teacher') {
            resourceRecurring.summary = currentStudentName + '-' + currentSubject + '(' + teacherName + ')';
            resourceRecurring.extendedProperties.private.teacher = teacherId;
          } else if (field === 'subject') {
            resourceRecurring.summary = currentStudentName + '-' + subjectName + '(' + currentTeacherName + ')';
            resourceRecurring.extendedProperties.private.subject = subjectId;
          }
          
          try { Calendar.Events.patch(resourceRecurring, calendarId, masterId); } catch (errMaster) {}
          var instancesResp = Calendar.Events.instances(calendarId, masterId, { maxResults: 2500 });
          if (instancesResp && instancesResp.items && instancesResp.items.length) {
            var eventsToWrite = [];
            for (var i = 0; i < instancesResp.items.length; i++) {
              var inst = instancesResp.items[i];
              try {
                Calendar.Events.patch(resourceRecurring, calendarId, inst.id);
                var attendanceInst = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.attendance) || '未點名';
                var finalStudentId = field === 'student' ? studentId : currentStudentId;
                var finalTeacherId = field === 'teacher' ? teacherId : currentTeacherId;
                var finalSubject = field === 'subject' ? subjectId : currentSubject;
                eventsToWrite.push({ calendarId: calendarId, eventId: inst.id, startDatetime: inst.start && (inst.start.dateTime || inst.start.date), endDatetime: inst.end && (inst.end.dateTime || inst.end.date), student: finalStudentId, teacher: finalTeacherId, subject: finalSubject, attendance: attendanceInst });
              } catch (errInst) {
                Logger.log('Failed to patch instance %s: %s', inst.id, errInst.message);
              }
            }
            if (eventsToWrite.length) {
              try { syncEventToSheet(); } catch (errBulk) { Logger.log('Bulk sync to sheet failed: %s', errBulk.message); }
            }
          }
        } catch (err2) {
          Logger.log('Error updating recurring instances: %s', err2.message);
        }
      } else {
        Calendar.Events.patch(resource, calendarId, eventId);
        syncEventToSheet();
      }
    }
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('已儲存欄位。'))).build();
  } catch (err) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('儲存失敗：' + err.message))).build();
  }
}

function saveAttendanceField(e) {
  var params = (e && e.parameters) || {};
  var calendarId = params.calendarId || 'primary';
  var eventId = params.eventId;
  var form = e && e.formInput ? e.formInput : {};
  var attendance = form.attendance || '未點名';
  var colorId = '8';
  if (attendance === '出席') colorId = '10';
  else if (attendance === '請假') colorId = '5';
  else if (attendance === '缺席') colorId = '11';
  else if (attendance === '補課') colorId = '6';
  else if (attendance === '自習') colorId = '3';
  else if (attendance === '試聽') colorId = '2';
  if (!eventId) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('找不到事件 ID，無法儲存點名。'))).build();
  }
  try {
    var event = Calendar.Events.get(calendarId, eventId);
    var resource = { extendedProperties: { private: { attendance: attendance } }, colorId: colorId };
    Calendar.Events.patch(resource, calendarId, eventId);
    syncEventToSheet();

    var prevAttendance = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.attendance) || '';
    var isPrevAttend = (prevAttendance === '出席' || prevAttendance === '缺席');
    var isNewAttend = (attendance === '出席' || attendance === '缺席');

    var amount = 0;

    if (isPrevAttend && !isNewAttend) {
      amount = 1;
    }else if (!isPrevAttend && isNewAttend) {
      amount = -1;
    }

    if (amount !== 0) {
      var sheetId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
      var ss = SpreadsheetApp.openById(sheetId);
      var sheet = ss.getSheetByName('StockHistory');
      if (!sheet) sheet = ss.insertSheet('StockHistory');

      var studentId = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.student) || '';

      var classDateRaw = event.start && (event.start.dateTime || event.start.date) || '';
      var classDateObj = classDateRaw ? new Date(classDateRaw) : null;
      var classDate = classDateObj ? (classDateObj.getFullYear() + '-' + ('0' + (classDateObj.getMonth() + 1)).slice(-2) + '-' + ('0' + classDateObj.getDate()).slice(-2)) : '';

      var randId = genId();

      var now = new Date();
      var yyyy = now.getFullYear();
      var mm = ('0' + (now.getMonth() + 1)).slice(-2);
      var dd = ('0' + now.getDate()).slice(-2);
      var hh = ('0' + now.getHours()).slice(-2);
      var min = ('0' + now.getMinutes()).slice(-2);
      var nowStr = yyyy + '-' + mm + '-' + dd + ' ' + hh + ':' + min;

      sheet.appendRow([randId, studentId, classDate, amount, 'addon', nowStr, 'addon', nowStr]);
    }
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText('點名已儲存')).build();
  } catch (err) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('點名儲存失敗：' + err.message))).build();
  }
}

function syncEventToSheet() {
  syncCalendarToSheet.syncCalendarToSheet();
}

function genId() {
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var id = '';
  for (var i = 0; i < 10; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}

function syncCalendarToSheet() {
  var calendarId = '0a437281002a546d7e17233cefa484100bd212d88da1c792e9f162bdf1be23e3@group.calendar.google.com';
  var sheetId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
  var sheetName = 'CalendarEvents';
  var ss = SpreadsheetApp.openById(sheetId);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  var header = ['CalendarId', 'EventId', 'StartDatetime', 'EndDatetime', 'Student', 'Teacher', 'Attendance'];
  sheet.clear();
  sheet.appendRow(header);

  var now = new Date();
  var timeMinDate = new Date(2025, 8, 1);
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

  function formatTaipei(dtStr) {
    if (!dtStr) return '';
    var dt = new Date(dtStr);
    // 轉換為台灣時區 (UTC+8)
    var taipei = new Date(dt.getTime() + 8 * 60 * 60 * 1000);
    var yyyy = taipei.getUTCFullYear();
    var mm = ('0' + (taipei.getUTCMonth() + 1)).slice(-2);
    var dd = ('0' + taipei.getUTCDate()).slice(-2);
    var hh = ('0' + taipei.getUTCHours()).slice(-2);
    var min = ('0' + taipei.getUTCMinutes()).slice(-2);
    var ss = ('0' + taipei.getUTCSeconds()).slice(-2);
    return yyyy + '-' + mm + '-' + dd + ' ' + hh + ':' + min + ':' + ss;
  }
  var rows = allEvents.map(function(ev) {
    var student = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.student) || '';
    var teacher = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.teacher) || '';
    var attendance = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.attendance) || '';
    var startRaw = ev.start && (ev.start.dateTime || ev.start.date) || '';
    var endRaw = ev.end && (ev.end.dateTime || ev.end.date) || '';
    return [
      calendarId,
      ev.id || '',
      formatTaipei(startRaw),
      formatTaipei(endRaw),
      student,
      teacher,
      attendance
    ];
  });
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }
}

function syncPurchaseHistoryToReport() {
	var srcId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
	var srcSs = SpreadsheetApp.openById(srcId);
	var srcSheet = srcSs.getSheetByName('StockHistory');
	if (!srcSheet) {
		throw new Error('找不到來源分頁 StockHistory');
	}

	// 讀取 Students 對照表
	var studentsSheet = srcSs.getSheetByName('Students');
	var idToName = {};
	if (studentsSheet) {
		var studentsData = studentsSheet.getDataRange().getValues();
		for (var i = 1; i < studentsData.length; i++) {
			var row = studentsData[i];
			var id = row[0] !== undefined ? String(row[0]) : '';
			var name = row[1] !== undefined ? String(row[1]) : '';
			if (id) idToName[id] = name;
		}
	}

	// 讀取來源資料（假設有 header）
	var data = srcSheet.getDataRange().getValues();
	if (data.length <= 1) {
		// 沒資料
		writeToDest([]);
		return;
	}

	var header = data[0];
	var rows = data.slice(1);

	// 找出來源欄位的 index（寬容處理）
	var idx = {};
	for (var h = 0; h < header.length; h++) {
		var key = (header[h] || '').toString().trim();
		idx[key] = h;
	}

	// 支援欄位名稱變體
	function colIndex(names) {
		for (var i = 0; i < names.length; i++) {
			if (idx.hasOwnProperty(names[i])) return idx[names[i]];
		}
		return -1;
	}

	var si = colIndex(['Student', 'student', 'StudentId', 'studentId']);
	var di = colIndex(['Desc', 'desc', 'Description', 'description']);
	var ai = colIndex(['Amount', 'amount']);
	var ui = colIndex(['UpdatedAt', 'Updated At', 'updatedAt', 'updated_at']);

	var out = [];
	for (var r = 0; r < rows.length; r++) {
		var row = rows[r];
		var studentId = si >= 0 ? String(row[si]) : '';
		var desc = di >= 0 ? row[di] : '';
		var amount = ai >= 0 ? Number(row[ai]) : 0;
		var updatedAtRaw = ui >= 0 ? row[ui] : '';

		var studentName = idToName[studentId] || studentId || '';
		var dateStr = '';
		var dateYYYYMM = '';
		if (updatedAtRaw) {
			try {
				var dt = new Date(updatedAtRaw);
				if (!isNaN(dt.getTime())) {
					dateStr = dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2) + '-' + ('0' + dt.getDate()).slice(-2);
					dateYYYYMM = dt.getFullYear() + '/' + ('0' + (dt.getMonth() + 1)).slice(-2);
				} else {
					// 如果不是可解析的日期，嘗試只取字串中的日期部分
					var s = String(updatedAtRaw);
					var m = s.match(/(\d{4})[-\/]?(\d{2})[-\/]?(\d{2})/);
					if (m) {
						dateStr = m[1] + '-' + m[2] + '-' + m[3];
						dateYYYYMM = m[1] + '/' + m[2];
					} else {
						dateStr = s;
						dateYYYYMM = '';
					}
				}
			} catch (e) {
				dateStr = String(updatedAtRaw);
				dateYYYYMM = '';
			}
		}
		var category = (amount > 0) ? '購買' : ('上課' + (dateYYYYMM ? (' ' + dateYYYYMM) : ''));
		var title = desc;
		var qty = amount;
		out.push([studentName, category, title, qty, dateStr]);
	}

	writeToDest(out);
}

function writeToDest(rows) {
    var destId = '1EsDKBw0malhT8o0s_OF28QMnNfScTerF-FK-9W36GOE';
    var destSs = SpreadsheetApp.openById(destId);
    var destSheet = destSs.getSheetByName('stock');
    if (!destSheet) destSheet = destSs.insertSheet('stock');
    // 清空並寫入 header + rows
    destSheet.clear();
	var headerOut = ['學生', 'category', 'title', '數量', '日期', 'InRange'];
	destSheet.appendRow(headerOut);
	if (rows && rows.length) {
		// 寫入資料
		destSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
		// 寫入 InRange 公式
		var inRangeCol = headerOut.length; // F 欄
		var dateCol = headerOut.indexOf('日期') + 1; // E 欄
		for (var i = 0; i < rows.length; i++) {
			var rowIdx = i + 2;
			// 公式：=AND(E2 >= 'stock-report'!B$1, E2 <= 'stock-report'!D$1)
			var formula = '=AND(E' + rowIdx + " >= 'stock-report'!B$1, E" + rowIdx + " <= 'stock-report'!D$1)";
			destSheet.getRange(rowIdx, inRangeCol).setFormula(formula);
		}
	}
}

function syncTeacherAttendanceToReport() {
  var srcId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
  var srcSs = SpreadsheetApp.openById(srcId);
  var eventsSheet = srcSs.getSheetByName('CalendarEvents');
  var teachersSheet = srcSs.getSheetByName('Teachers');
  if (!eventsSheet || !teachersSheet) throw new Error('來源分頁不存在');

  // 取得 Teacher ID->Name 對照表
  var teacherMap = {};
  var tData = teachersSheet.getDataRange().getValues();
  for (var i = 1; i < tData.length; i++) {
    var row = tData[i];
    var id = row[0] !== undefined ? String(row[0]) : '';
    var name = row[1] !== undefined ? String(row[1]) : '';
    if (id) teacherMap[id] = name;
  }

  // 取得 CalendarEvents 資料
  var data = eventsSheet.getDataRange().getValues();
  if (data.length <= 1) return;
  var header = data[0];
  var rows = data.slice(1);
  var idx = {};
  for (var h = 0; h < header.length; h++) idx[header[h]] = h;
  var ti = idx['Teacher'] || idx['teacher'];
  var ai = idx['Attendance'] || idx['attendance'];
  var si = idx['StartDatetime'] || idx['startDatetime'] || idx['startdatetime'];
  var ei = idx['EndDatetime'] || idx['endDatetime'] || idx['enddatetime'];

  // 過濾出席資料，並依老師+日期分組，只合併重疊或連續時段
  var groupMap = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (!row[ai] || String(row[ai]).trim() !== '出席') continue;
    var teacherId = row[ti] ? String(row[ti]) : '';
    var teacherName = teacherMap[teacherId] || teacherId;
    var startStr = row[si] ? String(row[si]) : '';
    var endStr = row[ei] ? String(row[ei]) : '';
    if (!teacherName || !startStr || !endStr) continue;
    
    // 取日期（yyyy-MM-dd）
    var dateKey = '';
    try {
      var dt = new Date(startStr);
      if (!isNaN(dt.getTime())) {
        dateKey = dt.getFullYear() + '-' + ('0' + (dt.getMonth() + 1)).slice(-2) + '-' + ('0' + dt.getDate()).slice(-2);
      } else {
        dateKey = startStr.substr(0, 10);
      }
    } catch (e) {
      dateKey = startStr.substr(0, 10);
    }
    
    var baseKey = teacherName + '|' + dateKey;
    var merged = false;
    
    // 檢查是否可與現有時段合併
    for (var existingKey in groupMap) {
      if (!existingKey.startsWith(baseKey + '|')) continue;
      var existing = groupMap[existingKey];
      
      // 檢查時段是否重疊或連續
      if (isOverlapOrContinuous(startStr, endStr, existing.start, existing.end)) {
        // 合併時段：取最早開始、最晚結束
        if (compareTime(startStr, existing.start) < 0) existing.start = startStr;
        if (compareTime(endStr, existing.end) > 0) existing.end = endStr;
        merged = true;
        break;
      }
    }
    
    // 若無法合併，新增時段
    if (!merged) {
      var counter = 1;
      var newKey = baseKey + '|' + counter;
      while (groupMap[newKey]) {
        counter++;
        newKey = baseKey + '|' + counter;
      }
      groupMap[newKey] = { teacher: teacherName, date: dateKey, start: startStr, end: endStr };
    }
  }
  
  // 檢查兩個時段是否重疊或連續
  function isOverlapOrContinuous(start1, end1, start2, end2) {
    var s1 = new Date(start1);
    var e1 = new Date(end1);
    var s2 = new Date(start2);
    var e2 = new Date(end2);
    
    if (isNaN(s1.getTime()) || isNaN(e1.getTime()) || isNaN(s2.getTime()) || isNaN(e2.getTime())) {
      return false;
    }
    
    // 重疊：start1 < end2 && start2 < end1
    // 連續：end1 == start2 || end2 == start1
    return (s1 < e2 && s2 < e1) || (e1.getTime() === s2.getTime()) || (e2.getTime() === s1.getTime());
  }

  // 整理輸出
  var out = [];
  for (var k in groupMap) {
    var g = groupMap[k];
    var hour = calcHour(g.start, g.end);
    var startFmt = formatDateTimeStr(g.start);
    var endFmt = formatDateTimeStr(g.end);
    // 強制轉為字串避免 Date 物件被自動格式化
    out.push([String(g.teacher), String(startFmt), String(endFmt), String(hour)]);
  }

  // 時間格式 yyyy/MM/dd HH:mm (減少15小時修正時區)
  function formatDateTimeStr(dtVal) {
    var dt = null;
    
    // 若是 Date 物件
    if (Object.prototype.toString.call(dtVal) === '[object Date]' && !isNaN(dtVal.getTime())) {
      dt = new Date(dtVal.getTime() - 15 * 60 * 60 * 1000); // 減少15小時
    }
    // 若是字串，先解析再減少15小時
    else if (typeof dtVal === 'string') {
      var tempDt = new Date(dtVal);
      if (!isNaN(tempDt.getTime())) {
        dt = new Date(tempDt.getTime() - 15 * 60 * 60 * 1000);
      }
    }
    // Google Sheets 內部日期數字
    else if (typeof dtVal === 'number') {
      var tempDt = new Date(Math.round((dtVal - 25569) * 86400 * 1000));
      if (!isNaN(tempDt.getTime())) {
        dt = new Date(tempDt.getTime() - 15 * 60 * 60 * 1000);
      }
    }
    
    if (dt && !isNaN(dt.getTime())) {
      var yyyy = dt.getFullYear();
      var mm = ('0' + (dt.getMonth() + 1)).slice(-2);
      var dd = ('0' + dt.getDate()).slice(-2);
      var hh = ('0' + dt.getHours()).slice(-2);
      var min = ('0' + dt.getMinutes()).slice(-2);
      return yyyy + '/' + mm + '/' + dd + ' ' + hh + ':' + min;
    }
    
    return String(dtVal);
  }

  // 寫入目標 teacher 分頁
  var destId = '1EsDKBw0malhT8o0s_OF28QMnNfScTerF-FK-9W36GOE';
  var destSs = SpreadsheetApp.openById(destId);
  var destSheet = destSs.getSheetByName('teacher');
  if (!destSheet) destSheet = destSs.insertSheet('teacher');
  destSheet.clear();
  destSheet.appendRow(['老師', '開始時間', '結束時間', '時數']);
  if (out.length) destSheet.getRange(2, 1, out.length, out[0].length).setValues(out);
}

function compareTime(a, b) {
    var ta = new Date(a);
    var tb = new Date(b);
    if (isNaN(ta.getTime()) || isNaN(tb.getTime())) return String(a).localeCompare(String(b));
    return ta - tb;
}

function calcHour(start, end) {
    var t1 = new Date(start);
    var t2 = new Date(end);
    if (isNaN(t1.getTime()) || isNaN(t2.getTime())) return '';
    var diff = (t2 - t1) / (1000 * 60 * 60);
    return Math.round(diff * 10) / 10;
}

function aggregateStockHistory() {
  var srcId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
  var srcSs = SpreadsheetApp.openById(srcId);
  
  // 取得來源分頁
  var studentPurchasesSheet = srcSs.getSheetByName('StudentPurchases');
  var calendarEventsSheet = srcSs.getSheetByName('CalendarEvents');
  var stockHistorySheet = srcSs.getSheetByName('StockHistory');
  
  if (!studentPurchasesSheet || !calendarEventsSheet) {
    throw new Error('來源分頁不存在');
  }
  
  // 如果 StockHistory 分頁不存在，建立它
  if (!stockHistorySheet) {
    stockHistorySheet = srcSs.insertSheet('StockHistory');
    // 設定標題列
    stockHistorySheet.appendRow(['ID', 'Student', 'Desc', 'Amount', 'CreatedUser', 'UpdatedUser', 'CreatedAt', 'UpdatedAt']);
  } else {
    // 清除現有資料（保留標題列）
    if (stockHistorySheet.getLastRow() > 1) {
      stockHistorySheet.getRange(2, 1, stockHistorySheet.getLastRow() - 1, stockHistorySheet.getLastColumn()).clear();
    }
  }
  
  var currentTime = new Date();
  var aggregatedData = [];
  
  // 處理 StudentPurchases 資料
  var purchasesData = studentPurchasesSheet.getDataRange().getValues();
  if (purchasesData.length > 1) {
    var purchasesHeader = purchasesData[0];
    var purchasesRows = purchasesData.slice(1);
    
    // 找出欄位索引
    var purchasesIdx = {};
    for (var h = 0; h < purchasesHeader.length; h++) {
      purchasesIdx[purchasesHeader[h]] = h;
    }
    
    var idIdx = purchasesIdx['ID'] !== undefined ? purchasesIdx['ID'] : purchasesIdx['id'];
    var studentIdx = purchasesIdx['Student'] !== undefined ? purchasesIdx['Student'] : purchasesIdx['student'];
    var descIdx = purchasesIdx['Desc'] !== undefined ? purchasesIdx['Desc'] : purchasesIdx['desc'];
    var purchasedQtyIdx = purchasesIdx['PurchasedQty'] !== undefined ? purchasesIdx['PurchasedQty'] : purchasesIdx['purchasedQty'];
    
    for (var i = 0; i < purchasesRows.length; i++) {
      var row = purchasesRows[i];
      if (row[idIdx] !== undefined && row[idIdx] !== '') {
        aggregatedData.push([
          row[idIdx],                    // ID
          row[studentIdx] || '',         // Student
          row[descIdx] || '',            // Desc
          row[purchasedQtyIdx] || 0,     // Amount (PurchasedQty)
          'aggregateStockHistory',       // CreatedUser
          'aggregateStockHistory',       // UpdatedUser
          currentTime,                   // CreatedAt
          currentTime                    // UpdatedAt
        ]);
      }
    }
  }
  
  // 處理 CalendarEvents 資料
  var eventsData = calendarEventsSheet.getDataRange().getValues();
  if (eventsData.length > 1) {
    var eventsHeader = eventsData[0];
    var eventsRows = eventsData.slice(1);
    
    // 找出欄位索引
    var eventsIdx = {};
    for (var h = 0; h < eventsHeader.length; h++) {
      eventsIdx[eventsHeader[h]] = h;
    }
    
    var eventIdIdx = eventsIdx['EventId'] !== undefined ? eventsIdx['EventId'] : eventsIdx['eventId'];
    var studentIdx = eventsIdx['Student'] !== undefined ? eventsIdx['Student'] : eventsIdx['student'];
    var attendanceIdx = eventsIdx['Attendance'] !== undefined ? eventsIdx['Attendance'] : eventsIdx['attendance'];
    var startDatetimeIdx = eventsIdx['StartDatetime'] !== undefined ? eventsIdx['StartDatetime'] : eventsIdx['startDatetime'];
    
    for (var i = 0; i < eventsRows.length; i++) {
      var row = eventsRows[i];
      // 只處理出席的資料
      if (row[attendanceIdx] && String(row[attendanceIdx]).trim() === '出席' && row[eventIdIdx] !== undefined && row[eventIdIdx] !== '') {
        // 格式化日期為 yyyyMMdd
        var desc = '';
        if (row[startDatetimeIdx]) {
          try {
            var startDate = new Date(row[startDatetimeIdx]);
            if (!isNaN(startDate.getTime())) {
              var yyyy = startDate.getFullYear();
              var mm = ('0' + (startDate.getMonth() + 1)).slice(-2);
              var dd = ('0' + startDate.getDate()).slice(-2);
              desc = yyyy + mm + dd;
            } else {
              desc = String(row[startDatetimeIdx]).substr(0, 8).replace(/[^0-9]/g, '');
            }
          } catch (e) {
            desc = String(row[startDatetimeIdx]).substr(0, 8).replace(/[^0-9]/g, '');
          }
        }
        
        aggregatedData.push([
          row[eventIdIdx],               // ID (EventId)
          row[studentIdx] || '',         // Student
          desc,                          // Desc (yyyyMMdd format)
          -1,                            // Amount (-1)
          'aggregateStockHistory',       // CreatedUser
          'aggregateStockHistory',       // UpdatedUser
          currentTime,                   // CreatedAt
          currentTime                    // UpdatedAt
        ]);
      }
    }
  }
  
  // 寫入 StockHistory
  if (aggregatedData.length > 0) {
    stockHistorySheet.getRange(2, 1, aggregatedData.length, 8).setValues(aggregatedData);
  }
  
  Logger.log('已處理 ' + aggregatedData.length + ' 筆資料到 StockHistory');
  return '成功處理 ' + aggregatedData.length + ' 筆資料';
}