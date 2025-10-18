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