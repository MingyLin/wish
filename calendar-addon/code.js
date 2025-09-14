function onEventOpen(e) {
  var calendarId = e.calendar.calendarId;
  var eventId = e.calendar.id;
  if (!eventId) {
    return [createInfoCard('請選擇既有事件以使用此外掛。')];
  }
  var studentValue = '';
  var teacherValue = '';
  var attendanceValue = '';
  var event = null;
  try {
    event = Calendar.Events.get(calendarId, eventId);
    studentValue = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.student) || '';
    teacherValue = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.teacher) || '';
    attendanceValue = (event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.attendance) || '';
  } catch (err) {
    event = null;
  }
  if (!event) {
    return [createInfoCard('找不到事件資料，請先選擇既有事件。')];
  }
  var sheetId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
  var studentOptions = fetchSheetOptions(sheetId, 'Students!B:B');
  var teacherOptions = fetchSheetOptions(sheetId, 'Teachers!B:B');
  var studentDropdown = createDropdown('student', '學生', studentOptions, studentValue);
  var teacherDropdown = createDropdown('teacher', '老師', teacherOptions, teacherValue, 3);
  var attendanceRadio = createAttendanceRadio(attendanceValue, calendarId, eventId);
  var studentBatchBtn = createUpdateButton('student', 'batch', calendarId, eventId, '批次更新');
  var teacherSingleBtn = createUpdateButton('teacher', 'single', calendarId, eventId, '單次更新');
  var teacherBatchBtn = createUpdateButton('teacher', 'batch', calendarId, eventId, '批次更新');
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
      .addWidget(teacherBatchBtn)
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

function fetchSheetOptions(sheetId, range) {
  var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + sheetId + '/values/' + encodeURIComponent(range);
  var params = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, params);
  var result = JSON.parse(response.getContentText());
  if (result.values && result.values.length > 1) {
    var arr = [];
    for (var i = 1; i < result.values.length; i++) {
      if (result.values[i][0]) arr.push(result.values[i][0]);
    }
    return arr;
  }
  return ['選項讀取失敗'];
}

function createDropdown(fieldName, title, options, selected, maxSelect) {
  var dropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName(fieldName)
    .setTitle(title);
  if (maxSelect) dropdown.setMultiSelectMaxSelectedItems(maxSelect);
  for (var i = 0; i < options.length; i++) {
    dropdown.addItem(options[i], options[i], selected === options[i]);
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
    .addItem('自習', '自習', selected === '自習')
    .setOnChangeAction(CardService.newAction().setFunctionName('saveAttendanceField').setParameters({ calendarId: calendarId, eventId: eventId }));
  return radio;
}

function createUpdateButton(field, updateType, calendarId, eventId, text) {
  return CardService.newTextButton()
    .setText(text)
    .setOnClickAction(CardService.newAction()
      .setFunctionName('saveCustomField')
      .setParameters({ field: field, updateType: updateType, calendarId: calendarId, eventId: eventId }));
}

function saveCustomField(e) {
  var params = (e && e.parameters) || {};
  var calendarId = params.calendarId || 'primary';
  var eventId = params.eventId;
  var form = e && e.formInput ? e.formInput : {};
  var student = form.student || '';
  var teacher = form.teacher || '';
  var field = params.field;
  var updateType = params.updateType;
  if (!eventId) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('找不到事件 ID，無法儲存。'))).build();
  }
  try {
    var event = Calendar.Events.get(calendarId, eventId);
    var resource = { extendedProperties: { private: {} }, colorId: '1' };
    if (field === 'student') {
      resource.extendedProperties.private.student = student;
      resource.summary = student + '_(' + ((event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.teacher) || '') + ')';
    } else if (field === 'teacher') {
      resource.extendedProperties.private.teacher = teacher;
      resource.summary = ((event.extendedProperties && event.extendedProperties.private && event.extendedProperties.private.student) || '') + '_(' + teacher + ')';
    }
    if (field === 'teacher' && updateType === 'single') {
      Calendar.Events.patch(resource, calendarId, eventId);
      syncEventToSheet();
    } else {
      var masterId = event && event.recurringEventId ? event.recurringEventId : null;
      if (masterId) {
        try {
          var resourceRecurring = {
            summary: student + '_(' + teacher + ')',
            extendedProperties: { private: { student: student, teacher: teacher } },
            colorId: '1'
          };
          try { Calendar.Events.patch(resourceRecurring, calendarId, masterId); } catch (errMaster) {}
          var instancesResp = Calendar.Events.instances(calendarId, masterId, { maxResults: 2500 });
          if (instancesResp && instancesResp.items && instancesResp.items.length) {
            var eventsToWrite = [];
            for (var i = 0; i < instancesResp.items.length; i++) {
              var inst = instancesResp.items[i];
              try {
                Calendar.Events.patch(resourceRecurring, calendarId, inst.id);
                var attendanceInst = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.attendance) || '未點名';
                eventsToWrite.push({ calendarId: calendarId, eventId: inst.id, startDatetime: inst.start && (inst.start.dateTime || inst.start.date), endDatetime: inst.end && (inst.end.dateTime || inst.end.date), student: student, teacher: teacher, attendance: attendanceInst });
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
  else if (attendance === '自習') colorId = '9';
  if (!eventId) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('找不到事件 ID，無法儲存點名。'))).build();
  }
  try {
    var event = Calendar.Events.get(calendarId, eventId);
    var resource = { extendedProperties: { private: { attendance: attendance } }, colorId: colorId };
    Calendar.Events.patch(resource, calendarId, eventId);
    syncEventToSheet();
    return CardService.newActionResponseBuilder().setNotification(CardService.newNotification().setText('點名已儲存')).build();
  } catch (err) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('點名儲存失敗：' + err.message))).build();
  }
}

function syncEventToSheet() {
  syncCalendarToSheet.syncCalendarToSheet();
}