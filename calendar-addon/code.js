function onEventOpen(e) {
  var calendarId = e.calendar.calendarId;
  var eventId = e.calendar.id;

  if (!eventId) {
    var card = CardService.newCardBuilder()
      .addSection(CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText('請點選 "既有事件" 以使用此外掛。')
      ))
      .build();
    return [card];
  }

  var existingValue = '';
  var existingValue2 = '';
  var existingAttendance = '';
  var ev = null;
  try {
    ev = Calendar.Events.get(calendarId, eventId);
    existingValue = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.student) || '';
    existingValue2 = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.teacher) || '';
    existingAttendance = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.attendance) || '';
    var attendanceAction = CardService.newAction()
      .setFunctionName('saveAttendanceField')
      .setParameters({ calendarId: calendarId, eventId: eventId });
    var attendanceInput = CardService.newSelectionInput()
      .setType(CardService.SelectionInputType.RADIO_BUTTON)
      .setFieldName('attendance')
      .setTitle('點名狀態')
      .addItem('未點名', '未點名', existingAttendance === '未點名' || !existingAttendance)
      .addItem('出席', '出席', existingAttendance === '出席')
      .addItem('缺席', '缺席', existingAttendance === '缺席')
      .addItem('請假', '請假', existingAttendance === '請假')
      .addItem('自習', '自習', existingAttendance === '自習')
      .setOnChangeAction(attendanceAction);
  } catch (err) {
    existingValue = '';
    existingValue2 = '';
    ev = null;
  }

  if (!ev) {
    var card = CardService.newCardBuilder()
      .addSection(CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText('找不到事件資料，請先點選 "既有事件"。')
      ))
      .build();
    return [card];
  }

  var sheetId = '15EbnrqcDcvhlKOJ3L0cZxzRLiiZqQp-BrYSdwq1tnZ8';
  var values = getSheetValues(sheetId, 'Students!B:B');
  var values2 = getSheetValues(sheetId, 'Teachers!B:B');

  var input = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName('student')
    .setTitle('學生');
  values.forEach(function(option) {
    input.addItem(option, option, existingValue === option);
  });

  var input2 = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setMultiSelectMaxSelectedItems(3)
    .setFieldName('teacher')
    .setTitle('老師');
  values2.forEach(function(option) {
    input2.addItem(option, option, existingValue2 === option);
  });

  var action = CardService.newAction()
    .setFunctionName('saveCustomField')
    .setParameters({ calendarId: calendarId, eventId: eventId });

  const studentBatchUpdateButton = CardService.newTextButton()
    .setText('批次更新')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('saveCustomField')
      .setParameters({
        field: 'student',
        updateType: 'batch',
        calendarId: calendarId,
        eventId: eventId
      }));

  const teacherSingleUpdateButton = CardService.newTextButton()
    .setText('單次更新')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('saveCustomField')
      .setParameters({
        field: 'teacher',
        updateType: 'single',
        calendarId: calendarId,
        eventId: eventId
      }));

  const teacherBatchUpdateButton = CardService.newTextButton()
    .setText('批次更新')
    .setOnClickAction(CardService.newAction()
      .setFunctionName('saveCustomField')
      .setParameters({
        field: 'teacher',
        updateType: 'batch',
        calendarId: calendarId,
        eventId: eventId
      }));

  var infoWidget = CardService.newKeyValue()
    .setTopLabel('目前標題')
    .setContent(ev.summary);

  var creatorWidget = CardService.newKeyValue()
    .setTopLabel('建立者')
    .setContent(ev.creator.email);

  var updatedWidget = CardService.newKeyValue()
    .setTopLabel('異動時間')
    .setContent(ev.updated);

  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('點名工具'))
    .addSection(
      CardService.newCardSection()
        .addWidget(infoWidget)
        .addWidget(creatorWidget)
        .addWidget(updatedWidget)
        .addWidget(input)
        .addWidget(studentBatchUpdateButton)
        .addWidget(input2)
        .addWidget(teacherSingleUpdateButton)
        .addWidget(teacherBatchUpdateButton)
        .addWidget(attendanceInput)
    )
    .build();

  return [card];
}

function getSheetValues(sheetId, range) {
  var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + sheetId + '/values/' + encodeURIComponent(range);
  var params = {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  var response = UrlFetchApp.fetch(url, params);
  var result = JSON.parse(response.getContentText());
  if (result.values && result.values.length > 1) {
    return result.values.slice(1).map(function(row) { return row[0]; }).filter(function(v){ return v; });
  } else {
    return ['選項讀取失敗'];
  }
}

function saveCustomField(e) {
  var params = (e && e.parameters) || {};
  var calendarId = params.calendarId || 'primary';
  var eventId = params.eventId;
  var form = e && e.formInput ? e.formInput : {};
  var value = form.student || '';
  var value2 = form.teacher || '';
  var field = params.field;
  var updateType = params.updateType;

  if (!eventId) {
    var card = CardService.newCardBuilder()
      .addSection(CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText('找不到事件 ID，無法儲存。')
      ))
      .build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(card)).build();
  }

  try {
    var ev = Calendar.Events.get(calendarId, eventId);
    var resource = {
      extendedProperties: { private: {} },
      colorId: "1"
    };
    if (field === 'student') {
      resource.extendedProperties.private.student = value;
      resource.summary = value + '_(' + (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.teacher || '') + ')';
    } else if (field === 'teacher') {
      resource.extendedProperties.private.teacher = value2;
      resource.summary = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.student || '') + '_(' + value2 + ')';
    }

    if (field === 'teacher' && updateType === 'single') {
      Calendar.Events.patch(resource, calendarId, eventId);
      syncEventToSheet();
    } else {
      var masterId = ev && ev.recurringEventId ? ev.recurringEventId : null;
      if (masterId) {
        try {
          var resourceRecurring = {
            summary: value + '_(' + value2 + ')',
            extendedProperties: {
              private: {
                student: value,
                teacher: value2
              }
            },
            colorId: "1"
          };
          try {
            Calendar.Events.patch(resourceRecurring, calendarId, masterId);
          } catch (errMaster) {
          }

          var instancesResp = Calendar.Events.instances(calendarId, masterId, { maxResults: 2500 });
          if (instancesResp && instancesResp.items && instancesResp.items.length) {
            var eventsToWrite = [];
            instancesResp.items.forEach(function(inst) {
              try {
                Calendar.Events.patch(resourceRecurring, calendarId, inst.id);
                var attendanceInst = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.attendance) || '未點名';
                eventsToWrite.push({
                  calendarId: calendarId,
                  eventId: inst.id,
                  startDatetime: inst.start && (inst.start.dateTime || inst.start.date),
                  endDatetime: inst.end && (inst.end.dateTime || inst.end.date),
                  student: value,
                  teacher: value2,
                  attendance: attendanceInst
                });
              } catch (errInst) {
                Logger.log('Failed to patch instance %s: %s', inst.id, errInst.message);
              }
            });

            if (eventsToWrite.length) {
              try {
                syncEventToSheet();
              } catch (errBulk) {
                Logger.log('Bulk sync to sheet failed: %s', errBulk.message);
              }
            }
          }
        } catch (err2) {
          Logger.log('Error updating recurring instances: %s', err2.message);
        }
      }else {
        syncEventToSheet();
      }
    }

    var card = CardService.newCardBuilder()
      .addSection(CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText('已儲存欄位。')
      ))
      .build();

    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(card)).build();
  } catch (err) {
    var card = CardService.newCardBuilder()
      .addSection(CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText('儲存失敗：' + err.message)
      ))
      .build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(card)).build();
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
    var card = CardService.newCardBuilder()
      .addSection(CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText('找不到事件 ID，無法儲存點名。')
      ))
      .build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(card)).build();
  }

  try {
    var ev = Calendar.Events.get(calendarId, eventId);
    var resource = {
      extendedProperties: {
        private: {
          attendance: attendance
        }
      },
      colorId: colorId
    };
    Calendar.Events.patch(resource, calendarId, eventId);

    syncEventToSheet();
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('點名已儲存'))
      .build();
  } catch (err) {
    var card = CardService.newCardBuilder()
      .addSection(CardService.newCardSection().addWidget(
        CardService.newTextParagraph().setText('點名儲存失敗：' + err.message)
      ))
      .build();
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(card)).build();
  }
}

function syncEventToSheet() {
  syncCalendarToSheet.syncCalendarToSheet();
}