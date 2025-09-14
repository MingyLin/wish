function onEventOpen(e) {
  var calendarId = e.calendar.calendarId;  
  var eventId = e.calendar.id;  

  // 若沒有 eventId（例如正在建立新事件），提示使用者選取既有事件
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
  var ev = null; // 事先宣告，避免後續存取時發生 ReferenceError
  try {
    ev = Calendar.Events.get(calendarId, eventId);

    existingValue = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.student) || '';
    existingValue2 = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.teacher) || '';
    existingAttendance = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.attendance) || '';
    // 建立點名 radio button
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
    // 取不到事件時，保持 existing 值為空並在下方顯示提示
    existingValue = '';
    existingValue2 = '';
    ev = null;
  }

  // 若無法取得事件物件，提示使用者選擇既有事件後再使用功能
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

  // 建立下拉選單
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

  // Update the "儲存欄位" button to "批次更新" for students
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

  // Add "單次更新" and "批次更新" buttons for teachers
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

  var countAction = CardService.newAction()
    .setFunctionName('showCountCardAction')
    .setParameters({ calendarId: calendarId, eventId: eventId, value: existingValue });
/*
  var countButton = CardService.newTextButton()
    .setText('查詢此學生課堂數')
    .setOnClickAction(countAction);
*/
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
    /*
    .addSection(
      CardService.newCardSection()
        .addWidget(countButton)
    )*/
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
    // 跳過第一列（標題）
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

    // 單次更新只更新目前 event
    if (field === 'teacher' && updateType === 'single') {
      Calendar.Events.patch(resource, calendarId, eventId);
      var attendance = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.attendance) || '未點名';
      syncEventToSheet();
    } else {
      // 若此事件屬於一個 recurring series，則同時更新該 series 的所有 instances，並寫入 Sheet
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
            // 可能無法更新 master（權限或不存在），忽略並繼續更新 instances
          }

          // 列出 instances 並逐一更新；將要寫入 Sheet 的 events 收集起來，最後一次批次寫入以加速
          var instancesResp = Calendar.Events.instances(calendarId, masterId, { maxResults: 2500 });
          if (instancesResp && instancesResp.items && instancesResp.items.length) {
            var eventsToWrite = [];
            instancesResp.items.forEach(function(inst) {
              try {
                Calendar.Events.patch(resourceRecurring, calendarId, inst.id);
                // 取得 instance 的 attendance
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
                // 個別 instance 更新失敗則記錄並繼續
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
        
        var attendance = (ev.extendedProperties && ev.extendedProperties.private && ev.extendedProperties.private.attendance) || '未點名';
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

  // 設定 colorId 對應
  var colorId = '8'; // 預設灰色
  if (attendance === '出席') colorId = '10'; // 綠色
  else if (attendance === '請假') colorId = '5'; // 黃色
  else if (attendance === '缺席') colorId = '11'; // 紅色
  else if (attendance === '自習') colorId = '9'; // 藍色

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

// Count events in a calendar whose extendedProperties.private.student equals fieldValue.
// Uses Advanced Calendar service (Calendar.Events.list). Returns integer count.
function countEventsByCustomField(calendarId, fieldValue, timeMin, timeMax) {
  calendarId = calendarId || 'primary';
  if (!fieldValue) return 0;

  var args = {
    singleEvents: true,
    maxResults: 2500,
    // filter by private extended property (Calendar API supports this)
    privateExtendedProperty: 'student=' + fieldValue
  };
  if (timeMin) args.timeMin = (new Date(timeMin)).toISOString();
  if (timeMax) args.timeMax = (new Date(timeMax)).toISOString();

  var count = 0;
  var pageToken = null;
  do {
    if (pageToken) args.pageToken = pageToken;
    var resp = Calendar.Events.list(calendarId, args);
    if (resp && resp.items && resp.items.length) {
      count += resp.items.length;
    }
    pageToken = resp && resp.nextPageToken ? resp.nextPageToken : null;
  } while (pageToken);

  return count;
}

// Simple card action to show the count for a given value (can be wired to a button or run manually).
function showCountCardAction(e) {
  // e may be undefined when run from editor. Support parameters or formInput.
  var params = (e && e.parameters) || {};
  var form = (e && e.formInput) || {};
  var calendarId = params.calendarId || 'primary';
  var value = form.student || params.value || '查詢值';
  var timeMin = params.timeMin || null;
  var timeMax = params.timeMax || null;

  var count = 0;
  try {
    count = countEventsByCustomField(calendarId, value, timeMin, timeMax);
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('學生：' + value + '，課堂數量：' + count))
      .build();
  } catch (err) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('查詢失敗：' + err.message))
      .build();
  }
}

// 批次寫入多筆 events 到 sheet：一次讀 header 和資料，使用 setValues 批次更新並 append 多列

// helper: 建立一列完整的 array，依 header 順序填入 ev 的欄位
function buildRowArrayFromHeader(header, ev) {
  var row = [];
  for (var i = 0; i < header.length; i++) {
    var col = header[i];
    switch(col) {
      case 'calendarId': row.push(ev.calendarId || ''); break;
      case 'eventId': row.push(ev.eventId || ''); break;
      case 'startDatetime': row.push(ev.startDatetime || ''); break;
      case 'endDatetime': row.push(ev.endDatetime || ''); break;
      case 'student': row.push(ev.student || ''); break;
      case 'teacher': row.push(ev.teacher || ''); break;
      case 'attendance': row.push(ev.attendance || ''); break;
      default: row.push(''); break;
    }
  }
  return row;
}