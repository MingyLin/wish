function onEventOpen(e) {
  var calendarId = e.calendar && e.calendar.calendarId;
  var eventId = e.calendar && e.calendar.id;
    
    var ALLOWED_CALENDAR_ID = getConfig().allowedCalendarId;
  if (calendarId !== ALLOWED_CALENDAR_ID) {
    return [createInfoCard('僅支援指定日曆，請在對應日曆中使用。')];
  }

  if (!eventId) {
    return [createInfoCard('請選擇既有事件。')];
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
    var sheetId = getConfig().mainSheetId;
    var studentOptions = fetchSheetOptions(sheetId, 'Students!A:C', true);
    var teacherOptions = fetchSheetOptions(sheetId, 'Teachers!A:B');
  var subjectOptions = [
    { id: '國文', name: '國文' },
    { id: '英文', name: '英文' },
    { id: '數學', name: '數學' },
    { id: '理化', name: '理化' },
    { id: '物理', name: '物理' },
    { id: '化學', name: '化學' },
    { id: '生物', name: '生物' },
    { id: '地科', name: '地科' },
    { id: '歷史', name: '歷史' },
    { id: '地理', name: '地理' },
    { id: '公民', name: '公民' }
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
  var updatedWidget = CardService.newKeyValue().setTopLabel('更新時間').setContent(event.updated);
  var card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('點名工具'))
    .addSection(CardService.newCardSection()
      .addWidget(infoWidget)
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
  // Faster implementation: use SpreadsheetApp to read ranges and CacheService to cache results.
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = 'fetchSheetOptions|' + sheetId + '|' + range + '|' + (includeThirdColumn ? '1' : '0');
    try {
      var cached = cache.get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (e) { /* ignore cache errors */ }

    // Parse sheet name and A1 range (e.g. 'Students!A:C')
    var sheetName = null;
    var rangePart = range;
    if (range.indexOf('!') !== -1) {
      var parts = range.split('!');
      sheetName = parts[0];
      rangePart = parts.slice(1).join('!');
    }

    var ss = SpreadsheetApp.openById(sheetId);
    var sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
    if (!sheet) return [{ id: '', name: '選項讀取失敗' }];

    // Determine start column and number of columns from rangePart (supports forms like 'A:C' or 'A:B')
    function colLetterToIndex(letter) {
      var s = letter.toUpperCase();
      var col = 0;
      for (var i = 0; i < s.length; i++) {
        col = col * 26 + (s.charCodeAt(i) - 64);
      }
      return col;
    }

    var rpParts = rangePart.split(':');
    var startLetters = (rpParts[0].match(/[A-Za-z]+/) || [])[0] || 'A';
    var endLetters = (rpParts.length > 1 && (rpParts[1].match(/[A-Za-z]+/) || [])[0]) || startLetters;
    var startCol = colLetterToIndex(startLetters);
    var endCol = colLetterToIndex(endLetters);
    var numCols = Math.max(1, endCol - startCol + 1);

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [{ id: '', name: '選項讀取失敗' }];

    var values = sheet.getRange(1, startCol, lastRow, numCols).getValues();
    var arr = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (row[0] && row[1]) {
        var option = { id: String(row[0]), name: String(row[1]) };
        // For students, the 3rd column is treated as Grade (數字) and shown directly
        if (includeThirdColumn && row[2] !== undefined && row[2] !== '') {
          option.grade = row[2];
          var gnum = Number(row[2]);
          if (!isNaN(gnum)) {
            option.name = option.id + '.' + option.name + '(' + gnum + ')';
          } else {
            option.name = option.id + '.' + option.name + '(' + String(row[2]) + ')';
          }
        } else {
          option.name = option.id + '.' + option.name;
        }
        arr.push(option);
      }
    }

    try { cache.put(cacheKey, JSON.stringify(arr), 300); } catch (e) { /* ignore cache set errors */ }
    if (arr.length) return arr;
    return [{ id: '', name: '選項讀取失敗' }];
  } catch (err) {
    return [{ id: '', name: '選項讀取失敗' }];
  }
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
  
  var sheetId = getConfig().mainSheetId;
  var studentOptions = fetchSheetOptions(sheetId, 'Students!A:C', true);
  var teacherOptions = fetchSheetOptions(sheetId, 'Teachers!A:B');
  var studentName = '';
  var teacherName = '';
  var subjectName = subjectId;
  for (var i = 0; i < studentOptions.length; i++) {
    if (studentOptions[i].id === studentId) {
      studentName = studentOptions[i].name;
      var studentGrade = studentOptions[i].grade !== undefined ? studentOptions[i].grade : '';
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
      resource.extendedProperties.private.grade = studentGrade || '';
      resource.summary = studentName + '-' + currentSubject + '(' + currentTeacherName + ')';
    } else if (field === 'teacher') {
      resource.extendedProperties.private.teacher = teacherId;
      resource.summary = currentStudentName + '-' + currentSubject + '(' + teacherName + ')';
    } else if (field === 'subject') {
      resource.extendedProperties.private.subject = subjectId;
      resource.summary = currentStudentName + '-' + subjectName + '(' + currentTeacherName + ')';
    }

    if ((field === 'teacher' || field === 'student' || field === 'subject') && updateType === 'single') {
      // If nothing changed, skip patch to save API calls
      var origStudent = currentStudentId || '';
      var origTeacher = currentTeacherId || '';
      var origSubject = currentSubject || '';
      var needPatch = false;
      if (field === 'student' && studentId !== origStudent) needPatch = true;
      if (field === 'teacher' && teacherId !== origTeacher) needPatch = true;
      if (field === 'subject' && subjectId !== origSubject) needPatch = true;
      if (needPatch) {
        Calendar.Events.patch(resource, calendarId, eventId);
      }
    } else if (updateType === 'future') {
      var masterId = event && event.recurringEventId ? event.recurringEventId : null;
      if (masterId) {
        try {
          var timeMax = new Date();
          timeMax.setMonth(timeMax.getMonth() + 6);
          var instancesResp = Calendar.Events.instances(calendarId, masterId, { timeMax: timeMax.toISOString(), maxResults: 2500 });
          if (instancesResp && instancesResp.items && instancesResp.items.length) {
            var items = instancesResp.items;
            // find index of current event to start from there
            var startIndex = 0;
            for (var idx = 0; idx < items.length; idx++) {
              if (items[idx].id === eventId) { startIndex = idx; break; }
            }
            var eventsToWrite = [];
            for (var ii = startIndex; ii < items.length; ii++) {
              var inst = items[ii];
              try {
                // compute existing values
                var existingStudent = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.student) || '';
                var existingTeacher = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.teacher) || '';
                var existingSubject = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.subject) || '';

                var finalStudentId = field === 'student' ? studentId : currentStudentId;
                var finalTeacherId = field === 'teacher' ? teacherId : currentTeacherId;
                var finalSubject = field === 'subject' ? subjectId : currentSubject;

                // skip if values already match (no-op)
                if (existingStudent === finalStudentId && existingTeacher === finalTeacherId && existingSubject === finalSubject) {
                  continue;
                }

                // compute display names for summary
                var finalStudentName = '';
                var finalStudentGrade = '';
                var finalTeacherName = '';
                // find student name in studentOptions
                for (var si = 0; si < studentOptions.length; si++) {
                  if (String(studentOptions[si].id) === String(finalStudentId)) { finalStudentName = studentOptions[si].name; break; }
                }
                // also capture grade if present
                for (var sgi = 0; sgi < studentOptions.length; sgi++) {
                  if (String(studentOptions[sgi].id) === String(finalStudentId) && studentOptions[sgi].grade !== undefined) { finalStudentGrade = studentOptions[sgi].grade; break; }
                }
                // find teacher name in teacherOptions
                for (var ti = 0; ti < teacherOptions.length; ti++) {
                  if (String(teacherOptions[ti].id) === String(finalTeacherId)) { finalTeacherName = teacherOptions[ti].name; break; }
                }
                // fallback to current names if not found
                if (!finalStudentName) finalStudentName = currentStudentName || finalStudentId || '';
                if (!finalTeacherName) finalTeacherName = currentTeacherName || finalTeacherId || '';
                var resSummary = finalStudentName + '-' + (finalSubject || '') + '(' + finalTeacherName + ')';
                var resourceForInst = { summary: resSummary, extendedProperties: { private: { student: finalStudentId, teacher: finalTeacherId, subject: finalSubject, grade: finalStudentGrade || (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.grade) || '' } }, colorId: '1' };
                // patch and remember to sync later
                Calendar.Events.patch(resourceForInst, calendarId, inst.id);

                var attendanceInst = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.attendance) || '未點名';
                eventsToWrite.push({ calendarId: calendarId, eventId: inst.id, startDatetime: inst.start && (inst.start.dateTime || inst.start.date), endDatetime: inst.end && (inst.end.dateTime || inst.end.date), student: finalStudentId, teacher: finalTeacherId, subject: finalSubject, attendance: attendanceInst });
              } catch (errInst) {
                Logger.log('Failed to patch instance %s: %s', inst.id, errInst.message);
              }
            }
          }
        } catch (err2) {
          Logger.log('Error updating future instances: %s', err2.message);
          return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('儲存失敗：' + err.message))).build();
        }
      } else {
        Calendar.Events.patch(resource, calendarId, eventId);
      }
    } else {
      var masterId = event && event.recurringEventId ? event.recurringEventId : null;
      if (masterId) {
        try {
          var timeMax = new Date();
          timeMax.setMonth(timeMax.getMonth() + 6);
          var instancesResp = Calendar.Events.instances(calendarId, masterId, { timeMax: timeMax.toISOString(), maxResults: 2500 });
          if (instancesResp && instancesResp.items && instancesResp.items.length) {
            var items = instancesResp.items;
            var eventsToWrite = [];
            for (var jj = 0; jj < items.length; jj++) {
              var inst = items[jj];
              try {
                var existingStudent = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.student) || '';
                var existingTeacher = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.teacher) || '';
                var existingSubject = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.subject) || '';

                var finalStudentId = field === 'student' ? studentId : currentStudentId;
                var finalTeacherId = field === 'teacher' ? teacherId : currentTeacherId;
                var finalSubject = field === 'subject' ? subjectId : currentSubject;

                if (existingStudent === finalStudentId && existingTeacher === finalTeacherId && existingSubject === finalSubject) {
                  continue;
                }

                // compute display names for summary
                var finalStudentName = '';
                var finalTeacherName = '';
                for (var si2 = 0; si2 < studentOptions.length; si2++) {
                  if (String(studentOptions[si2].id) === String(finalStudentId)) { finalStudentName = studentOptions[si2].name; break; }
                }
                for (var ti2 = 0; ti2 < teacherOptions.length; ti2++) {
                  if (String(teacherOptions[ti2].id) === String(finalTeacherId)) { finalTeacherName = teacherOptions[ti2].name; break; }
                }
                if (!finalStudentName) finalStudentName = currentStudentName || finalStudentId || '';
                if (!finalTeacherName) finalTeacherName = currentTeacherName || finalTeacherId || '';
                var resSummary2 = finalStudentName + '-' + (finalSubject || '') + '(' + finalTeacherName + ')';
                var finalStudentGrade = '';
                for (var sgi3 = 0; sgi3 < studentOptions.length; sgi3++) {
                  if (String(studentOptions[sgi3].id) === String(finalStudentId) && studentOptions[sgi3].grade !== undefined) { finalStudentGrade = studentOptions[sgi3].grade; break; }
                }
                var resourceForInst = { summary: resSummary2, extendedProperties: { private: { student: finalStudentId, teacher: finalTeacherId, subject: finalSubject, grade: finalStudentGrade || (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.grade) || '' } }, colorId: '1' };
                Calendar.Events.patch(resourceForInst, calendarId, inst.id);

                var attendanceInst = (inst.extendedProperties && inst.extendedProperties.private && inst.extendedProperties.private.attendance) || '未點名';
                eventsToWrite.push({ calendarId: calendarId, eventId: inst.id, startDatetime: inst.start && (inst.start.dateTime || inst.start.date), endDatetime: inst.end && (inst.end.dateTime || inst.end.date), student: finalStudentId, teacher: finalTeacherId, subject: finalSubject, attendance: attendanceInst });
              } catch (errInst) {
                Logger.log('Failed to patch instance %s: %s', inst.id, errInst.message);
                return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('儲存失敗：' + errInst.message))).build();
              }
            }
          }
        } catch (err2) {
          Logger.log('Error updating recurring instances: %s', err2.message);
          return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('儲存失敗：' + err2.message))).build();
        }
      } else {
        Calendar.Events.patch(resource, calendarId, eventId);
      }
    }

    try {
      return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('已儲存欄位。'))).build();
    } catch (err) {
      return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('儲存失敗：' + err.message))).build();
    }
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
  var colorId = '1';
  // 0 日曆顏色
  // 1 薰衣草色
  // 2 鼠尾草綠
  // 3 葡萄紫
  // 4 紅鶴色
  // 5 香蕉黃
  // 6 橙橘色
  // 7 孔雀藍
  // 8 石墨黑
  // 9 藍莓色
  // 10 羅勒綠
  // 11 番茄紅
  if (attendance === '出席') colorId = '10';
  else if (attendance === '請假') colorId = '5';
  else if (attendance === '缺席') colorId = '11';
  else if (attendance === '補課') colorId = '6';
  else if (attendance === '自習') colorId = '3';
  else if (attendance === '試聽') colorId = '7';
  if (!eventId) {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().pushCard(createInfoCard('找不到事件 ID，無法儲存點名。'))).build();
  }
  try {
    var event = Calendar.Events.get(calendarId, eventId);
    var resource = { extendedProperties: { private: { attendance: attendance } }, colorId: colorId };
    Calendar.Events.patch(resource, calendarId, eventId);
    
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

      var sheetId = getConfig().mainSheetId;
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

function genId() {
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  var id = '';
  for (var i = 0; i < 10; i++) {
    id += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return id;
}
