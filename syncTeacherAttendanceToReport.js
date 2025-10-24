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
      // Apps Script environment may not support String.startsWith on all runtimes; coerce to string and use indexOf
      if (String(existingKey).indexOf(baseKey + '|') !== 0) continue;
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