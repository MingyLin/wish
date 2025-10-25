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

  // 取得 Student ID->Grade 對照表（如果有 Students 分頁）
  var studentGradeMap = {};
  var studentsSheet = srcSs.getSheetByName('Students');
  if (studentsSheet) {
    var sData = studentsSheet.getDataRange().getValues();
    if (sData.length > 1) {
      var sHeader = sData[0] || [];
      var gradeIdx = -1;
      var idIdx = 0;
      for (var sh = 0; sh < sHeader.length; sh++) {
        var key = String(sHeader[sh] || '').toLowerCase();
        if (key === 'grade' || key === '年級' || key === 'gradelevel') gradeIdx = sh;
        if (key === 'id' || key === 'student' || key === 'studentid' || key === '學號') idIdx = sh;
      }
      if (gradeIdx === -1) gradeIdx = 2; // fallback to column C if header not explicit
      for (var sr = 1; sr < sData.length; sr++) {
        var srow = sData[sr];
        var sid = srow[idIdx] !== undefined ? String(srow[idIdx]) : '';
        var gval = srow[gradeIdx];
        var gnum = Number(gval);
        if (sid) studentGradeMap[sid] = isNaN(gnum) ? String(gval) : gnum;
      }
    }
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
  var sti = idx['Student'] || idx['student'] || idx['StudentId'] || idx['studentId'] || idx['學號'];
  var si = idx['StartDatetime'] || idx['startDatetime'] || idx['startdatetime'];
  var ei = idx['EndDatetime'] || idx['endDatetime'] || idx['enddatetime'];

  // 過濾出席資料，並依老師+日期分組，只合併重疊或連續時段
  var groupMap = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var attendanceVal = row[ai] ? String(row[ai]).trim() : '';
    // include both '出席' and '試聽'
    if (attendanceVal !== '出席' && attendanceVal !== '試聽') continue;
    var teacherId = row[ti] ? String(row[ti]) : '';
    var teacherName = zeroPad && typeof zeroPad === 'function' ? zeroPad(teacherId, 3) + ' ' + (teacherMap[teacherId] || teacherId) : (teacherMap[teacherId] || teacherId);
    var studentId = sti !== undefined ? (row[sti] ? String(row[sti]) : '') : '';
    var gradeVal = studentId && studentGradeMap[studentId] !== undefined ? studentGradeMap[studentId] : '';
    var gradeGroupVal = getGradeGroup(gradeVal);
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
    
  var baseKey = teacherName + '|' + gradeGroupVal + '|' + dateKey;
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
      groupMap[newKey] = { teacher: teacherName, grade: gradeGroupVal, date: dateKey, start: startStr, end: endStr };
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
  // build sortable array and include grade
  var tmp = [];
  for (var k in groupMap) {
    var g = groupMap[k];
    var hour = calcHour(g.start, g.end);
    var startFmt = formatDateTimeStr(g.start);
    var endFmt = formatDateTimeStr(g.end);
    tmp.push({ teacher: String(g.teacher), grade: String(g.grade || ''), startRaw: g.start, startFmt: String(startFmt), endFmt: String(endFmt), hour: String(hour) });
  }

  // sort by teacher, grade (小, 中, 大, 其他), start time
  var gradeOrder = { '小': 0, '中': 1, '大': 2 };
  tmp.sort(function(a, b) {
    if (a.teacher < b.teacher) return -1;
    if (a.teacher > b.teacher) return 1;
    var ga = gradeOrder[a.grade] !== undefined ? gradeOrder[a.grade] : 99;
    var gb = gradeOrder[b.grade] !== undefined ? gradeOrder[b.grade] : 99;
    if (ga < gb) return -1;
    if (ga > gb) return 1;
    return compareTime(a.startRaw, b.startRaw);
  });

  for (var tii = 0; tii < tmp.length; tii++) {
    var e = tmp[tii];
    out.push([e.teacher, e.grade, e.startFmt, e.endFmt, e.hour]);
  }

  // 寫入目標 teacher 分頁
  var destId = '1EsDKBw0malhT8o0s_OF28QMnNfScTerF-FK-9W36GOE';
  var destSs = SpreadsheetApp.openById(destId);
  var destSheet = destSs.getSheetByName('teacher');
  if (!destSheet) destSheet = destSs.insertSheet('teacher');
  destSheet.clear();
  destSheet.appendRow(['老師', 'Grade', '開始時間', '結束時間', '時數']);
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

function formatDateTimeStr(dtVal) {
  var dt = null;
  
  // 若是 Date 物件
  if (Object.prototype.toString.call(dtVal) === '[object Date]' && !isNaN(dtVal.getTime())) {
    dt = new Date(dtVal.getTime());
  }
  else if (typeof dtVal === 'string') {
    var tempDt = new Date(dtVal);
    if (!isNaN(tempDt.getTime())) {
      dt = new Date(tempDt.getTime());
    }
  }
  else if (typeof dtVal === 'number') {
    var tempDt = new Date(Math.round((dtVal - 25569) * 86400 * 1000));
    if (!isNaN(tempDt.getTime())) {
      dt = new Date(tempDt.getTime());
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

// Map a student grade value to a grade group string: '小' (<=6), '中' (7-9), '大' (>=10)
function getGradeGroup(gradeVal) {
  if (gradeVal === null || gradeVal === undefined || gradeVal === '') return '';
  var n = Number(gradeVal);
  if (!isNaN(n)) {
    if (n <= 6) return '小';
    if (n >= 7 && n <= 9) return '中';
    if (n >= 10) return '高';
  }
  // If it's already one of the labels, return normalized
  var s = String(gradeVal).trim();
  if (s === '小' || s === '中' || s === '高') return s;
  return '';
}