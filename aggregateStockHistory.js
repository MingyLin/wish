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
