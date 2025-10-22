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