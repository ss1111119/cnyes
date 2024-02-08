function parseNewsData() {
  var sheetName = "鉅亨網新聞";
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  Logger.log('處理工作表: ' + sheetName);
  
  // 設置開始和結束時間戳
  var startDateStr = "20240207"; // 需要轉換的開始日期
  var endDateStr = "20240208"; // 需要轉換的結束日期

  var startAt = dateToTimestamp(startDateStr); // 轉換後的開始時間戳
  var endAt = dateToTimestamp(endDateStr); // 轉換後的結束時間戳
  Logger.log('開始時間戳: ' + startAt + ', 結束時間戳: ' + endAt);

  var limit = 30; // 每頁新聞數量
  var page = 1; // 初始頁碼
  var hasMore = true; // 標記是否還有更多頁面
  
  // 讀取工作表中已存在的數據
  var existingTitles = new Set();
  var lastRow = sheet.getLastRow();
  var existingData = lastRow > 0 ? sheet.getRange(1, 1, lastRow, sheet.getLastColumn()).getValues() : [];
  Logger.log('已存在數據行數: ' + lastRow);
  
  existingData.forEach(function(row) {
    existingTitles.add(row[1]); // 假設標題在第二列
  });

  while (hasMore) {
    var apiUrl = `https://news.cnyes.com/api/v3/news/category/tw_stock?startAt=${startAt}&endAt=${endAt}&limit=${limit}&page=${page}`;
    Logger.log('正在從URL獲取數據: ' + apiUrl);
    var response = UrlFetchApp.fetch(apiUrl);
    var jsonData = JSON.parse(response.getContentText());

    if (!jsonData || !jsonData.items || !jsonData.items.data || jsonData.items.data.length === 0) {
      Logger.log('第 ' + page + ' 頁沒有更多新聞數據');
      break;
    }

    var rowsToAdd = [];
    jsonData.items.data.forEach(function(item) {
      var title = item.title;
      // 檢查是否已存在
      if (!existingTitles.has(title)) {
        var publishDate = new Date(item.publishAt * 1000);
        var formattedDate = Utilities.formatDate(publishDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        var keywords = item.keyword ? item.keyword.join(", ") : "";
        var summary = item.summary;
        var link = "https://news.cnyes.com/news/id/" + item.newsId;
        // 修改這部分，從 market 陣列中提取 code 值
        var codes = item.market ? item.market.map(function(market) { return market.code; }).join(", ") : "";
        
        rowsToAdd.push([formattedDate, title, keywords, summary, codes, link]);
        existingTitles.add(title); // 添加到已存在的標題集合中，避免未來重複
      }
    });

    // 將新行批量添加到工作表
    if (rowsToAdd.length > 0) {
      sheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
      lastRow += rowsToAdd.length; // 更新最後一行的編號，為可能的下一次迭代做準備
      Logger.log(rowsToAdd.length + ' 行新數據已添加');
    } else {
      Logger.log('沒有新數據需要添加');
    }

    page += 1;
    hasMore = jsonData.items.data.length === limit;
  }
}

function dateToTimestamp(dateStr) {
  // 將日期字符串從"YYYYMMDD"格式轉換為"YYYY-MM-DD"
  var formattedDateStr = dateStr.substring(0, 4) + '-' + dateStr.substring(4, 6) + '-' + dateStr.substring(6, 8);

  // 創建一個Date對象
  var date = new Date(formattedDateStr + 'T00:00:00Z'); // 加'T00:00:00Z'確保時間為UTC

  // 獲取UNIX時間戳（秒為單位）
  var timestamp = date.getTime() / 1000;

  return timestamp.toString();
}
