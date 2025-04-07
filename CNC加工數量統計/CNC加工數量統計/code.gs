// 連結 HTML 檔案
function doGet(){
  var html = HtmlService.createTemplateFromFile("form");
  var check = html.evaluate();
  var show = check.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return show;
}

//登記///////////////////////////////////////////////////////////////////////////////
//抓機台
function get_mc_List() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("data");
  var data = ws.getRange(2, 1, ws.getLastRow() - 1).getValues(); // 取得 A 欄，忽略標題列
  var mcList = data.flat().filter(String); // 將二維陣列展平成一維並過濾空白
  return mcList;
}

// 抓取產品編號
function getProductList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("data");
  var data = ws.getRange(2, 2, ws.getLastRow() - 1).getValues(); // 取得 B 欄資料，忽略標題列
  var pdList = data.flat().filter(String); // 將二維陣列展平成一維並過濾空白
  return pdList;
}

// 根據選取的產品編號抓取訂單編號
function getOrderNumbers(productNumber) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("order_number");
  var data = ws.getRange(2, 1, ws.getLastRow() - 1, ws.getLastColumn()).getValues(); // 取得資料範圍
  var orderNumbers = [];

  // 遍歷資料，找到符合的產品編號並擷取其訂單編號
  data.forEach(function(row) {
    if (row[0] === productNumber) { // 假設 A 欄包含產品編號
      orderNumbers = row.slice(1); // 將該列的訂單編號加入（從第二欄開始）
    }
  });
  
  return orderNumbers.filter(String); // 過濾空白
}


//新增資料到工作表
function addData(rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("record_sheet");
  var currentDate = new Date();

  /*// 初始化 rowData
  rowData = {
    dt: "2024/10/28",
    mc: "MT02",
    pd: "0AS-025",
    hr: "11",
    min: "30",
    num: "500"
  };*/

  // 檢查工作表是否有資料（至少一行）
  if (ws.getLastRow() < 2) { // 如果沒有資料（只有標題行）
    console.log("未找到現有資料，將在第二行新增資料。");
    ws.appendRow([rowData.dt, rowData.mc, rowData.pd, rowData.or, rowData.hr, rowData.min, rowData.num, rowData.com, currentDate]);
    sortSheet(); // 新增資料後進行排序
    return true; // 成功新增資料
  }

  // 取得現有資料
  var dataRange = ws.getRange(2, 1, ws.getLastRow() - 1, 9).getValues(); // 取得 A、B、C 欄，忽略標題列

  // 格式化 rowData.dt 為 "YYYY/MM/DD"
  var inputDate = new Date(rowData.dt);
  var formattedInputDate = inputDate.getFullYear() + '/' + (inputDate.getMonth() + 1) + '/' + inputDate.getDate();

  // 調試輸出
  console.log("輸入日期: " + formattedInputDate);
  console.log("機台: " + rowData.mc);
  console.log("產品編號: " + rowData.pd);
  console.log("現有資料行數: " + dataRange.length);

  // 檢查是否已有相同的日期、機台和產品編號
  for (var i = 0; i < dataRange.length; i++) {
    // 格式化現有日期為 "YYYY/MM/DD"
    var existingDate = new Date(dataRange[i][0]);
    var formattedExistingDate = existingDate.getFullYear() + '/' + (existingDate.getMonth() + 1) + '/' + existingDate.getDate();

    console.log("檢查現有項目: " + formattedExistingDate + ", 機台: " + dataRange[i][1] + ", 產品編號: " + dataRange[i][2]);

    if (formattedExistingDate === formattedInputDate && dataRange[i][1] === rowData.mc && dataRange[i][2] === rowData.pd && dataRange[i][3] === rowData.or) {
      console.log("發現重複資料，不新增。");
      return false; // 找到重複資料，返回 false
    }
  }

  // 如果沒有重複，則新增資料
  ws.appendRow([formattedInputDate, rowData.mc, rowData.pd,rowData.or, rowData.hr, rowData.min, rowData.num, rowData.com, currentDate]);
  sortSheet(); // 新增資料後進行排序
  return true; // 成功新增資料
}

//排序資料
function sortSheet() {
  // 獲取指定名稱的工作表
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("record_sheet");
  
  // 獲取 A2 到 H 欄的最後一行範圍
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange("A2:H" + lastRow);
  
  // 根據 A 欄（1）和 C 欄（3）進行排序
  range.sort([{column: 1, ascending: true}, {column: 3, ascending: true}]);
}

//刪除最後一筆
function deleteLastRow() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("record_sheet");
    const lastRow = sheet.getLastRow();

    // 若有多於 1 列資料，則刪除最後一列並回傳 true；否則回傳 false
    if (lastRow > 1) {
        sheet.deleteRow(lastRow);
        return true;
    } else {
        return false;
    }
}

//查詢///////////////////////////////////////////////////////////////////////////////

// 抓取產品編號
function getProductList2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("data");
  var data = ws.getRange(2, 2, ws.getLastRow() - 1).getValues(); // 取得 B 欄資料，忽略標題列
  var pdList = data.flat().filter(String); // 將二維陣列展平成一維並過濾空白
  return pdList;
}

// 查詢範圍內資料
function getRecordData(dt1, dt2, pd) {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("record_sheet");
  var data = ws.getDataRange().getValues(); // 取得所有資料
  var filteredData = [];

  // 遍歷資料並根據日期和產品編號篩選
  for (var i = 0; i < data.length; i++) {
    var recordDate = new Date(data[i][0]); // A欄: 日期
    var machineCode = data[i][1]; // B欄: 機台編號
    var productCode = data[i][2]; // C欄: 產品編號
    var quantity = data[i][6]; // G欄: 數量

    // 轉換為 YYYY/M/D 格式
    var formattedDate = recordDate.getFullYear() + '/' + (recordDate.getMonth() + 1) + '/' + recordDate.getDate();
    
    // 比對條件：日期範圍和產品編號
    if (formattedDate >= dt1 && formattedDate <= dt2 && productCode === pd) {
      productCode=productCode +'  .  ' + machineCode
      filteredData.push([formattedDate, productCode, quantity]);
    }
  }
  
  return filteredData; // 回傳符合條件的資料
}

////////////////////////////////////////////////////////////////////////////////////////////////////
// 查詢依照 pd2 產品編號抓最後日期的資料
function getLastData(pd) {
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("record_sheet");
  var data = ws.getDataRange().getValues(); // 取得所有資料
  var filteredData = [];
  var latestDate = null;

  // 先找出該產品編號的最後日期
  for (var i = 0; i < data.length; i++) {
    var recordDate = new Date(data[i][0]); // A欄: 日期
    var productCode = data[i][2]; // C欄: 產品編號

    // 比對產品編號
    if (productCode === pd) {
      // 更新最新日期
      if (!latestDate || recordDate > latestDate) {
        latestDate = recordDate;
      }
    }
  }

  // 如果找到了最新日期，篩選出該日期的資料
  if (latestDate) {
    var formattedLatestDate = latestDate.getFullYear() + '/' + (latestDate.getMonth() + 1) + '/' + latestDate.getDate();

    for (var j = 0; j < data.length; j++) {
      var recordDate = new Date(data[j][0]); // A欄: 日期
      var machineCode = data[j][1]; // B欄: 機台編號
      var productCode = data[j][2]; // C欄: 產品編號
      var quantity = data[j][6]; // G欄: 數量

      // 格式化日期
      var formattedDate = recordDate.getFullYear() + '/' + (recordDate.getMonth() + 1) + '/' + recordDate.getDate();

      // 比對條件：最新日期和產品編號
      if (formattedDate === formattedLatestDate && productCode === pd) {
        productCode = productCode + '  .  ' + machineCode; // 顯示機台編號
        filteredData.push([formattedDate, productCode, quantity]);
      }
    }
  }
  
  return filteredData; // 回傳符合條件的資料
}
