// 連結 HTML 檔案
function doGet(){
  var html = HtmlService.createTemplateFromFile("form");
  var check = html.evaluate();
  var show = check.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return show;
}


//抓機台
function get_mc_List() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("data");
  var data = ws.getRange(2, 1, ws.getLastRow() - 1).getValues(); // 取得 A 欄，忽略標題列
  var mcList = data.flat().filter(String); // 將二維陣列展平成一維並過濾空白
  return mcList;
}

//抓產品編號
function getProductList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("data");
  var data = ws.getRange(2, 2, ws.getLastRow() - 1).getValues(); // 取得 B 欄，忽略標題列
  var pdList = data.flat().filter(String); // 將二維陣列展平成一維並過濾空白
  return pdList;
}

//新增資料到工作表
function addData(rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName("record_sheet");

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
    ws.appendRow([rowData.dt, rowData.mc, rowData.pd, rowData.hr, rowData.min, rowData.num, rowData.com]);
    return true; // 成功新增資料
  }

  // 取得現有資料
  var dataRange = ws.getRange(2, 1, ws.getLastRow() - 1, 3).getValues(); // 取得 A、B、C 欄，忽略標題列

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

    if (formattedExistingDate === formattedInputDate && dataRange[i][1] === rowData.mc && dataRange[i][2] === rowData.pd) {
      console.log("發現重複資料，不新增。");
      return false; // 找到重複資料，返回 false
    }
  }

  // 如果沒有重複，則新增資料
  ws.appendRow([formattedInputDate, rowData.mc, rowData.pd, rowData.hr, rowData.min, rowData.num, rowData.com]);
  return true; // 成功新增資料
}



