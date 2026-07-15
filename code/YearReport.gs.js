/* =========================================================
 * 年度收益總覽
 * - 使用 Code.gs 的共用引擎 computeLedger_(移動平均成本法),
 * 數字與庫存總覽 Dashboard 完全一致
 * - 台股 (TWD) 與美股 (USD) 分開統計,避免幣別混算
 * - 需搭配 Code.gs(共用 withRetry_ / resetSheet_ / computeLedger_)
 * ========================================================= */

function createYearlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  if (!ss.getSheetByName("個股交易紀錄")) {
    ui.alert("❌ 錯誤:找不到『個股交易紀錄』分頁!");
    return;
  }

  let result;
  try {
    result = rebuildYearly_(ss, true);
  } catch (e) {
    ui.alert("❌ Google 試算表服務暫時異常,已自動重試 3 次仍失敗。\n請稍候一分鐘後再執行一次即可。\n\n技術訊息:" + e);
    return;
  }

  if (!result) {
    ui.alert("⚠ 找不到有效的交易資料,請檢查『個股交易紀錄』的日期與類型欄位。");
    return;
  }

  const ledger = result;
  if (ledger.warnings.length > 0) {
    const shown = ledger.warnings.slice(0, 10).join("\n");
    const more = ledger.warnings.length > 10 ?
      "\n...(共 " + ledger.warnings.length + " 筆異常)" : "";
    ui.alert("⚠ 資料異常提醒(不影響報表產生):\n\n" + shown + more);
  } else {
    ui.alert("🎉 年度收益總覽產生完成!");
  }
}

// 靜默重建年度收益總覽(選單 / 新增交易共用)
// 回傳 ledger;若無有效交易資料則回傳 null
function rebuildYearly_(ss, focus) {
  const recordSheet = ss.getSheetByName("個股交易紀錄");
  if (!recordSheet) throw new Error("找不到『個股交易紀錄』分頁");

  const ledger = computeLedger_(recordSheet);
  const years = Object.keys(ledger.yearly).map(Number).sort(function (a, b) { return a - b; });
  if (years.length === 0) return null;

  withRetry_(function () {
    const rptSheet = resetSheet_(ss, "年度收益總覽");
    writeYearlyReport_(rptSheet, ledger, years);
    if (focus) ss.setActiveSheet(rptSheet);
  }, "建立年度收益總覽");
  return ledger;
}

function writeYearlyReport_(rptSheet, ledger, years) {
  const colWidths = [90, 170, 170, 150, 150, 130, 90];
  colWidths.forEach(function (w, i) { rptSheet.setColumnWidth(i + 1, w); });

  rptSheet.getRange("A1").setValue("年度收益總覽(已實現損益)")
    .setFontWeight("bold").setFontSize(16).setFontColor("#1e293b");
  rptSheet.getRange("A2").setValue("僅計算已賣出部分,採移動平均成本法;未實現損益請參考『庫存總覽 Dashboard』。")
    .setFontColor("#64748b").setFontSize(10);
  const header = ["年度", "台股已實現損益 (TWD)", "美股已實現損益 (USD)", "買入總額 (混合幣別)", "賣出總額 (混合幣別)", "手續費合計", "交易筆數"];
  const headerRow = 4;
  rptSheet.getRange(headerRow, 1, 1, header.length).setValues([header])
    .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  const rows = years.map(function (yr) {
    const y = ledger.yearly[yr];
    return [yr, y.realizedTw, y.realizedUs, y.buyAmt, y.sellAmt, y.fees, y.count];
  });
  const totalRow = ["合計",
    rows.reduce(function (s, r) { return s + r[1]; }, 0),
    rows.reduce(function (s, r) { return s + r[2]; }, 0),
    rows.reduce(function (s, r) { return s + r[3]; }, 0),
    rows.reduce(function (s, r) { return s + r[4]; }, 0),
    rows.reduce(function (s, r) { return s + r[5]; }, 0),
    rows.reduce(function (s, r) { return s + r[6]; }, 0)
  ];
  const dataStart = headerRow + 1;
  rptSheet.getRange(dataStart, 1, rows.length, header.length).setValues(rows);
  const totalRowIdx = dataStart + rows.length;
  rptSheet.getRange(totalRowIdx, 1, 1, header.length).setValues([totalRow])
    .setFontWeight("bold").setBackground("#f1f5f9");

  const numRows = rows.length + 1;
  // 含合計列
  rptSheet.getRange(dataStart, 1, numRows, 1).setHorizontalAlignment("center");
  rptSheet.getRange(dataStart, 2, numRows, 2).setNumberFormat('[Red]$#,##0.00;[Green]-$#,##0.00;$0.00').setHorizontalAlignment("right");
  rptSheet.getRange(dataStart, 4, numRows, 3).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
  rptSheet.getRange(dataStart, 7, numRows, 1).setNumberFormat("#,##0").setHorizontalAlignment("center");
  // 紅漲綠跌以數字格式的 [Red]/[Green] 呈現,不使用條件式格式(本文件該服務層異常)

  // 自動擴展欄寬並套用最小寬度限制
  autoResizeColumnsWithMin_(rptSheet, 1, 7, [90, 170, 170, 150, 150, 130, 90]);

  rptSheet.getRange(totalRowIdx + 2, 1).setValue(
    "產生時間:" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm")
  ).setFontColor("#94a3b8").setFontSize(9);
}