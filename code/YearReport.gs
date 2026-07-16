/* =========================================================
 * 年度收益總覽
 * - 使用 Code.gs 的共用引擎 computeLedger_(移動平均成本法),
 * 數字與庫存總覽 Dashboard 完全一致
 * - 台股 (TWD) 與美股 (USD) 分開統計,避免幣別混算
 * - 頂部為全量 KPI 摘要卡;年度總計行下可展開該年度
 * 有賣出交易的個股明細(列群組折疊)
 * - 需搭配 Code.gs(共用 withRetry_ / resetSheet_ / computeLedger_)
 * ========================================================= */

function createYearlyReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  // 自動初始化（如果不存在）
  if (typeof checkAndInitSheets_ === "function") {
    checkAndInitSheets_();
  }

  const recordSheet = ss.getSheetByName("個股交易紀錄");
  if (!recordSheet) {
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
    // sheet.clear() 不會移除列群組,須趁內容還在時先清除,
    // 否則重複執行會使 shiftRowGroupDepth 的深度不斷疊加
    const existing = ss.getSheetByName("年度收益總覽");
    if (existing) removeAllRowGroups_(existing);
    const rptSheet = resetSheet_(ss, "年度收益總覽");
    writeYearlyReport_(rptSheet, ledger, years);
    if (focus) ss.setActiveSheet(rptSheet);
  }, "建立年度收益總覽");
  return ledger;
}

// 移除分頁上所有列群組;先全部展開,避免收合中的群組移除後資料列仍被隱藏
function removeAllRowGroups_(sheet) {
  try {
    sheet.expandAllRowGroups();
    const maxRow = sheet.getLastRow();
    for (let row = 1; row <= maxRow; row++) {
      while (sheet.getRowGroupDepth(row) > 0) {
        const group = sheet.getRowGroup(row, 1);
        if (!group) break;
        group.remove();
      }
    }
  } catch (e) {
    Logger.log("移除既有列群組失敗(忽略): " + e);
  }
}

function writeYearlyReport_(rptSheet, ledger, years) {
  const NUM_COLS = 6; // A 年度/代號 + 已實現損益/買入/賣出/手續費/筆數
  // A 欄需容納縮排的個股標籤(如「   TPE:2330 (台積電)」),故較寬
  const colWidths = [180, 140, 130, 130, 110, 90];
  colWidths.forEach(function (w, i) { rptSheet.setColumnWidth(i + 1, w); });

  // sheet.clear() 不會解除合併儲存格,先整區還原,避免與舊版面殘留的 merge 衝突
  try { rptSheet.getRange(1, 1, 500, 11).breakApart(); } catch (e) { Logger.log("解除舊合併儲存格失敗(忽略): " + e); }

  // ---- 頂部大字報:全量 KPI 摘要卡 (A1:G4) ----
  rptSheet.getRange("A1:G4").setBackground("#f8fafc");

  let totalRealizedTw = 0;
  let totalRealizedUs = 0;
  Object.keys(ledger.tickers).forEach(function (t) {
    const info = ledger.tickers[t];
    if (info.currency === "TWD") totalRealizedTw += info.realized;
    else totalRealizedUs += info.realized;
  });
  const totalFees = years.reduce(function (s, yr) { return s + ledger.yearly[yr].fees; }, 0);
  const totalCount = years.reduce(function (s, yr) { return s + ledger.yearly[yr].count; }, 0);

  rptSheet.getRange("A1").setValue("🇹🇼 台股歷史已實現 (TWD)").setFontWeight("bold").setFontColor("#64748b");
  rptSheet.getRange("A2").setValue(totalRealizedTw).setFontWeight("bold").setFontSize(14)
    .setNumberFormat('[Red]"NT$"#,##0;[Green]-"NT$"#,##0;"NT$"0');

  rptSheet.getRange("C1").setValue("🇺🇸 美股歷史累計已實現 (USD)").setFontWeight("bold").setFontColor("#64748b");
  rptSheet.getRange("C2").setValue(totalRealizedUs).setFontWeight("bold").setFontSize(14)
    .setNumberFormat('[Red]"US$"#,##0.00;[Green]-"US$"#,##0.00;"US$"0.00');

  rptSheet.getRange("E1").setValue("歷史累計手續費").setFontWeight("bold").setFontColor("#64748b");
  rptSheet.getRange("E2").setValue(totalFees).setFontWeight("bold").setFontSize(14)
    .setNumberFormat("$#,##0.00");

  rptSheet.getRange("G1").setValue("歷史累計交易筆數").setFontWeight("bold").setFontColor("#64748b");
  rptSheet.getRange("G2").setValue(totalCount).setFontWeight("bold").setFontSize(14)
    .setNumberFormat("#,##0");

  // 產生時間(置於 KPI 摘要卡下方)
  rptSheet.getRange("A3").setValue(
    "產生時間:" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm")
  ).setFontColor("#94a3b8").setFontSize(9);

  // ---- 台股/美股上下分區:各自獨立的年度表格 ----
  const groupRanges = []; // 全部明細列群組範圍(絕對列號),兩個分區共用

  // 渲染單一市場分區(分區標題 + 欄位表頭 + 年度行/明細行),回傳實際使用到的最後一列列號
  function renderSection_(startRow, title, currency) {
    rptSheet.getRange(startRow, 1, 1, NUM_COLS).merge()
      .setValue(title)
      .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold")
      .setFontSize(12).setHorizontalAlignment("center").setVerticalAlignment("middle");
    rptSheet.getRange(startRow + 1, 1, 1, NUM_COLS)
      .setValues([["年度", "已實現損益", "買入總額", "賣出總額", "手續費", "交易筆數"]])
      .setBackground("#334155").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

    const dataStart = startRow + 2;
    const rows = [];
    const yearRowIdxs = [];
    const sectionGroups = [];
    years.forEach(function (yr) {
      const y = ledger.yearly[yr];
      const agg = { realized: 0, buyAmt: 0, sellAmt: 0, fees: 0, count: 0 };
      const detailTickers = [];
      Object.keys(y.tickers).forEach(function (t) {
        const yt = y.tickers[t];
        if (yt.currency !== currency) return;
        agg.realized += yt.realized;
        agg.buyAmt += yt.buyAmt;
        agg.sellAmt += yt.sellAmt;
        agg.fees += yt.fees;
        agg.count += yt.count;
        // 只列出該年度有賣出的個股(含損益剛好打平為 0 者);有買無賣者不列。
        // realized 只在賣出時變動,加上此檢查可涵蓋售價 0 的沖銷(revenue 為 0)極端情況
        if (yt.sellAmt !== 0 || yt.realized !== 0) detailTickers.push(t);
      });
      if (agg.count === 0) return; // 該市場該年度無交易,不列該年度

      const parentRowIdx = dataStart + rows.length; // 年度行的實際列號
      yearRowIdxs.push(parentRowIdx);
      rows.push([yr, agg.realized, agg.buyAmt, agg.sellAmt, agg.fees, agg.count]);

      detailTickers.sort().forEach(function (t) {
        const yt = y.tickers[t];
        rows.push(["   " + yt.ticker + " (" + yt.name + ")", yt.realized, yt.buyAmt, yt.sellAmt, yt.fees, yt.count]);
      });
      if (detailTickers.length > 0) {
        sectionGroups.push({ start: parentRowIdx + 1, count: detailTickers.length });
      }
    });

    if (rows.length === 0) {
      rptSheet.getRange(dataStart, 1).setValue("（尚無交易紀錄）")
        .setFontColor("#94a3b8").setFontSize(10).setHorizontalAlignment("left");
      return dataStart;
    }

    rptSheet.getRange(dataStart, 1, rows.length, NUM_COLS).setValues(rows);

    // ---- 數字格式與對齊 ----
    const numRows = rows.length;
    rptSheet.getRange(dataStart, 1, numRows, 1).setHorizontalAlignment("center");
    // 已實現損益:紅漲綠跌(以數字格式呈現,不使用條件式格式,本文件該服務層異常)
    rptSheet.getRange(dataStart, 2, numRows, 1).setNumberFormat('[Red]$#,##0.00;[Green]-$#,##0.00;$0.00').setHorizontalAlignment("right");
    rptSheet.getRange(dataStart, 3, numRows, 3).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
    rptSheet.getRange(dataStart, 6, numRows, 1).setNumberFormat("#,##0").setHorizontalAlignment("center");

    // 年度行加粗、明細行縮小灰階;明細 A 欄靠左(置中會使縮排空白失效)
    yearRowIdxs.forEach(function (r) {
      rptSheet.getRange(r, 1, 1, NUM_COLS).setFontWeight("bold").setFontSize(11);
    });
    sectionGroups.forEach(function (g) {
      rptSheet.getRange(g.start, 1, g.count, NUM_COLS)
        .setFontSize(10).setFontColor("#64748b").setBackground("#f8fafc");
      rptSheet.getRange(g.start, 1, g.count, 1).setHorizontalAlignment("left");
      groupRanges.push(g);
    });

    return dataStart + rows.length - 1;
  }

  const twEnd = renderSection_(6, "🇹🇼 台股 (TWD)", "TWD");
  // 美股分區固定從第 11 列起;台股資料較多時自動往下順延(與台股區隔一列空白)
  const usStart = Math.max(11, twEnd + 2);
  renderSection_(usStart, "🇺🇸 美股 (USD)", "USD");

  // ---- 建立折疊列群組(折疊鈕顯示在年度行) ----
  // 分組僅屬視覺輔助:任一群組 API 失敗只記錄不中斷,明細行仍完整呈現
  if (groupRanges.length > 0) {
    try {
      rptSheet.setRowGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.BEFORE);
    } catch (e) {
      Logger.log("設定列群組折疊鈕位置失敗(忽略,折疊鈕將顯示於群組下方): " + e);
    }
    groupRanges.forEach(function (g) {
      try {
        rptSheet.getRange(g.start, 1, g.count, 1).shiftRowGroupDepth(1);
      } catch (e) {
        Logger.log("建立列群組失敗(忽略,明細仍會顯示): " + e);
      }
    });
  }

  // 自動擴展欄寬並套用最小寬度限制
  autoResizeColumnsWithMin_(rptSheet, 1, 6, colWidths);
}
