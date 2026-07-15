/* =========================================================
 * 系統設定與開關
 * ========================================================= */
// 調整欄位寬度模式：
// - 'PRESET'  : 直接套用預設最合適欄寬 (極快，推薦！完全不會有卡頓感)
// - 'DYNAMIC' : 動態依儲存格內容調整並限制最小寬度 (較慢，為原本的模式)
// - 'OFF'     : 完全關閉欄寬調整功能，保留使用者手動調整的欄寬 (不耗費任何時間)
const COLUMN_RESIZE_MODE = 'PRESET';

function onOpen() {
  buildMenu_();
}

// 動態建立選單，以呈現當前選用的介面樣式勾選狀態
function buildMenu_() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('⚡️ 股票小幫手');
  
  menu.addItem('💰 買賣輸入', 'showSidebar');
  menu.addItem('📊 建立/更新庫存總覽 Dashboard', 'createPortfolioDashboard');
  menu.addItem('📅 建立/更新年度收益總覽', 'createYearlyReport');
  menu.addSeparator();
  menu.addItem('⚙ 設定全域預設手續費率', 'promptGlobalFeeRate');
  menu.addItem('🛠️ 初始化交易紀錄分頁', 'initializeTransactionSheet');
  
  // 讀取當前的介面模式設定
  const props = PropertiesService.getUserProperties();
  const mode = props.getProperty('UI_MODE') || 'SIDEBAR';
  
  // 建立設定子選單，並在目前選中的模式前面加上勾選符號 (✅ / ⬜)
  const subMenu = ui.createMenu('⚙ 設定輸入介面樣式');
  subMenu.addItem((mode === 'SIDEBAR' ? '✅ ' : '⬜ ') + '側邊欄 (Sidebar)', 'setModeSidebar_');
  subMenu.addItem((mode === 'MODELESS' ? '✅ ' : '⬜ ') + '浮動視窗 (Modeless Dialog)', 'setModeModeless_');
  subMenu.addItem((mode === 'MODAL' ? '✅ ' : '⬜ ') + '對話框 (Modal Dialog)', 'setModeModal_');
  
  menu.addSubMenu(subMenu);
  menu.addToUi();
}

function setModeSidebar_() {
  PropertiesService.getUserProperties().setProperty('UI_MODE', 'SIDEBAR');
  buildMenu_();
  SpreadsheetApp.getActiveSpreadsheet().toast('已將輸入介面切換為：側邊欄 (Sidebar)', '⚙ 介面設定');
}

function setModeModeless_() {
  PropertiesService.getUserProperties().setProperty('UI_MODE', 'MODELESS');
  buildMenu_();
  SpreadsheetApp.getActiveSpreadsheet().toast('已將輸入介面切換為：浮動視窗 (Modeless Dialog)', '⚙ 介面設定');
}

function setModeModal_() {
  PropertiesService.getUserProperties().setProperty('UI_MODE', 'MODAL');
  buildMenu_();
  SpreadsheetApp.getActiveSpreadsheet().toast('已將輸入介面切換為：對話框 (Modal Dialog)', '⚙ 介面設定');
}

// 彈出視窗供使用者設定全域手續費率
function promptGlobalFeeRate() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getUserProperties();
  const currentRate = props.getProperty('GLOBAL_FEE_RATE') || '0.001425';
  const response = ui.prompt(
    '⚙ 設定全域手續費率',
    '請輸入預設的手續費百分比率 (小數格式)：\n\n' +
    '・台股標準手續費率為 0.001425 (即 0.1425%)\n' +
    '・若券商有手續費折扣，請自行換算後輸入，例如：\n' +
    '　　三折 → 0.001425 × 0.3 = 0.0004275\n' +
    '　　五折 → 0.001425 × 0.5 = 0.0007125\n\n' +
    '目前的設定為：' + currentRate,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText().trim();
    const rate = Number(input);
    // 用 Number 嚴格解析並改存清洗後的數字，避免「0.001425abc」這類輸入
    // 通過驗證後被直接串入 Dashboard 公式導致整欄解析錯誤
    if (input === '' || isNaN(rate) || rate < 0 || rate >= 1) {
      ui.alert('❌ 錯誤：請輸入 0 至 1 之間的小數 (例如 0.001425)。');
      return;
    }
    props.setProperty('GLOBAL_FEE_RATE', String(rate));
    ui.alert('🎉 全域預設手續費率已更新為：' + rate);
  }
}

function getGlobalFeeRate() {
  return PropertiesService.getUserProperties().getProperty('GLOBAL_FEE_RATE') || '0.001425';
}

function showSidebar() {
  // 自動檢查並建立「個股交易紀錄」分頁
  checkAndInitSheets_();

  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('新增交易紀錄')
    .setWidth(320)
    .setHeight(600); // 浮動視窗模式需要有高度設定，側邊欄則會自動忽略
    
  const mode = PropertiesService.getUserProperties().getProperty('UI_MODE') || 'SIDEBAR';
  const ui = SpreadsheetApp.getUi();
  
  if (mode === 'MODELESS') {
    ui.showModelessDialog(html, '新增交易紀錄');
  } else if (mode === 'MODAL') {
    ui.showModalDialog(html, '新增交易紀錄');
  } else {
    // 預設與 'SIDEBAR' 模式
    ui.showSidebar(html);
  }
}

// 手動觸發的初始化選單項目
function initializeTransactionSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("個股交易紀錄");
  if (sheet) {
    SpreadsheetApp.getUi().alert("ℹ️ 提示：『個股交易紀錄』分頁已存在，無需重複初始化。");
    return;
  }
  checkAndInitSheets_();
}

// 自動檢查並建立「個股交易紀錄」分頁的共用函式
function checkAndInitSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("個股交易紀錄");
  if (!sheet) {
    sheet = ss.insertSheet("個股交易紀錄");
    const headers = [
      "交易日期",
      "股票代號",
      "股票名稱",
      "交易類型",
      "交易單價",
      "交易股數",
      "手續費 / 稅金",
      "損益/收支"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold")
      .setBackground("#f1f5f9")
      .setHorizontalAlignment("center");
      
    // 凍結第一列表頭
    sheet.setFrozenRows(1);
    
    // 設定常用格式與對齊
    sheet.getRange("A2:A").setHorizontalAlignment("center");
    sheet.getRange("B2:B").setHorizontalAlignment("center");
    sheet.getRange("C2:C").setHorizontalAlignment("left");
    sheet.getRange("D2:D").setHorizontalAlignment("center");
    sheet.getRange("E2:E").setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
    sheet.getRange("F2:F").setNumberFormat("#,##0").setHorizontalAlignment("right");
    sheet.getRange("G2:G").setNumberFormat("#,##0.00").setHorizontalAlignment("right"); // 美股手續費有小數
    sheet.getRange("H2:H").setNumberFormat('[Red]$#,##0.00;[Green]-$#,##0.00;$0.00').setHorizontalAlignment("right");
    
    // 設定初始合適欄寬防止初次開啟縮排
    const initWidths = [110, 100, 150, 90, 110, 100, 110, 120];
    for (let col = 1; col <= initWidths.length; col++) {
      sheet.setColumnWidth(col, initWidths[col - 1]);
    }
    
    // 自動建立篩選器（範圍涵蓋 A1:H 欄），方便使用者個別篩選查詢
    try {
      if (!sheet.getFilter()) {
        sheet.getRange("A1:H").createFilter();
      }
    } catch (e) {
      Logger.log("自動建立篩選器失敗: " + e);
    }
    
    ss.toast('已自動建立並初始化『個股交易紀錄』分頁！', '⚡️ 系統初始化');
  }
  return sheet;
}

/* =========================================================
 * 共用工具
 * ========================================================= */

// 自動重試:遇到 Google 服務暫時性錯誤時,等待後重試最多 3 次
function withRetry_(fn, label) {
  let lastErr;
  for (let i = 0; i < 3; i++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      Logger.log("withRetry_ [" + label + "] 第 " + (i + 1) + " 次失敗: " + e);
      if (i < 2) Utilities.sleep(1500 * (i + 1)); // 最後一次失敗直接拋出,不再等待
    }
  }
  throw lastErr;
}

// 清空重用分頁(不刪除重建,避免服務競態;分頁位置也不會跑掉)
function resetSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (sheet) {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
    try { sheet.setConditionalFormatRules([]); } catch (e) { Logger.log("清除條件式格式失敗(忽略): " + e); }
    sheet.clear();
    return sheet;
  }
  return ss.insertSheet(name);
}

// 自動調整欄寬並限制最小寬度，支援不同模式以提升執行速度
function autoResizeColumnsWithMin_(sheet, startCol, endCol, minWidths) {
  try {
    const mode = (typeof COLUMN_RESIZE_MODE !== 'undefined') ? COLUMN_RESIZE_MODE : 'PRESET';
    
    if (mode === 'OFF') {
      // 關閉此功能，直接跳過以求最快速度
      return;
    }
    
    if (mode === 'DYNAMIC') {
      // 模式 1: 動態調整寬度並套用最小限制 (慢)
      sheet.autoResizeColumns(startCol, endCol - startCol + 1);
      for (let col = startCol; col <= endCol; col++) {
        const idx = col - startCol;
        const minW = minWidths[idx] || 100;
        const currentW = sheet.getColumnWidth(col);
        if (currentW < minW) {
          sheet.setColumnWidth(col, minW);
        }
      }
    } else {
      // 模式 2: PRESET (預設) - 直接套用預設的欄寬陣列 (極快，跳過動態字元長度計算)
      for (let col = startCol; col <= endCol; col++) {
        const idx = col - startCol;
        const minW = minWidths[idx] || 100;
        sheet.setColumnWidth(col, minW);
      }
    }
  } catch (e) {
    Logger.log("調整欄寬失敗: " + e);
  }
}

// 支援 Date 物件與 "2026/07/15"、"2026-07-15" 字串
function parseTxDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "string") {
    const m = v.trim().match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  return null;
}

/* =========================================================
 * 共用計算引擎:讀取交易紀錄,以「移動平均成本法」逐筆計算
 * - Dashboard 與年度收益總覽共用,確保兩邊數字一致
 * - 回傳:各股票目前部位 / 各年度損益 / 資料異常警告
 * ========================================================= */
function computeLedger_(recordSheet) {
  const result = { tickers: {}, tickerOrder: [], yearly: {}, warnings: [] };
  const lastRow = recordSheet.getLastRow();
  if (lastRow < 2) return result;
  const values = recordSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  const txs = [];
  values.forEach(function (row, i) {
    const date = parseTxDate_(row[0]);
    const ticker = String(row[1] || "").trim();
    const name = String(row[2] || "").trim();
    const type = String(row[3] || "").trim();
    const price = Number(row[4]) || 0;
    const shares = Number(row[5]) || 0;
    const fee = Number(row[6]) || 0;
    if (!date || !ticker || (type !== "買入" && type !== "賣出") || shares <= 0) return;
    txs.push({ date: date, order: i, ticker: ticker, name: name, type: type, price: price, shares: shares, fee: fee });
  });
  // 依日期排序;同日依輸入順序
  txs.sort(function (a, b) { return a.date - b.date || a.order - b.order; });
  txs.forEach(function (tx) {
    const year = tx.date.getFullYear();
    if (!result.yearly[year]) {
      result.yearly[year] = { realizedTw: 0, realizedUs: 0, buyAmt: 0, sellAmt: 0, fees: 0, count: 0 };
    }
    const y = result.yearly[year];
    const isTw = tx.ticker.indexOf("TPE:") === 0;

    if (!result.tickers[tx.ticker]) {
      result.tickers[tx.ticker] = {
        name: tx.name || tx.ticker,
        qty: 0,          // 目前持有股數
        cost: 0,           // 目前部位成本
        buyCostTotal: 0,   // 歷史累計買入成本(報酬率分母)
        realized: 0,       // 累計已實現損益
        currency: isTw ? "TWD" : "USD"
      };
      result.tickerOrder.push(tx.ticker);
    }
    const pos = result.tickers[tx.ticker];
    if (tx.name) pos.name = tx.name;

    y.count++;
    y.fees += tx.fee;

    if (tx.type === "買入") {
      const totalCost = tx.price * tx.shares + tx.fee; // 手續費計入成本
      y.buyAmt += totalCost;
      pos.cost += totalCost;
      pos.qty += tx.shares;
      pos.buyCostTotal += totalCost;
    } else {
      const revenue = tx.price * tx.shares - tx.fee;
      y.sellAmt += revenue;

      const avgCost = pos.qty > 0 ? pos.cost / pos.qty : 0;
      const matchedQty = Math.min(tx.shares, pos.qty);
      if (matchedQty < tx.shares) {
        result.warnings.push(
          Utilities.formatDate(tx.date, Session.getScriptTimeZone(), "yyyy/MM/dd") +
          " " + tx.ticker + " 賣出 " + tx.shares + " 股,但當時持有僅 " + pos.qty + " 股(超賣部分成本以 0 計算)"
        );
      }
      const costBasis = avgCost * matchedQty;
      const realized = revenue - costBasis;
      pos.cost -= costBasis;
      pos.qty -= matchedQty;
      pos.realized += realized;

      if (isTw) y.realizedTw += realized;
      else y.realizedUs += realized;
    }
  });

  return result;
}

/* =========================================================
 * 股票搜尋模組
 * - 台股(上市):證交所官方 codeQuery API → 原生繁體中文,雙向模糊查詢
 * - 台股(上櫃):櫃買中心 OpenAPI 全清單 + 本地比對(快取 6 小時)
 * - 美股:Yahoo Finance 全球搜尋(英文名稱屬正常)
 * ========================================================= */

const SEARCH_CACHE_SEC = 21600; // 6 小時
const MAX_RESULTS = 10;

function hasCJK_(str) {
  return /[\u4e00-\u9fff]/.test(str);
}

function searchTwseListed_(query) {
  const url = "https://www.twse.com.tw/rwd/zh/api/codeQuery?query=" + encodeURIComponent(query);
  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "application/json"
      }
    });
    if (res.getResponseCode() !== 200) return [];
    const json = JSON.parse(res.getContentText("UTF-8"));
    if (!json || !Array.isArray(json.suggestions)) return [];
    return json.suggestions
      .map(function (s) {
        const parts = String(s).split("\t");
        if (parts.length < 2) return null;
        return { symbol: parts[0].trim(), name: parts[1].trim(), market: "TPE" };
      })
      .filter(function (x) { return x && x.symbol && x.name; });
  } catch (e) {
    Logger.log("searchTwseListed_ Error: " + e);
    return [];
  }
}

function getOtcList_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get("otc_list_v1");
  if (hit) {
    try { return JSON.parse(hit);
    } catch (e) { /* 快取損毀則重新下載 */ }
  }
  try {
    const res = UrlFetchApp.fetch("https://www.tpex.org.tw/openapi/v1/tpex_mainboard_quotes", {
      muteHttpExceptions: true,
      headers: { "Accept": "application/json" }
    });
    if (res.getResponseCode() === 200) {
      const arr = JSON.parse(res.getContentText("UTF-8"));
      if (!Array.isArray(arr)) return [];
      const list = arr
        .map(function (r) {
          return {
            symbol: String(r.SecuritiesCompanyCode || r.Code || "").trim(),
            name: String(r.CompanyName || r.Name || "").trim(),
            market: "TPE"
          };
        })
        .filter(function (x) { return x.symbol && x.name; });
      try {
        cache.put("otc_list_v1", JSON.stringify(list), SEARCH_CACHE_SEC);
      } catch (e) { /* 超過快取大小上限就不快取,功能不受影響 */ }
      return list;
    }
  } catch (e) {
    Logger.log("getOtcList_ Error: " + e);
  }
  return [];
}

function searchOtc_(query) {
  const q = String(query).toUpperCase();
  return getOtcList_().filter(function (x) {
    return x.symbol.indexOf(q) === 0 || x.name.indexOf(query) !== -1;
  });
}

function searchYahooFinance(query) {
  if (!query) return [];
  const url = "https://query2.finance.yahoo.com/v1/finance/search?q=" + encodeURIComponent(query);
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
      }
    });
    if (response.getResponseCode() !== 200) return [];
    const json = JSON.parse(response.getContentText("UTF-8"));
    if (!json || !Array.isArray(json.quotes)) return [];
    return json.quotes
      .filter(function (q) {
        return q.symbol && (q.quoteType === "EQUITY" || q.quoteType === "ETF");
      })
      .map(function (q) {
        let symbol = q.symbol;
        let market = "美股";
        if (symbol.endsWith(".TW") || symbol.endsWith(".TWO")) {
          symbol = symbol.replace(".TW", "").replace(".TWO", "");
          market = "TPE";
        }
        return {
          symbol: symbol,
          name: q.shortname || q.longname || symbol,
          market: market
        };
      });
  } catch (e) {
    Logger.log("searchYahooFinance Error: " + e);
    return [];
  }
}

function mergeUnique_(a, b) {
  const seen = {};
  return a.concat(b).filter(function (x) {
    const key = x.market + ":" + x.symbol;
    if (seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function unifiedSearch(query) {
  if (!query) return [];
  const q = String(query).trim();
  if (!q) return [];

  const cache = CacheService.getScriptCache();
  const cacheKey = ("q_" + encodeURIComponent(q)).slice(0, 240);
  const hit = cache.get(cacheKey);
  if (hit) {
    try { return JSON.parse(hit);
    } catch (e) { /* ignore */ }
  }

  let results = [];
  const looksTaiwan = hasCJK_(q) || /^\d{3,6}[A-Z]?$/i.test(q);

  if (looksTaiwan) {
    results = mergeUnique_(searchTwseListed_(q), searchOtc_(q));
    if (results.length === 0) results = searchYahooFinance(q);
  } else {
    results = searchYahooFinance(q);
    if (results.length === 0) results = mergeUnique_(searchTwseListed_(q), searchOtc_(q));
  }

  results = results.slice(0, MAX_RESULTS);
  try { cache.put(cacheKey, JSON.stringify(results), SEARCH_CACHE_SEC);
  } catch (e) { /* ignore */ }
  return results;
}

function searchByTicker(ticker, selectedMarket) {
  if (!ticker) return null;
  const cleanTicker = String(ticker).trim().toUpperCase();
  const results = unifiedSearch(cleanTicker);
  if (results && results.length > 0) {
    const exactMarketMatch = results.find(function (r) { return r.symbol === cleanTicker && r.market === selectedMarket; });
    if (exactMarketMatch) return exactMarketMatch;
    const exactMatch = results.find(function (r) { return r.symbol === cleanTicker; });
    if (exactMatch) return exactMatch;
    const marketMatch = results.find(function (r) { return r.market === selectedMarket; });
    if (marketMatch) return marketMatch;
    return results[0];
  }
  return null;
}

function searchByName(name) {
  return unifiedSearch(name);
}

function getStockName(ticker) {
  let cleanTicker = ticker;
  let market = "美股";
  if (ticker.startsWith("TPE:")) {
    cleanTicker = ticker.replace("TPE:", "");
    market = "TPE";
  }
  const res = searchByTicker(cleanTicker, market);
  return res ? res.name : cleanTicker; // 查無結果時回傳去除 TPE: 前綴的代號
}

// 錯誤一律以 throw 回報,讓前端 failureHandler 接手(保留使用者輸入,不清空表單)
function addTransaction(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("個股交易紀錄");
  if (!sheet) throw new Error("找不到『個股交易紀錄』分頁,請先從選單執行初始化!");

  const price = parseFloat(data.price);
  const shares = parseInt(data.shares, 10);
  const fee = parseFloat(data.fee) || 0;
  if (isNaN(price) || isNaN(shares) || price <= 0 || shares <= 0) {
    throw new Error("單價與股數必須為正數!");
  }
  if (fee < 0) {
    throw new Error("手續費 / 稅金不可為負數!");
  }

  let tickerForSheet = data.ticker.trim().toUpperCase();
  if (data.market === "TPE" && !tickerForSheet.startsWith("TPE:")) {
    tickerForSheet = "TPE:" + tickerForSheet;
  }

  const stockName = data.name ? data.name : getStockName(tickerForSheet);

  // 文件鎖:避免並發送出時取得相同列號而互相覆蓋
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const nextRow = sheet.getLastRow() + 1;
    sheet.getRange(nextRow, 1, 1, 7).setValues([[
      data.date,
      tickerForSheet,
      stockName,
      data.type,
      price,
      shares,
      fee
    ]]);
    sheet.getRange(nextRow, 8).setFormula(
      `=IF(D${nextRow}="買入", -(E${nextRow}*F${nextRow}+G${nextRow}), (E${nextRow}*F${nextRow}-G${nextRow}))`
    );
    SpreadsheetApp.flush(); // 確保寫入落地後才釋放鎖
  } finally {
    lock.releaseLock();
  }

  // 自動更新 Dashboard 與年度收益總覽 (靜默更新，不切換焦點分頁)
  const failed = [];
  try {
    rebuildDashboard_(ss, false);
  } catch (e) {
    failed.push("Dashboard");
    Logger.log("自動重建 Dashboard 失敗: " + e);
  }
  try {
    rebuildYearly_(ss, false);
  } catch (e) {
    failed.push("年度收益總覽");
    Logger.log("自動重建年度收益總覽失敗: " + e);
  }

  if (failed.length > 0) {
    return "⚠ 交易已寫入,但自動更新 " + failed.join("、") + " 失敗,請稍後從選單手動重建。";
  }
  return "🎉 成功新增交易紀錄，且已自動更新 Dashboard 與年度收益總覽！";
}

/* =========================================================
 * 庫存總覽 Dashboard(穩定版)
 * - 統計數字由 Apps Script 在記憶體算好後寫入「靜態值」,
 * 試算表內只留每列一條輕量 GOOGLEFINANCE 抓現價 + 簡單算式,
 * 徹底移除原本 LET+REDUCE+LAMBDA 巨型公式造成的服務不穩定
 * - 已實現損益採移動平均成本法(與年度收益總覽一致)
 * - 摘要卡:台股 (TWD) / 美股 (USD) 分開統計
 * - L 欄為隱藏的「累計買入成本」,僅供報酬率計算
 * ========================================================= */
function createPortfolioDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // 自動初始化（如果不存在）
  checkAndInitSheets_();
  
  const recordSheet = ss.getSheetByName("個股交易紀錄");
  if (!recordSheet) {
    ui.alert("❌ 錯誤:找不到『個股交易紀錄』分頁!");
    return;
  }

  let ledger;
  try {
    ledger = rebuildDashboard_(ss, true);
  } catch (e) {
    ui.alert("❌ Google 試算表服務暫時異常,已自動重試 3 次仍失敗。\n請稍候一分鐘後再執行一次即可。\n\n技術訊息:" + e);
    return;
  }

  let msg = "🎉 庫存總覽 Dashboard 建立完成!(台股 TWD / 美股 USD 分開統計)";
  if (ledger.warnings.length > 0) {
    msg += "\n\n⚠ 發現 " + ledger.warnings.length + " 筆資料異常(如超賣),已以持有股數為上限計算,詳見執行紀錄。";
    ledger.warnings.forEach(function (w) { Logger.log("資料異常: " + w); });
  }
  ui.alert(msg);
}

// 靜默重建 Dashboard(選單 / 新增交易共用);回傳 ledger
function rebuildDashboard_(ss, focus) {
  const recordSheet = ss.getSheetByName("個股交易紀錄");
  if (!recordSheet) throw new Error("找不到『個股交易紀錄』分頁");

  // 步驟 1:先在記憶體完成所有計算(此階段不寫入任何分頁,失敗不留半成品)
  const ledger = computeLedger_(recordSheet);
  const holdings = ledger.tickerOrder
    .map(function (t) { return { ticker: t, info: ledger.tickers[t] }; })
    .filter(function (x) { return x.info.buyCostTotal > 0; })
    .sort(function (a, b) {
      if (a.info.currency !== b.info.currency) return a.info.currency === "TWD" ? -1 : 1;
      return a.ticker < b.ticker ? -1 : (a.ticker > b.ticker ? 1 : 0);
    });
  // 步驟 2:重設分頁並寫入(自動重試)
  withRetry_(function () {
    const dbSheet = resetSheet_(ss, "庫存總覽 Dashboard");
    writeDashboard_(dbSheet, holdings);
    if (focus) ss.setActiveSheet(dbSheet);
  }, "建立庫存總覽 Dashboard");
  return ledger;
}

function writeDashboard_(dbSheet, holdings) {
  const DATA_START = 7;
  // 第 6 列表頭,第 7 列起為資料
  // (欄寬統一由結尾的 autoResizeColumnsWithMin_ 設定)

  // ---- 摘要卡 ----
  dbSheet.getRange("A1:K4").setBackground("#f8fafc");
  // 台股 (TWD) — 第 1~2 列
  dbSheet.getRange("A1").setValue("🇹🇼 台股持倉市值 (TWD)").setFontWeight("bold").setFontColor("#64748b");
  dbSheet.getRange("A2").setFormula('=SUMIFS(F7:F, K7:K, "TWD")')
    .setFontWeight("bold").setFontSize(14).setNumberFormat('"NT$"#,##0');

  dbSheet.getRange("C1").setValue("台股未實現損益").setFontWeight("bold").setFontColor("#64748b");
  dbSheet.getRange("C2").setFormula('=SUMIFS(G7:G, K7:K, "TWD")')
    .setFontWeight("bold").setFontSize(14).setNumberFormat('[Red]"NT$"#,##0;[Green]-"NT$"#,##0;"NT$"0');

  dbSheet.getRange("E1").setValue("台股已實現損益").setFontWeight("bold").setFontColor("#64748b");
  dbSheet.getRange("E2").setFormula('=SUMIFS(H7:H, K7:K, "TWD")')
    .setFontWeight("bold").setFontSize(14).setNumberFormat('[Red]"NT$"#,##0;[Green]-"NT$"#,##0;"NT$"0');

  dbSheet.getRange("G1").setValue("台股總損益 (TWD)").setFontWeight("bold").setFontColor("#1e293b");
  dbSheet.getRange("G2").setFormula("=C2+E2")
    .setFontWeight("bold").setFontSize(16).setNumberFormat('[Red]"NT$"#,##0;[Green]-"NT$"#,##0;"NT$"0');

  // 美股 (USD) — 第 3~4 列
  dbSheet.getRange("A3").setValue("🇺🇸 美股持倉市值 (USD)").setFontWeight("bold").setFontColor("#64748b");
  dbSheet.getRange("A4").setFormula('=SUMIFS(F7:F, K7:K, "USD")')
    .setFontWeight("bold").setFontSize(14).setNumberFormat('"US$"#,##0.00');

  dbSheet.getRange("C3").setValue("美股未實現損益").setFontWeight("bold").setFontColor("#64748b");
  dbSheet.getRange("C4").setFormula('=SUMIFS(G7:G, K7:K, "USD")')
    .setFontWeight("bold").setFontSize(14).setNumberFormat('[Red]"US$"#,##0.00;[Green]-"US$"#,##0.00;"US$"0.00');

  dbSheet.getRange("E3").setValue("美股已實現損益").setFontWeight("bold").setFontColor("#64748b");
  dbSheet.getRange("E4").setFormula('=SUMIFS(H7:H, K7:K, "USD")')
    .setFontWeight("bold").setFontSize(14).setNumberFormat('[Red]"US$"#,##0.00;[Green]-"US$"#,##0.00;"US$"0.00');

  dbSheet.getRange("G3").setValue("美股總損益 (USD)").setFontWeight("bold").setFontColor("#1e293b");
  dbSheet.getRange("G4").setFormula("=C4+E4")
    .setFontWeight("bold").setFontSize(16).setNumberFormat('[Red]"US$"#,##0.00;[Green]-"US$"#,##0.00;"US$"0.00');
  // ---- 表頭 ----
  const header = ["股票代號", "股票名稱", "目前現價", "持有股數", "平均買入成本", "目前市值", "未實現損益", "已實現損益", "累計總損益", "總報酬率", "幣別", "累計買入成本"];
  dbSheet.getRange(6, 1, 1, header.length).setValues([header])
    .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center");

  // ---- 資料列:靜態值 + 輕量公式 ----
  if (holdings.length > 0) {
    const staticRows = holdings.map(function (h) {
      const info = h.info;
      const avgCost = info.qty > 0 ? info.cost / info.qty : 0;
      // C/F/G/I/J 先留空,稍後以公式覆蓋
      return [h.ticker, info.name, "", info.qty, avgCost, "", "", info.realized, "", "", info.currency, info.buyCostTotal];
    });
    dbSheet.getRange(DATA_START, 1, staticRows.length, 12).setValues(staticRows);

    const fPrice = [], fMktVal = [], fUnreal = [], fTotal = [], fRoi = [];
    // 防呆:舊版可能存過非純數字的費率字串,解析失敗就退回台股法定費率
    let feeRate = parseFloat(getGlobalFeeRate());
    if (isNaN(feeRate) || feeRate < 0) feeRate = 0.001425;
    for (let i = 0; i < holdings.length; i++) {
      const r = DATA_START + i;
      // 現價抓不到時留空(而非當 0),避免未實現損益顯示成全額虧損
      fPrice.push(['=IFERROR(GOOGLEFINANCE($A' + r + ',"price"),"")']);
      fMktVal.push(['=IF(ISNUMBER($C' + r + '),C' + r + '*D' + r + ',"")']);
      // 扣除預估賣出手續費與證交稅以對齊券商 APP 淨損益,僅台股適用
      // (一般股票 0.3%,ETF 代號 00 開頭 0.1%;美股多為零手續費,不預扣)
      const sellCostRate = 'IF(LEFT($A' + r + ',4)="TPE:",' + feeRate + '+IF(MID($A' + r + ',5,2)="00",0.001,0.003),0)';
      fUnreal.push(['=IF(ISNUMBER($C' + r + '),F' + r + '-D' + r + '*E' + r + '-F' + r + '*' + sellCostRate + ',"")']);
      // SUM 會自動略過留空的未實現損益,現價抓不到時仍顯示已實現部分
      fTotal.push(['=SUM(G' + r + ',H' + r + ')']);
      fRoi.push(['=IF($L' + r + '=0,0,I' + r + '/$L' + r + ')']);
    }
    dbSheet.getRange(DATA_START, 3, holdings.length, 1).setFormulas(fPrice);   // C 現價
    dbSheet.getRange(DATA_START, 6, holdings.length, 1).setFormulas(fMktVal);
    // F 市值
    dbSheet.getRange(DATA_START, 7, holdings.length, 1).setFormulas(fUnreal);  // G 未實現
    dbSheet.getRange(DATA_START, 9, holdings.length, 1).setFormulas(fTotal);
    // I 累計總損益
    dbSheet.getRange(DATA_START, 10, holdings.length, 1).setFormulas(fRoi);
    // J 報酬率
  }

  // ---- 數字格式 ----
  const endRow = Math.max(DATA_START + holdings.length - 1, 120);
  dbSheet.getRange("C7:J" + endRow).setHorizontalAlignment("right");
  dbSheet.getRange("C7:C" + endRow).setNumberFormat("$#,##0.00");
  dbSheet.getRange("D7:D" + endRow).setNumberFormat("#,##0");
  dbSheet.getRange("E7:F" + endRow).setNumberFormat("$#,##0.00");
  dbSheet.getRange("G7:I" + endRow).setNumberFormat('[Red]$#,##0.00;[Green]-$#,##0.00;$0.00');
  dbSheet.getRange("J7:J" + endRow).setNumberFormat('[Red]0.00%;[Green]-0.00%;0.00%');
  dbSheet.getRange("K7:K" + endRow).setHorizontalAlignment("center").setFontColor("#64748b");

  // 自動擴展欄寬（限制最小寬度防止 GOOGLEFINANCE 載入中縮排）
  autoResizeColumnsWithMin_(dbSheet, 1, 11, [110, 160, 100, 90, 110, 120, 120, 120, 120, 100, 70]);

  dbSheet.hideColumns(12); // L 欄僅供報酬率計算,隱藏不顯示

  // 紅漲綠跌:一律以數字格式的 [Red]/[Green] 呈現。
  // 本文件的條件式格式服務層已確認異常(系統診斷定位),故正式功能完全不使用條件式格式。
}