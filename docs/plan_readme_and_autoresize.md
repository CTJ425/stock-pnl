# 實作計畫：建立 README.md 與實作表格欄寬自動擴展功能

## 目的與問題描述
1. **建立說明文件**：新增 [README.md](file:///mnt/c/Users/user/Desktop/股票/README.md) 用以詳細說明此 Google Apps Script 股票小幫手系統的安裝、設定與使用方法。
2. **自動擴展表格（自動欄寬調整）**：目前 Dashboard 與年度收益表使用固定的欄寬。當遇到較長的股票名稱或較大的損益金額時，容易造成文字折行或數字顯示為 `###` 的情況。
   * *解決方案*：我們將在 `code.gs.js` 中實作一個通用工具函式 `autoResizeColumnsWithMin_`。它會呼叫 Google Sheets 的 `autoResizeColumns` API 自動調整欄寬，同時套用**最小欄寬限制**（Min Widths），防止 `GOOGLEFINANCE` 函式在加載中（Loading 狀態）時導致欄寬過度縮小。

---

## 使用者審查項目 (User Review Required)
> [!NOTE]
> **最小寬度設計**：
> 由於 `GOOGLEFINANCE` 獲取現價是非同步的，剛寫入公式時可能呈現為空或載入中，這會使 `autoResizeColumns` 把該欄縮得極小。因此，我們設定了預設的「最小安全寬度」，確保即使在載入中，版面依然整齊且不出現 `###`。

---

## 預計修改內容

### 1. 新增 README.md 說明文件
#### [NEW] [README.md](file:///mnt/c/Users/user/Desktop/股票/README.md)
建立一個完整的中文使用手冊，涵蓋功能介紹、試算表格式要求、安裝步驟及常見問題。

---

### 2. 表格自動欄寬擴展 (Auto-resize with Min Width)

#### [MODIFY] [code.gs.js](file:///mnt/c/Users/user/Desktop/股票/code.gs.js)
* 在「共用工具」區塊新增 `autoResizeColumnsWithMin_` 輔助函式。
* 修改 `writeDashboard_` 函式，在寫入資料與公式後，調用 `autoResizeColumnsWithMin_` 進行欄寬自適應。

**程式碼修改點 1 (約第 48 行，新增共用工具)：**
```javascript
// 自動調整欄寬並限制最小寬度，防止 GoogleFinance 載入中導致欄位縮小
function autoResizeColumnsWithMin_(sheet, startCol, endCol, minWidths) {
  try {
    sheet.autoResizeColumns(startCol, endCol - startCol + 1);
    for (let col = startCol; col <= endCol; col++) {
      const idx = col - startCol;
      const minW = minWidths[idx] || 100;
      const currentW = sheet.getColumnWidth(col);
      if (currentW < minW) {
        sheet.setColumnWidth(col, minW);
      }
    }
  } catch (e) {
    Logger.log("自動調整欄寬失敗: " + e);
  }
}
```

**程式碼修改點 2 (約第 421 - 426 行，修改 `writeDashboard_` 中設定寬度的邏輯)：**
```javascript
function writeDashboard_(dbSheet, holdings) {
  const DATA_START = 7;
  // 第 6 列表頭,第 7 列起為資料

  const colWidths = [110, 160, 100, 90, 110, 120, 120, 120, 120, 100, 70];
  // 先設定基礎寬度
  colWidths.forEach(function (w, i) { dbSheet.setColumnWidth(i + 1, w); });
```
*在 `writeDashboard_` 底部（約第 508 行後，隱藏 L 欄前）加入：*
```javascript
  // 自動擴展欄寬（限制最小寬度防止 GOOGLEFINANCE 載入中縮排）
  autoResizeColumnsWithMin_(dbSheet, 1, 11, colWidths);

  dbSheet.hideColumns(12); // L 欄僅供報酬率計算,隱藏不顯示
```

---

#### [MODIFY] [YearReport.gs.js](file:///mnt/c/Users/user/Desktop/股票/YearReport.gs.js)
* 修改 `writeYearlyReport_` 函式，在寫入年度報告後調用 `autoResizeColumnsWithMin_` 進行欄寬自適應。

**程式碼修改點 (約第 97 行後)：**
```javascript
  rptSheet.getRange(dataStart, 7, numRows, 1).setNumberFormat("#,##0").setHorizontalAlignment("center");
  // 紅漲綠跌以數字格式的 [Red]/[Green] 呈現,不使用條件式格式(本文件該服務層異常)

  // 自動擴展欄寬並套用最小寬度限制
  autoResizeColumnsWithMin_(rptSheet, 1, 7, colWidths);

  rptSheet.getRange(totalRowIdx + 2, 1).setValue(
```

---

## 驗證計畫 (Verification Plan)
1. **程式碼完整性**：確認 `autoResizeColumnsWithMin_` 能在 `code.gs.js` 全域範圍被 `YearReport.gs.js` 呼叫。
2. **手動測試**：
   * 在交易紀錄中填入一個**名稱特別長**的股票（如美股 `Direxion Daily Semiconductor Bull 3X Shares`）。
   * 點擊自訂選單或透過側邊欄新增交易。
   * 檢查 Dashboard 的 B 欄（股票名稱）是否會自動拉寬以容納完整名稱，而不會被裁切或折行。
   * 檢查其餘數字欄位是否維持在設定的最小寬度以上，沒有因為 `GOOGLEFINANCE` 顯示載入中而縮成極小，亦無 `###` 的現象。
