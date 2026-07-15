# 實作計畫：修復側邊欄新增交易後 Dashboard 未自動同步問題

## 目的與問題描述
目前使用者透過側邊欄（Sidebar）新增交易紀錄時，資料確實會寫入「個股交易紀錄」分頁，但「庫存總覽 Dashboard」與「年度收益總覽」並不會同步更新，使用者必須手動點選上方選單進行更新。本計畫旨在修改 `addTransaction` 函式，使其在成功寫入交易紀錄後，自動且靜默地（不強制切換焦點分頁）重新建立/更新 Dashboard 與年度收益總覽，維持資料的即時同步。

---

## 使用者審查項目 (User Review Required)
> [!IMPORTANT]
> **靜默更新設計**：
> 在自動同步時，更新函式將以 `focus = false` 呼叫（如 `rebuildDashboard_(ss, false)`）。這能確保當使用者在填寫側邊欄時，試算表的當前選取分頁不會被強制切換，避免干擾使用者操作。

---

## 預計修改內容

### [Backend Logic]
修改 `code.gs.js` 中的 `addTransaction` 函式，在成功寫入資料並設定公式後，調用 `rebuildDashboard_` 與 `rebuildYearly_`，並調整返回給前端 Sidebar 的提示訊息。

#### [MODIFY] [code.gs.js](file:///mnt/c/Users/user/Desktop/股票/code.gs.js)

**修改前 (約第 358 - 363 行)：**
```javascript
  sheet.getRange(nextRow, 8).setFormula(
    `=IF(D${nextRow}="買入", -(E${nextRow}*F${nextRow}+G${nextRow}), (E${nextRow}*F${nextRow}-G${nextRow}))`
  );

  return "🎉 成功新增交易紀錄!(記得重建 Dashboard 以更新統計)";
}
```

**修改後：**
```javascript
  sheet.getRange(nextRow, 8).setFormula(
    `=IF(D${nextRow}="買入", -(E${nextRow}*F${nextRow}+G${nextRow}), (E${nextRow}*F${nextRow}-G${nextRow}))`
  );

  // 自動更新 Dashboard 與年度收益總覽 (靜默更新，不切換焦點分頁)
  try {
    rebuildDashboard_(ss, false);
  } catch (e) {
    Logger.log("自動重建 Dashboard 失敗: " + e);
  }
  try {
    rebuildYearly_(ss, false);
  } catch (e) {
    Logger.log("自動重建年度收益總覽失敗: " + e);
  }

  return "🎉 成功新增交易紀錄，且已自動更新 Dashboard 與年度收益總覽！";
}
```

---

## 驗證計畫 (Verification Plan)

### 手動驗證流程
由於本專案為 Google Apps Script (GAS)，無法直接在本地終端運行單元測試，建議在 Google 試算表專案中更新此代碼後，進行以下手動驗證：
1. 開啟試算表，點擊自訂選單 **`⚡️ 股票小幫手`** -> **`💰 買賣輸入`** 開啟側邊欄。
2. 停留在「個股交易紀錄」或其他分頁（非 Dashboard 分頁）。
3. 填入一筆測試交易（例如：台股 `2330` 買入 1000 股），點擊確認送出。
4. 觀察側邊欄提示是否顯示為 `"🎉 成功新增交易紀錄，且已自動更新 Dashboard 與年度收益總覽！"`。
5. 切換到「庫存總覽 Dashboard」與「年度收益總覽」，確認該筆新交易的數據已自動被計入，且先前未強制切換分頁。
