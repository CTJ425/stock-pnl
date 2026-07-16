# 實作結果說明 (Walkthrough)

此專案已完成台股未實現損益計算精準度（四捨五入 + 無條件捨去）的對齊，以及台美股數字格式的自動分區段格式化設定。異動內容已通過程式碼審查 (Code Review)，並已成功推送至 **`dev`** 分支。

---

## 變更項目 (Changes Made)

### 1. 核心計算公式優化 (code.gs)
在 [code.gs](file:///mnt/c/Users/user/Documents/Configuration/stock-pnl/code/code.gs) 的 `writeDashboard_` 函式中：
* 針對台股交易列，將預估的賣出手續費與證交稅修改為分別以 `FLOOR()` 函數進行無條件捨去（與券商 App 及 Sidebar `calculateFee` 邏輯同構）。
* 於台股損益的最外層套用 `ROUND(..., 0)` 函數，將移動平均成本中因除權息或分割造成的浮點尾數四捨五入收整。
* 美股（USD）維持原本無估算稅費的計算公式。

### 2. 多幣別連續區塊數字格式設定 (code.gs)
移除原先整欄套用小數點的設定，改為先計算台股數量 `twCount`，並依據 `holdings` 的排序特性（台股在前、美股在後）分區段套用格式：
* **台股區塊 (TWD)**：
  * 現價（Column C）與平均成本（Column E）保留兩位小數（如 `NT$143.50`，避免誤導）。
  * 目前市值（Column F）與損益（Column G 至 I）設定為不帶小數的整數格式（如 `NT$7,293`）。
* **美股區塊 (USD)**：
  * 現價、平均成本、市值與損益全部保留兩位小數點並冠以 `US$` 前綴。

---

## 驗證結果與推送紀錄 (Validation & Git History)

1. **程式碼審查 (Self Code Review)**：
   * 確保所有公式的括號皆配對正確，無 ReferenceError，且原先被刪除的現價拉取 `fPrice` 與市值計算 `fMktVal` 公式已全數還原到位。
2. **Git 版本控制**：
   * 成功切換並同步至 `dev` 分支。
   * 提交變更並順利推送到遠端儲存庫的 `dev` 分支（`dev -> dev`），並**未**動到 `main` 分支。

```bash
[dev 2dbd187] feat: implement PnL rounding and currency-specific block formatting for TWD/USD
 2 files changed, 137 insertions(+), 7 deletions(-)
 create mode 100644 docs/plan_20260716_v2.md
```
