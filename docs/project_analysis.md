# 專案功能與架構分析：股票交易與庫存管理系統

此專案是一個基於 **Google Apps Script (GAS)** 的股票交易紀錄與庫存管理系統，專為 Google 試算表（Google Sheets）設計。它主要協助使用者追蹤台股（TWD）及美股（USD）的交易紀錄，自動計算持股成本、市值、未實現損益以及年度已實現損益，並提供友好的側邊欄交易輸入介面與聯動股票搜尋功能。

---

## 系統架構與檔案關聯

本專案由三個主要檔案組成：

1. **[code.gs.js](file:///mnt/c/Users/user/Desktop/股票/code.gs.js)**: 主程式邏輯與工具函式庫。
2. **[YearReport.gs.js](file:///mnt/c/Users/user/Desktop/股票/YearReport.gs.js)**: 年度收益總覽報表的獨立模組。
3. **[Sidebar.html](file:///mnt/c/Users/user/Desktop/股票/Sidebar.html)**: 交易輸入側邊欄的 HTML UI。

其相互呼叫與關聯架構如下圖所示：

```mermaid
graph TD
    subgraph Google Sheet UI
        Menu["⚡️ 股票小幫手 (自訂選單)"]
    end

    subgraph Sidebar.html (側邊欄 UI)
        Form["交易輸入表單 (日期/市場/代號/單價/股數/手續費)"]
        SearchUI["模糊搜尋與雙向反查 UI"]
    end

    subgraph code.gs (主程式)
        onOpen["onOpen() <br> 建立選單"]
        showSidebar["showSidebar() <br> 顯示側邊欄"]
        addTx["addTransaction() <br> 寫入交易紀錄"]
        compLedger["computeLedger_() <br> 核心引擎 (移動平均成本法)"]
        dbDashboard["createPortfolioDashboard() <br> 建立庫存總覽"]
        searchModule["統一搜尋模組 (unifiedSearch/searchByTicker)"]
        caching["快取機制 (CacheService)"]
    end

    subgraph YearReport.gs (年度收益)
        yearlyReport["createYearlyReport() <br> 建立年度收益總覽"]
    end

    subgraph External APIs (外部 API)
        TWSE["證交所 codeQuery API (上市)"]
        TPEx["櫃買中心 OpenAPI (上櫃)"]
        Yahoo["Yahoo Finance API (美股/全球)"]
        GF["GOOGLEFINANCE 函式 (即時現價)"]
    end

    %% 關聯線
    Menu -->|點擊買賣輸入| showSidebar
    showSidebar -->|載入| Form
    Menu -->|點擊建立Dashboard| dbDashboard
    Menu -->|點擊建立年度收益| yearlyReport

    Form -->|送出表單| addTx
    addTx -->|寫入分頁| RecordSheet[("個股交易紀錄 分頁")]
    
    SearchUI -->|反查代號/名稱| searchModule
    searchModule -->|查詢台股上市| TWSE
    searchModule -->|查詢台股上櫃| TPEx
    searchModule -->|查詢美股/全球| Yahoo
    searchModule -->|暫存結果| caching

    dbDashboard -->|讀取交易紀錄| compLedger
    dbDashboard -->|寫入靜態值與輕量公式| DashboardSheet[("庫存總覽 Dashboard 分頁")]
    DashboardSheet -->|現價查詢| GF

    yearlyReport -->|讀取交易紀錄| compLedger
    yearlyReport -->|寫入已實現統計| YearlySheet[("年度收益總覽 分頁")]
```

---

## 核心功能說明

### 1. 核心計算引擎 (移動平均成本法)
* 系統統一使用 `computeLedger_(recordSheet)` 進行帳簿計算。
* 依交易日期與輸入順序進行排序，採用**移動平均成本法 (Moving Average Cost Method)** 計算持有部位與成本。
* 當進行「賣出」時，會根據目前的平均成本計算**已實現損益**，並扣減持有股數。
* **資料異常警告機制**：若發生「超賣」（賣出股數大於當時持有股數），系統會記錄警告，並以持有股數為上限進行計算（超賣部分成本以 $0 估算），避免計算溢出。

### 2. 庫存總覽 Dashboard
* **台美股獨立統計**：為避免幣別混算造成混亂，摘要卡（KPI Cards）將台股（TWD）與美股（USD）分開統計。
* **效能優化 (穩定版)**：由於 Google 試算表的大型陣列公式（如 `LET+REDUCE+LAMBDA`）容易導致服務不穩定或載入緩慢，此系統改採「**記憶體預算 + 輕量公式**」策略：
  * 持有股數、均價、已實現損益等靜態數據，由 Apps Script 直接計算後寫入。
  * 目前現價則動態寫入輕量的 `=IFERROR(GOOGLEFINANCE(ticker,"price"),0)` 公式，以保持現價即時更新。
  * 未實現損益與報酬率透過簡單的試算表公式（如 `市值 = 現價 * 股數`）在儲存格間運算，保證高流暢度。
* **無條件式格式**：因 Google 試算表對條件式格式服務層有時不穩定，系統不使用試算表內建的條件格式規則，而是直接將紅字/綠字顏色編碼寫入儲存格的**數字格式**（Number Format）中，如 `[Red]$#,##0.00;[Green]-$#,##0.00;$0.00`，實現更穩定的「紅漲綠跌」效果。

### 3. 年度收益總覽
* 透過 `createYearlyReport()` 自動彙整歷年已實現的收益。
* 列出每一年度的：
  1. 台股已實現損益 (TWD)
  2. 美股已實現損益 (USD)
  3. 買入總額 & 賣出總額（混合幣別）
  4. 手續費合計
  5. 交易筆數
* 底部附帶自動計算的「合計」列，方便使用者進行年度報稅或長期資產評估。

### 4. 股票搜尋模組 & 側邊欄聯動
* 側邊欄提供極佳的雙向輸入聯動體驗：
  * **代號反查名稱**：輸入股票代號（如 `2330` 或 `AAPL`）並移開焦點後，自動向後端發送請求，取得對應中文名稱（如 `台積電`）與市場（TPE/美股）。
  * **名稱模糊搜尋**：輸入中文名稱時，會觸發防抖（Debounce, 300ms）並自動列出匹配的股票下拉清單（台美股混合搜尋），點選後自動填入代號、市場與名稱。
* **外部 API 整合與快取**：
  * **台股上市**：呼叫證交所 `codeQuery` API。
  * **台股上櫃**：呼叫櫃買中心 OpenAPI 取得全清單，並利用 `CacheService` 進行本地比對，快取時間為 6 小時，避免頻繁請求。
  * **美股**：呼叫 Yahoo Finance Search API。

---

## 接下來的計畫 / 步驟

由於您發送了 `/plan` 指令，且目前的目的是「先瞭解專案是用來幹嘛的」，目前系統已處於可運作狀態。如果您有進一步的開發需求，以下是一些常見的優化或擴充方向供您參考：

### 潛在擴充方案
1. **匯率自動折算**：目前台股（TWD）與美股（USD）是完全拆開計算與顯示的。若希望有「總資產折合台幣」功能，可新增一欄或摘要卡，引入外匯報價進行折算。
2. **股利與配息追蹤**：目前交易類型僅支援「買入」與「賣出」。未來可擴充「配股」、「配息」等交易類型，自動扣減持股成本或列入已實現收益。
3. **歷史圖表分析**：新增資產配置圓餅圖（Sector/Stock Asset Allocation）或歷史資產淨值走勢圖。
4. **手續費折扣折讓計算**：針對台股的不同券商折讓（如 28 折、6 折）提供設定選項，使成本計算更貼近實際交割款。

---

> [!NOTE]
> 本分析已建立完成。請您確認此專案的用途與您的預期是否相符。
> 若您有具體的功能修改、Bug 修復或新功能開發需求，請告訴我，我將為您編寫具體的**實作計畫書（Implementation Plan）**供您審查！
