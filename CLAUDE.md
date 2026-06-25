# 記帳系統 — 專案說明（給 Claude Code 的常駐脈絡）

## 這是什麼
個人記帳系統。後端是一份 Google 試算表，前端是部署成「網頁 App」的 Google Apps Script，
使用者在 iPhone 用「加入主畫面」當 App 開啟。**一個月一張分頁**。
程式碼以 `src/`(GAS 本體) ＋ `docs/`(GitHub Pages 啟動頁) 結構放在 git repo，是**自架範本**：
每個人各自部署自己的後台（自己的試算表＋Apps Script），資料互不相通；可共用同一個啟動頁（只決定外框與圖示）。

## 檔案
- `src/Code.js` — Apps Script 後端：`setup()`、`getConfig()`、`addTransaction()`、`getMonthData()`、
  `getStats()`、`getBalances()`、`deleteRow()`、`updateRow()`、`getCardConfig()`、`setCardConfig()`、`getCardStatement()`、`recordCardPayment()`，以及試算表自動排版（`styleConfigSheet_` / `styleMonthSheet_`）。
  - `getMonthData(month)`：某月交易明細（記帳頁「本月最近」用）。
  - `getStats(gran, key)`：依 `day`/`week`/`month`/`year` 聚合收入、支出、結餘與該期間明細（統計頁用），回傳 `label`（週會給「起訖日」字串）。分布圓餅改由前端從 `txns` 即時算（可依 支出/收入 × 帳戶 篩選），後端不再回傳 pie。
  - `weekRange_(dateStr)`：以**週一為一週開始**；給週內任一天，回傳該週起訖日與涵蓋的月份分頁（跨月的週會涵蓋兩張）。
  - `getBalances()`：各帳戶餘額＋淨資產（帳戶頁用）；底層仍是 `getBalances_()`。信用卡額外附 `closeDay/dueDay/currentDue(本期未繳)/pending(上期已結帳未繳)`。
  - `updateRow(oldMonth, row, t)`：編輯已存紀錄；沿用原「建立時間」；**若日期被改到別的月份會自動把該筆搬到正確月份分頁**。
  - `readTxns_(sh)`：讀單張月份分頁的交易（getMonthData/getStats 共用）；每筆帶 `sheet` 欄，讓前端能用 `sheet`+`row` 定位編輯/刪除（年/週檢視會跨多張分頁）。
  - **信用卡帳單週期**：`getCardConfig()`/`setCardConfig(card,close,due)` 讀寫每張卡的結帳日/繳款日（存在設定分頁 I~K 欄，見資料模型）；`getCardStatement(card, ym)` 回傳某結帳月的帳單（依**真實結帳日→結帳日**週期：明細、合計、我/家分攤、繳款日、`paid`/`isClosed`）；`recordCardPayment(card, ym, mineAcct, famAcct, mineAmtIn, famAmtIn)` 依付款方把該期拆成「我」「家」各一筆轉帳(來源→卡)，**金額以前端填的「銀行帳單實際金額」為準**(沒填才用估算)，note 標 `繳款·<結帳月>`(用月份，日期微調不影響) 供 `isCyclePaid_` 判定已繳。`reverseCardPayment(card, ym)` 撤銷某期繳款(刪掉該期繳款轉帳、還原帳戶)；`setChargeBill(sheet, row, billYm)` 設某筆刷卡的「帳單月(J欄)」做延期。週期計算：`cycleByCloseMonth_`/`openCloseMonth_`/`dueDateFor_`/`naturalCloseMonth_`/`addMonths_`/`cardCycleData_`(考慮延期、回傳 checked/deferred)/`cardCurrentDue_`/`cardPendingClosed_`。
  - **延期(帳單頁勾選)**：每筆刷卡有「有效結帳月」＝`帳單月`(J欄)覆蓋值，沒有就用消費日的自然結帳月(`naturalCloseMonth_`)。帳單頁取消勾選某筆＝把它的帳單月設成下一結帳月(延到下一期)；勾選＝設回本期。合計/分攤/本期已刷只算「計入本期」的。
  - **重要觀念**：App 用「消費日」歸期，銀行用「入帳日」，兩者必有落差，所以 App 的每期金額是**估算**（帳戶卡片顯示「本期已刷」）。真相來源是銀行帳單：繳款時填銀行實際金額即可，**與已記刷卡的差額會自然留在卡片累計餘額、滾到下一期**，不需追入帳日。淨資產靠「累計餘額＋每期照實繳」維持正確。
- `src/Index.html` — 前端 UI（HtmlService）。三個分頁：
  - **記帳**：交易表單 ＋「本月最近」短清單（點一筆 → 操作選單可編輯/刪除）。
  - **統計**：`日/週/月/年` ＋ 指定日期/週(選週內一天)/月份/年份 → 收入/支出/結餘、分布圓餅（支出/收入切換、可再依帳戶篩選）、該期間明細（可編輯/刪除）。
  - **帳戶**：淨資產 ＋ 各帳戶餘額（跨月累計）。信用卡顯示**本期未繳**（有設結帳日才算；否則退回累計未繳）＋結帳/繳款日＋上期待繳提醒；**點卡片 → 卡片帳單浮層**（切換帳單週期、明細、我/家分攤、設定結帳/繳款日、`記錄繳款`）。另有「本月卡費分攤」浮層（依付款方拆我/家）。
  前端用 `google.script.run` 呼叫後端；已移除頂部月份下拉，期間選擇只在統計頁。
- `src/appsscript.json` — Apps Script 設定檔，非必要勿動（`doGet()` 已加 `setXFrameOptionsMode(ALLOWALL)` 讓 App 可被啟動頁 iframe 嵌入）。
- `docs/` — GitHub Pages 啟動頁（**可選**；只有把 App 設成「任何人可用」時才用得上）：
  - `docs/index.html`：全螢幕 iframe 嵌入 App，後台一律由網址參數 `?app=<exec>` 指定（公開 repo 內**不寫死**任何人的後台網址）；只接受 `script.google.com/macros/.../exec`。
  - `docs/apple-touch-icon.png`：啟動頁的主畫面圖示（咬一口扁平甜甜圈、180×180）。原理：GAS 把 App 包在沙箱 iframe，App 內的 `apple-touch-icon` 傳不到最上層，所以靠這個「自己掌握最上層 `<head>`」的啟動頁。
  - **限制**：私有(「只有我自己」)的 App 在跨網域 iframe 會被第三方 Cookie 擋而回 **403**，所以啟動頁方案**只在 App 設為匿名可用時有效**。
- **最終選擇＝維持 App 私有 ＋ 直接 PWA**：用 Safari 開 `/exec` → 加入主畫面，得到全螢幕、可用、私有的 App；主畫面圖示是「記」字。
  - 這是避不掉的**三選二**：私有 GAS App 在「自訂圖示／全螢幕 standalone／私有」最多同時取兩個。圖示要自訂得掌握最上層 `<head>`，但 GAS 私有 App 最上層由 Google 控制（只給設標題＝「記」），塞不進 `apple-touch-icon`。
  - **iOS 捷徑法試過、不採用**：會強制進瀏覽器、失去全螢幕（且在預設瀏覽器帳號不對時會撞 Google 權限錯誤）。
  - `docs/` 啟動頁＋`doGet` 的 `ALLOWALL` 是「**日後若把 App 改成匿名可用**」才用得到的現成選項（那時可同時拿到甜甜圈圖示＋全螢幕）；目前對私有 App 停用，但保留不刪。
- `assets/` 圖示原始檔；`README.md` 安裝/更新說明；`update.sh` 朋友的一鍵更新（`git pull && clasp push && clasp deploy`，自動抓自己的部署 ID）。

## 資料模型
每筆交易欄位：日期時間、類型(收入/支出/轉帳)、出帳帳戶、入帳帳戶、金額、類別、備註、建立時間、家裡負擔%。
- 收入：入帳帳戶＝資產帳戶；出帳留空。
- 支出：出帳帳戶＝資產或信用卡；入帳留空。
- 轉帳：出帳 → 入帳。
- **日期時間**(A欄)：記錄到「分」(前端 `datetime-local`，預設帶當下時刻)，用來分析消費習慣；月份分頁仍依此欄的 `yyyy-MM` 歸屬。改版前的舊資料無時間，顯示為中午 12:00。`parseDate_` 同時吃純日期與含時間字串。
- **建立時間**(H欄)：系統記錄當下的時間戳，編輯時不會變動，與「日期時間」分開。
- **家裡負擔%**(I欄)：付款方，只對「刷卡支出」有意義（0=我全出、100=家裡全出、其它=拆帳的家裡比例）；空白=不適用/我全出。
- **帳單月**(J欄)：信用卡延期標記，`yyyy-MM`=把這筆延到該結帳月、空白=照消費日歸期。後端用 `CREATED_COL=8`、`BILL_COL=10` 固定抓欄位（加欄後不可用 `TXN_HEADERS.length`）。

帳戶與類別都定義在試算表「設定」分頁（帳戶在 A 欄起、類別在 F 欄起，**資料皆從第 6 列開始**）。
帳戶類型欄：`資產` / `負債`(信用卡) / `外部`。
**信用卡帳單設定**放在「設定」分頁的 **I~K 欄**（卡名/結帳日/繳款日），刻意避開 getConfig 讀取的 A~D、F~G 範圍；首次呼叫 `getCardConfig` 會由 `ensureCardSection_` 自動建好標題與各負債卡列（日先空白），也可在 App 卡片帳單頁內設定。

## 改程式時必須遵守的規則
- **信用卡走「帳戶版」**：刷卡＝支出且出帳帳戶＝該卡（卡餘額為負＝未繳）；繳費＝轉帳（銀行→卡）；
  華南卡家人代繳＝轉帳（外部／他人 → 華南卡）。
- 餘額是**跨所有月份分頁累計**（含設定分頁的期初餘額）；`外部` 類型不列入淨資產。
- 月份分頁名稱格式固定為 `YYYY-MM`，`getBalances_` 靠這個 regex 篩選分頁，**不要改這個格式**。
- 這是 GAS HtmlService 環境，前端**不要用 localStorage / sessionStorage**（也用不到）。
- UI 維持現有風格：乾淨、中性配色、好按的手機介面，沿用既有 CSS 變數，不要大改視覺方向。

## 標準工作流程（每次改完都要做）——**兩條線都要推**
1. 改 `src/Code.js` / `src/Index.html`（或 `docs/`）。
2. **推到 Apps Script（Google）**：`clasp push` →
   `clasp deploy -i AKfycbwn-cmg2Lv8E80aXT6fIkzeN1BX-uVUsaEMRBMKQTMoCM30CXutZkiFDWRzorPfYAh-jA`
   （rootDir 已設 `src`，clasp 只推 src/ 三檔；這個是版本化 Web App 部署 ID，另有 `@HEAD` 是測試用，`clasp deployments` 可查。）
3. **推到 GitHub**：`git push`（GitHub Pages 服務 `docs/` 啟動頁；改 `docs/` 後所有使用者自動更新）。
- 只動 `docs/`（啟動頁/圖示）→ 只做第 3 步即可。動到 `src/`（App 本體）→ 第 2、3 步都要做。

## 注意
- `clasp push` 會用本機覆蓋雲端、`clasp pull` 會用雲端覆蓋本機。若曾在 Apps Script 網頁編輯器
  臨時改過，先 `clasp pull` 再動本機，以免蓋掉雲端的改動。
- routine 小改用 Sonnet 即可，不用 Opus。
- 專案擁有者帳號＝ hoffforliving@gmail.com；clasp 已用此帳號登入。
