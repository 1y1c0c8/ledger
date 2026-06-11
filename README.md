# 記帳 — 個人記帳系統（Google Apps Script，自架範本）

放在 Google Drive 的個人記帳 App：後端是一份 Google 試算表，前端用 Apps Script 部署成網頁 App，
iPhone 用「加到主畫面」當 App 開。**一個月一張分頁**。

> **這是「自架範本」，不是共用服務。** 每個人各自部署自己的一份（自己的試算表＋自己的 Apps Script），
> 資料互不相通。大家可以共用同一個「啟動頁」（只決定圖示與外框），但後台各連各的。

---

## 目錄結構
```
src/    Apps Script 本體（clasp 管理）：Code.js、Index.html、appsscript.json
docs/   GitHub Pages 啟動頁：index.html（全螢幕 iframe 嵌入你的 App）＋ apple-touch-icon.png（主畫面圖示）
assets/ 圖示原始檔（icon-180/512.png）
update.sh  一鍵更新腳本
```

---

## A. 第一次安裝（每個人各做一次）

1. **裝工具**：`npm i -g @google/clasp`，然後 `clasp login`（登入你自己的 Google 帳號）。
2. **建立你自己的後台**：在 Apps Script 新建一個專案（或用 `clasp create --type standalone`），
   複製 `.clasp.json.example` 成 `.clasp.json`，把裡面的 `scriptId` 換成你自己的（`rootDir` 維持 `src`）。
3. **推程式並初始化**：`clasp push` →在 Apps Script 編輯器手動執行一次 `setup()`
   （會建立「設定」分頁與當月分頁；帳戶/類別直接在「設定」分頁增修）。
4. **部署網頁 App**：Apps Script → 部署 → 新增部署作業 → 類型「網頁應用程式」，
   執行身分「**我**」、存取權限視需要（見下方注意）。記下你的 **`/exec` 網址**。
5. **加到主畫面**：用 iPhone Safari 開
   `https://<維護者帳號>.github.io/<repo>/?app=<你的 /exec 網址>`
   → 分享 → 加入主畫面。圖示就是甜甜圈，內容是你自己的後台。

> 不想依賴別人的啟動頁？把這個 repo fork 一份、開自己的 GitHub Pages（Settings → Pages → 來源選 `main` 的 `/docs`），
> 用 `https://<你的帳號>.github.io/<repo>/?app=<你的exec>` 即可，完全自主。

---

## B. 更新（不用重裝、資料不動）

程式有新版時，在你本機 repo 跑：
```bash
./update.sh        # = git pull + clasp push + clasp deploy（自動沿用你原本的部署，/exec 不變）
```
`setup()` 與你的試算表資料都不會動，只是換成新程式。iPhone 上的 App 重新整理就是新版。

- 啟動頁/圖示（`docs/`）的更新，由維護者推到 GitHub Pages 後**所有人自動拿到**，免動作。
- App 本體（`src/`）是各自的 Apps Script，需各自跑一次 `./update.sh`。

---

## C. 維護者（repo 擁有者）改完程式的標準流程 —— 兩邊都要推

| 改了什麼 | 指令 | 更新到 |
|---|---|---|
| `src/`（App 本體） | `clasp push` → `clasp deploy -i <你的部署ID>` | 你的 Apps Script（Google） |
| 任何檔（含 `src/`、`docs/`、README） | `git push` | GitHub（Pages 服務 `docs/`） |

> 也就是：**動到 App 本體要 clasp＋git 兩邊推**；只動啟動頁/圖示則 git 一邊即可。

---

## 注意
- **為什麼要啟動頁**：GAS 把網頁包在沙箱 iframe，iOS 主畫面讀不到 App 內的 `apple-touch-icon`，
  所以圖示一定要靠這個「自己掌握最上層 `<head>`」的 GitHub Pages 啟動頁。`doGet()` 已加
  `setXFrameOptionsMode(ALLOWALL)` 允許被嵌入。
- **私有存取 × iframe**：App 若部署成「只有我自己」，在 github.io 的跨網域 iframe 中，iOS 可能因擋第三方 Cookie 卡登入。
  若中招：改成「任何人」存取（資料仍在你私人試算表、網址夠亂），或改用 iOS 捷徑法。
- **別提交憑證**：`.clasp.json`（含 scriptId）與 `~/.clasprc.json` 已被 `.gitignore` 排除。
