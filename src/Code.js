/*************************************************************
 * 記帳系統 — Google Apps Script 後端 (Code.gs)
 * 一個月一張分頁、信用卡走「帳戶版」(記欠款、繳費當轉帳)
 * 後台試算表會自動套用易讀排版(表頭、斑馬紋、顏色標示、下拉選單)
 *************************************************************/

const TZ = 'Asia/Taipei';
const CFG_SHEET = '設定';
const TXN_HEADERS = ['日期', '類型', '出帳帳戶', '入帳帳戶', '金額', '類別', '備註', '建立時間', '家裡負擔%'];
const CREATED_COL = 8;   // 「建立時間」固定在第 8 欄（不可用 TXN_HEADERS.length，加欄後會跑掉）

const COLORS = {
  header:'#0f766e', headerText:'#ffffff', section:'#e3efed', sectionText:'#0b5c55',
  expenseBg:'#fdecea', expenseFg:'#b23a23',
  incomeBg:'#e9f4ee',  incomeFg:'#2f7d4f',
  transferBg:'#eff0f2',transferFg:'#5b6066',
  line:'#e7e5de', muted:'#8a877f'
};

// 第一次跑 setup() 會把以下清單寫進「設定」分頁。
// 之後要新增/修改帳戶或類別，直接改「設定」分頁就好，不用動程式。
//   類型：資產 / 負債(信用卡) / 外部(不列入淨值，例如家人代繳的來源)
const DEFAULT_ACCOUNTS = [
  // [名稱, 分類, 類型, 期初餘額]
  ['郵局',         '帳戶',     '資產', 0],
  ['台新',         '帳戶',     '資產', 0],
  ['錢包',         '現金',     '資產', 0],
  ['學生證',       '現金儲值', '資產', 0],
  ['LinePay',      '現金儲值', '資產', 0],
  ['iPassMoney',   '現金儲值', '資產', 0],
  ['國泰Cube卡',   '信用卡',   '負債', 0],
  ['華南超鑽卡',   '信用卡',   '負債', 0],
  ['永豐證券戶',   '投資',     '資產', 0],
  ['零錢罐',       '其他',     '資產', 0],
  ['外部／他人',   '外部',     '外部', 0],   // 家人代繳卡費的來源；不列入你的淨值
];

const DEFAULT_CATEGORIES = [
  // [名稱, 收支類型]
  ['飲食',       '支出'],
  ['交通',       '支出'],
  ['日用',       '支出'],
  ['娛樂',       '支出'],
  ['學習',       '支出'],
  ['醫療',       '支出'],
  ['通訊',       '支出'],
  ['儲值',       '支出'],
  ['零錢',       '支出'],
  ['其他支出',   '支出'],
  ['薪資',       '收入'],
  ['獎助學金',   '收入'],
  ['利息／配息', '收入'],
  ['退款',       '收入'],
  ['其他收入',   '收入'],
];

/** 部署成網頁 App 後的進入點 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('記帳')
    .addMetaTag('viewport',
      'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)   // 允許被 GitHub Pages 啟動頁的 iframe 嵌入（自訂主畫面圖示用）
    .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/favicon3.ico');
}

/** 第一次安裝：建立並美化「設定」分頁與當月分頁。在編輯器裡手動執行一次。 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfg = ss.getSheetByName(CFG_SHEET);
  if (!cfg) cfg = ss.insertSheet(CFG_SHEET, 0);
  styleConfigSheet_(cfg);
  getOrCreateMonthSheet_(currentMonth_());
  ss.setActiveSheet(cfg);
  ss.toast('初始化完成 ✓ 接著做「部署 → 新增部署作業 → 網頁應用程式」', '記帳系統', 8);
}

/** 美化「設定」分頁並填入預設帳戶/類別 */
function styleConfigSheet_(cfg) {
  cfg.clear();
  cfg.getBandings().forEach(b => b.remove());
  cfg.setConditionalFormatRules([]);
  cfg.setHiddenGridlines(true);

  // 標題列
  cfg.getRange('A1:H1').merge().setValue('⚙  設定 — 在這裡管理帳戶與類別')
     .setBackground(COLORS.header).setFontColor(COLORS.headerText)
     .setFontSize(13).setFontWeight('bold').setVerticalAlignment('middle');
  cfg.setRowHeight(1, 36);
  cfg.getRange('A2:H2').merge()
     .setValue('新增帳戶或類別：在表格最後一列往下接著打就好，App 下次打開即生效。'
             + '「類型」「收支類型」欄請點儲存格用下拉選單。類型＝資產／負債(信用卡)／外部(家人代繳來源，不計淨值)。')
     .setFontColor(COLORS.muted).setFontSize(10).setWrap(true).setVerticalAlignment('middle');
  cfg.setRowHeight(2, 34);

  // 區塊標題
  cfg.getRange('A4').setValue('帳戶').setFontWeight('bold').setFontColor(COLORS.sectionText).setFontSize(12);
  cfg.getRange('F4').setValue('類別').setFontWeight('bold').setFontColor(COLORS.sectionText).setFontSize(12);

  // 帳戶表
  cfg.getRange(5, 1, 1, 4).setValues([['名稱', '分類', '類型', '期初餘額']])
     .setBackground(COLORS.section).setFontWeight('bold').setFontColor(COLORS.sectionText);
  cfg.getRange(6, 1, DEFAULT_ACCOUNTS.length, 4).setValues(DEFAULT_ACCOUNTS);

  // 類別表
  cfg.getRange(5, 6, 1, 2).setValues([['名稱', '收支類型']])
     .setBackground(COLORS.section).setFontWeight('bold').setFontColor(COLORS.sectionText);
  cfg.getRange(6, 6, DEFAULT_CATEGORIES.length, 2).setValues(DEFAULT_CATEGORIES);

  // 欄寬
  cfg.setColumnWidth(1, 120); cfg.setColumnWidth(2, 90);
  cfg.setColumnWidth(3, 70);  cfg.setColumnWidth(4, 100);
  cfg.setColumnWidth(5, 28);
  cfg.setColumnWidth(6, 120); cfg.setColumnWidth(7, 90);

  // 期初餘額格式
  cfg.getRange('D6:D200').setNumberFormat('$#,##0');

  // 下拉驗證
  const vType = SpreadsheetApp.newDataValidation()
    .requireValueInList(['資產', '負債', '外部'], true).setAllowInvalid(false).build();
  cfg.getRange('C6:C200').setDataValidation(vType);
  const vKind = SpreadsheetApp.newDataValidation()
    .requireValueInList(['收入', '支出'], true).setAllowInvalid(false).build();
  cfg.getRange('G6:G200').setDataValidation(vKind);

  // 邊框
  cfg.getRange(5, 1, DEFAULT_ACCOUNTS.length + 1, 4)
     .setBorder(true, true, true, true, true, true, COLORS.line, SpreadsheetApp.BorderStyle.SOLID);
  cfg.getRange(5, 6, DEFAULT_CATEGORIES.length + 1, 2)
     .setBorder(true, true, true, true, true, true, COLORS.line, SpreadsheetApp.BorderStyle.SOLID);

  cfg.getRange('A1:H1').setFontFamily('Noto Sans TC');
  cfg.getRange(4, 1, DEFAULT_CATEGORIES.length + 3, 7).setFontFamily('Noto Sans TC');
  cfg.setFrozenRows(2);
}

function currentMonth_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM');
}

function parseDate_(s) {
  if (!s) return new Date();
  // 接受 'yyyy-MM-dd'(舊資料/純日期) 或 'yyyy-MM-ddTHH:mm'(datetime-local)
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ](\d{2}):(\d{2}))?/);
  if (!m) return new Date();
  const hasTime = m[4] !== undefined;
  // 沒給時間就用當地中午，避免時區把日期推到前一天；有給就用實際時間
  const hh = hasTime ? Number(m[4]) : 12;
  const mi = hasTime ? Number(m[5]) : 0;
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), hh, mi, 0);
}

function getOrCreateMonthSheet_(month) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(month);
  if (!sh) {
    sh = ss.insertSheet(month);
    styleMonthSheet_(sh);
  }
  return sh;
}

/** 美化某張月份分頁 */
function styleMonthSheet_(sh) {
  sh.setHiddenGridlines(true);
  sh.getBandings().forEach(b => b.remove());
  sh.setConditionalFormatRules([]);

  // 表頭
  sh.getRange(1, 1, 1, TXN_HEADERS.length).setValues([TXN_HEADERS])
    .setBackground(COLORS.header).setFontColor(COLORS.headerText)
    .setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);

  // 欄寬
  const W = [130, 64, 110, 110, 100, 90, 220, 140, 80];
  W.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // 格式
  sh.getRange('A2:A').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.getRange('E2:E').setNumberFormat('$#,##0').setHorizontalAlignment('right');
  sh.getRange('H2:H').setNumberFormat('yyyy-mm-dd hh:mm').setFontColor(COLORS.muted);
  sh.getRange('B2:B').setHorizontalAlignment('center');

  // 斑馬紋（不蓋到表頭）
  sh.getRange('A2:I2000').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

  // 用顏色標示交易類型
  const rule = (t, bg, fg) => SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo(t).setBackground(bg).setFontColor(fg).setBold(true)
    .setRanges([sh.getRange('B2:B2000')]).build();
  sh.setConditionalFormatRules([
    rule('支出', COLORS.expenseBg, COLORS.expenseFg),
    rule('收入', COLORS.incomeBg, COLORS.incomeFg),
    rule('轉帳', COLORS.transferBg, COLORS.transferFg)
  ]);

  sh.getRange('A1:I2000').setFontFamily('Noto Sans TC');
}

/** 給前端：帳戶、類別、可用月份 */
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cfg = ss.getSheetByName(CFG_SHEET);
  if (!cfg) { setup(); cfg = ss.getSheetByName(CFG_SHEET); }
  const n = Math.max(cfg.getLastRow() - 5, 1);

  const accounts = cfg.getRange(6, 1, n, 4).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({ name: String(r[0]).trim(), group: r[1], type: String(r[2]).trim(), opening: Number(r[3]) || 0 }));

  const categories = cfg.getRange(6, 6, n, 2).getValues()
    .filter(r => r[0] !== '')
    .map(r => ({ name: String(r[0]).trim(), kind: String(r[1]).trim() }));

  const monthRe = /^\d{4}-\d{2}$/;
  let months = ss.getSheets().map(s => s.getName()).filter(nm => monthRe.test(nm));
  const cur = currentMonth_();
  if (months.indexOf(cur) === -1) months.push(cur);
  months.sort().reverse();

  return { accounts, categories, months, currentMonth: cur };
}

/** 新增一筆交易 */
function addTransaction(t) {
  const amt = Number(t.amount);
  if (!t.type) throw new Error('缺少交易類型');
  if (!(amt > 0)) throw new Error('金額需大於 0');

  const dateObj = parseDate_(t.date);
  const month = Utilities.formatDate(dateObj, TZ, 'yyyy-MM');
  const sh = getOrCreateMonthSheet_(month);
  // 確保日期欄顯示到分鐘（含早於本次改版就存在的舊分頁，會在下次寫入時自動補上）
  sh.getRange('A2:A').setNumberFormat('yyyy-mm-dd hh:mm');
  sh.setColumnWidth(1, 130);
  sh.appendRow([
    dateObj,
    t.type,
    t.accountOut || '',
    t.accountIn || '',
    amt,
    t.category || (t.type === '轉帳' ? '轉帳' : ''),
    t.note || '',
    new Date(),
    famPct_(t)               // 家裡負擔%（只對刷卡支出有意義，其餘留空）
  ]);
  return { ok: true, month: month };
}

/** 讀某張月份分頁的所有交易，整理成前端用的物件陣列（最新在前） */
function readTxns_(sh) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, TXN_HEADERS.length).getValues();
  return vals.map((r, i) => {
    const d = r[0] ? new Date(r[0]) : null;
    return {
      row: i + 2,
      sheet: sh.getName(),   // 編輯/刪除要知道這筆在哪張分頁（年檢視會跨多張）
      date: d ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : '',
      time: d ? Utilities.formatDate(d, TZ, 'HH:mm') : '',                 // 顯示用
      dt:   d ? Utilities.formatDate(d, TZ, "yyyy-MM-dd'T'HH:mm") : '',    // 編輯帶入 datetime-local 用
      type: r[1],
      accountOut: r[2],
      accountIn: r[3],
      amount: Number(r[4]) || 0,
      category: r[5],
      note: r[6],
      famPct: (r[8] === '' || r[8] == null) ? null : Number(r[8])   // 家裡負擔%（null=未設/我全出）
    };
  }).filter(x => x.type).reverse();
}

/** 把前端傳來的 famPct 正規化成存進試算表的值（空字串=不適用） */
function famPct_(t) {
  return (t.famPct === '' || t.famPct == null) ? '' : Number(t.famPct);
}

/** 記帳頁用：某月的交易明細（最新在前） */
function getMonthData(month) {
  month = month || currentMonth_();
  const sh = getOrCreateMonthSheet_(month);
  return { month: month, txns: readTxns_(sh) };
}

/** 統計頁用：依「日/月/年」聚合收入、支出、結餘、分布與明細
 *  gran：'day' | 'month' | 'year'
 *  key ：day→'yyyy-MM-dd'、month→'yyyy-MM'、year→'yyyy' */
function getStats(gran, key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const monthRe = /^\d{4}-\d{2}$/;
  let sheetNames, inRange, label = key;

  if (gran === 'year') {
    sheetNames = ss.getSheets().map(s => s.getName())
      .filter(n => monthRe.test(n) && n.slice(0, 4) === key);
    inRange = function () { return true; };
  } else if (gran === 'week') {
    const wr = weekRange_(key);                       // 以週一為一週開始
    sheetNames = wr.months;                            // 跨月的週會涵蓋兩張分頁
    inRange = function (t) { return t.date >= wr.start && t.date <= wr.end; };
    label = wr.start + ' ~ ' + wr.end;
  } else if (gran === 'day') {
    sheetNames = [key.slice(0, 7)];
    inRange = function (t) { return t.date === key; };
  } else { // month
    sheetNames = [key.slice(0, 7)];
    inRange = function () { return true; };
  }

  let income = 0, expense = 0;
  let txns = [];
  sheetNames.sort().forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    readTxns_(sh).forEach(t => {
      if (!inRange(t)) return;
      txns.push(t);
      if (t.type === '支出') expense += t.amount;
      else if (t.type === '收入') income += t.amount;
    });
  });
  // 跨多張分頁時重排，最新在前
  txns.sort((a, b) => (a.date + a.time < b.date + b.time ? 1 : -1));

  // 分布圓餅由前端依「支出/收入 × 選定帳戶」即時從 txns 計算，後端只給總額與明細
  return {
    gran: gran, key: key, label: label,
    income: income, expense: expense, net: income - expense,
    txns: txns
  };
}

/** 以「週一」為一週開始：給週內任一天 'yyyy-MM-dd'，回傳該週起訖日與涵蓋的月份分頁 */
function weekRange_(dateStr) {
  const m = String(dateStr).match(/^(\d{4})-(\d{2})-(\d{2})/);
  const d = m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0) : new Date();
  const dow = (d.getDay() + 6) % 7;                 // 週一=0 … 週日=6
  const start = new Date(d); start.setDate(d.getDate() - dow);
  const end = new Date(start); end.setDate(start.getDate() + 6);
  const fmtD = x => Utilities.formatDate(x, TZ, 'yyyy-MM-dd');
  const fmtM = x => Utilities.formatDate(x, TZ, 'yyyy-MM');
  const months = [fmtM(start)];
  if (fmtM(end) !== fmtM(start)) months.push(fmtM(end));
  return { start: fmtD(start), end: fmtD(end), months: months };
}

/** 跨所有月份累計各帳戶餘額（含期初） */
function getBalances_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const bal = {};
  cfg.accounts.forEach(a => bal[a.name] = a.opening);

  const monthRe = /^\d{4}-\d{2}$/;
  ss.getSheets().forEach(sh => {
    if (!monthRe.test(sh.getName())) return;
    const last = sh.getLastRow();
    if (last < 2) return;
    const vals = sh.getRange(2, 1, last - 1, TXN_HEADERS.length).getValues();
    vals.forEach(r => {
      const type = r[1], out = r[2], inn = r[3], amt = Number(r[4]) || 0;
      if (!type) return;
      if (type === '支出') { if (out in bal) bal[out] -= amt; }
      else if (type === '收入') { if (inn in bal) bal[inn] += amt; }
      else if (type === '轉帳') {
        if (out in bal) bal[out] -= amt;
        if (inn in bal) bal[inn] += amt;
      }
    });
  });

  return cfg.accounts
    .filter(a => a.type !== '外部')
    .map(a => ({ name: a.name, group: a.group, type: a.type, balance: bal[a.name] }));
}

/** 帳戶頁用：各帳戶餘額＋淨資產（外部不計入）；信用卡再附上帳單週期資訊 */
function getBalances() {
  const list = getBalances_();
  let networth = 0;
  list.forEach(b => networth += b.balance);
  const cfgMap = getCardConfig();
  list.forEach(b => {
    if (b.type !== '負債') return;
    const c = cfgMap[b.name] || {};
    b.closeDay = c.close || 0;
    b.dueDay = c.due || 0;
    if (c.close) {
      const cur = cardCurrentDue_(b.name, c.close);
      b.currentDue = cur.due;                     // 本期(未結帳)累積刷卡額
      b.nextClose = cur.closeDate;                // 本期預計結帳日
      const pend = cardPendingClosed_(b.name, c.close, c.due);
      b.pending = pend.amount;                    // 上一期(已結帳未繳)金額，0=無
      b.pendingDue = pend.dueDate;                // 上一期繳款日
    } else {
      b.currentDue = null;                        // 沒設結帳日 → 前端退回顯示累計未繳
      b.pending = 0; b.nextClose = ''; b.pendingDue = '';
    }
  });
  return { balances: list, networth: networth };
}

/* ===== 信用卡帳單週期 ===== */

const PAY_TAG = '繳款·';   // 繳款轉帳的 note 前綴，後接該期結帳日(yyyy-MM-dd)，用來判定「已繳」

function clampDay_(year, monthIdx, day) {
  const lastDay = new Date(year, monthIdx + 1, 0).getDate();
  return Math.min(day, lastDay);
}

/** 列出 startStr..endStr(yyyy-MM-dd) 涵蓋的月份分頁名(yyyy-MM) */
function monthsBetween_(startStr, endStr) {
  let [y, m] = startStr.slice(0, 7).split('-').map(Number);
  const [ey, em] = endStr.slice(0, 7).split('-').map(Number);
  const out = [];
  while (y < ey || (y === ey && m <= em)) {
    out.push(y + '-' + String(m).padStart(2, '0'));
    m++; if (m > 12) { m = 1; y++; }
  }
  return out;
}

/** 在 (year,monthIdx) 月以 closeDay 結帳的帳單週期 [start,end]（含端點）＋結帳日 */
function cycleByCloseMonth_(closeDay, year, monthIdx) {
  const endD = new Date(year, monthIdx, clampDay_(year, monthIdx, closeDay), 12, 0, 0);
  const ps = new Date(year, monthIdx - 1, clampDay_(year, monthIdx - 1, closeDay), 12, 0, 0);
  ps.setDate(ps.getDate() + 1);   // 上一期結帳日的隔天
  const f = x => Utilities.formatDate(x, TZ, 'yyyy-MM-dd');
  return { start: f(ps), end: f(endD), closeDate: f(endD) };
}

/** 以 ref 為準，目前「未結帳(open)」週期是哪個結帳月 {year, monthIdx} */
function openCloseMonth_(closeDay, ref) {
  const y = ref.getFullYear(), m = ref.getMonth(), d = ref.getDate();
  if (d <= clampDay_(y, m, closeDay)) return { year: y, monthIdx: m };
  const nx = new Date(y, m + 1, 1);
  return { year: nx.getFullYear(), monthIdx: nx.getMonth() };
}

/** 某期(結帳日 closeDate)的繳款日：dueDay 大於結帳日的「日」→同月，否則→下個月 */
function dueDateFor_(closeDate, dueDay) {
  const [y, m, d] = closeDate.split('-').map(Number);
  let yy = y, mm = m - 1;
  if (dueDay <= d) { const nx = new Date(yy, mm + 1, 1); yy = nx.getFullYear(); mm = nx.getMonth(); }
  return Utilities.formatDate(new Date(yy, mm, clampDay_(yy, mm, dueDay), 12, 0, 0), TZ, 'yyyy-MM-dd');
}

/** 該期(以結帳月 ym='yyyy-MM' 識別)是否已記錄繳款：找「轉入該卡且 note 帶 繳款·ym」的轉帳。
 *  用「月份」而非確切日期，這樣結帳日在 23/24 間微調也不會弄亂已繳判定。 */
function isCyclePaid_(card, ym) {
  if (!ym) return false;
  const tag = PAY_TAG + ym;
  const [y, m] = ym.split('-').map(Number);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (let k = 0; k <= 2; k++) {
    const dt = new Date(y, (m - 1) + k, 1);
    const sh = ss.getSheetByName(Utilities.formatDate(dt, TZ, 'yyyy-MM'));
    if (!sh) continue;
    const hit = readTxns_(sh).some(t => t.type === '轉帳' && t.accountIn === card && String(t.note || '').indexOf(tag) >= 0);
    if (hit) return true;
  }
  return false;
}

/** 某卡某週期內的刷卡總額 */
function sumCardCharges_(card, start, end) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sum = 0;
  monthsBetween_(start, end).forEach(mn => {
    const sh = ss.getSheetByName(mn); if (!sh) return;
    readTxns_(sh).forEach(t => {
      if (t.type === '支出' && t.accountOut === card && t.date >= start && t.date <= end) sum += t.amount;
    });
  });
  return sum;
}

/** 本期(未結帳)累積刷卡額 */
function cardCurrentDue_(card, closeDay) {
  const oc = openCloseMonth_(closeDay, new Date());
  const cyc = cycleByCloseMonth_(closeDay, oc.year, oc.monthIdx);
  return { due: sumCardCharges_(card, cyc.start, cyc.end), closeDate: cyc.closeDate };
}

/** 上一期(已結帳)若未繳，回傳 {amount, dueDate}；已繳或無則 amount 0 */
function cardPendingClosed_(card, closeDay, dueDay) {
  const oc = openCloseMonth_(closeDay, new Date());
  const d = new Date(oc.year, oc.monthIdx - 1, 1);    // open 月的前一個月＝最近結帳的那期
  const cyc = cycleByCloseMonth_(closeDay, d.getFullYear(), d.getMonth());
  const ym = cyc.closeDate.slice(0, 7);
  if (isCyclePaid_(card, ym)) return { amount: 0, dueDate: '' };
  return {
    amount: sumCardCharges_(card, cyc.start, cyc.end),
    dueDate: dueDay ? dueDateFor_(cyc.closeDate, dueDay) : ''
  };
}

/** 讀「信用卡帳單設定」：{卡名: {close, due}}（放在設定分頁 I~K 欄，避開帳戶/類別的讀取範圍） */
function getCardConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(CFG_SHEET);
  if (!cfg) return {};
  ensureCardSection_(cfg);
  const n = Math.max(cfg.getLastRow() - 5, 1);
  const rows = cfg.getRange(6, 9, n, 3).getValues();   // I,J,K = 卡名,結帳日,繳款日
  const map = {};
  rows.forEach(r => {
    if (r[0] === '') return;
    map[String(r[0]).trim()] = { close: Number(r[1]) || 0, due: Number(r[2]) || 0 };
  });
  return map;
}

/** 確保設定分頁有「信用卡帳單」區塊(I4 起)；首次呼叫會建好標題並把目前的負債卡各放一列（日先空白） */
function ensureCardSection_(cfg) {
  if (cfg.getRange('I4').getValue()) return;
  cfg.getRange('I4').setValue('信用卡帳單').setFontWeight('bold').setFontColor(COLORS.sectionText).setFontSize(12);
  cfg.getRange(5, 9, 1, 3).setValues([['卡名', '結帳日', '繳款日']])
     .setBackground(COLORS.section).setFontWeight('bold').setFontColor(COLORS.sectionText);
  const cards = getBalances_().filter(a => a.type === '負債').map(a => [a.name, '', '']);
  if (cards.length) cfg.getRange(6, 9, cards.length, 3).setValues(cards);
  cfg.setColumnWidth(8, 28);
  cfg.setColumnWidth(9, 120); cfg.setColumnWidth(10, 70); cfg.setColumnWidth(11, 70);
  cfg.getRange(4, 9, cards.length + 2, 3).setFontFamily('Noto Sans TC');
}

/** 前端設定某卡的結帳日/繳款日（沒有該列就新增一列） */
function setCardConfig(card, close, due) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = ss.getSheetByName(CFG_SHEET);
  ensureCardSection_(cfg);
  const n = Math.max(cfg.getLastRow() - 5, 1);
  const names = cfg.getRange(6, 9, n, 1).getValues();
  let rowIdx = -1, lastFilled = 5;
  for (let i = 0; i < names.length; i++) {
    if (names[i][0] !== '') lastFilled = 6 + i;
    if (String(names[i][0]).trim() === card) { rowIdx = 6 + i; break; }
  }
  if (rowIdx === -1) { rowIdx = lastFilled + 1; cfg.getRange(rowIdx, 9).setValue(card); }
  cfg.getRange(rowIdx, 10).setValue(close ? Number(close) : '');
  cfg.getRange(rowIdx, 11).setValue(due ? Number(due) : '');
  return { ok: true };
}

/** 卡片帳單頁：某卡、某結帳月(ym='yyyy-MM'，省略=目前未結帳期) 的帳單 */
function getCardStatement(card, ym) {
  const conf = getCardConfig()[card] || { close: 0, due: 0 };
  const closeDay = conf.close, dueDay = conf.due;
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

  let year, monthIdx;
  if (ym) { const p = ym.split('-'); year = +p[0]; monthIdx = +p[1] - 1; }
  else if (closeDay) { const oc = openCloseMonth_(closeDay, new Date()); year = oc.year; monthIdx = oc.monthIdx; }
  else { const n = new Date(); year = n.getFullYear(); monthIdx = n.getMonth(); }

  let cyc;
  if (closeDay) cyc = cycleByCloseMonth_(closeDay, year, monthIdx);
  else cyc = {
    start: Utilities.formatDate(new Date(year, monthIdx, 1, 12, 0, 0), TZ, 'yyyy-MM-dd'),
    end: Utilities.formatDate(new Date(year, monthIdx + 1, 0, 12, 0, 0), TZ, 'yyyy-MM-dd'),
    closeDate: ''
  };
  const dueDate = (closeDay && dueDay) ? dueDateFor_(cyc.closeDate, dueDay) : '';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let txns = [], total = 0, mine = 0, fam = 0;
  monthsBetween_(cyc.start, cyc.end).forEach(mn => {
    const sh = ss.getSheetByName(mn); if (!sh) return;
    readTxns_(sh).forEach(t => {
      if (t.date < cyc.start || t.date > cyc.end) return;
      if (t.type === '支出' && t.accountOut === card) {
        txns.push(t);
        const f = (t.famPct == null ? 0 : t.famPct) / 100;
        total += t.amount; fam += t.amount * f; mine += t.amount * (1 - f);
      }
    });
  });
  txns.sort((a, b) => (a.date + a.time < b.date + b.time ? 1 : -1));
  const ymOut = year + '-' + String(monthIdx + 1).padStart(2, '0');

  return {
    card: card, closeDay: closeDay, dueDay: dueDay,
    closeDate: cyc.closeDate, dueDate: dueDate, start: cyc.start, end: cyc.end,
    ym: ymOut,
    txns: txns, total: total, mine: mine, fam: fam,
    paid: closeDay ? isCyclePaid_(card, ymOut) : false,
    isClosed: !!(cyc.closeDate && cyc.closeDate < today)
  };
}

/** 記錄繳款：依該期付款方拆「我/家」各產生一筆轉帳(來源帳戶→卡)，note 標記該期以利判定已繳 */
function recordCardPayment(card, ym, mineAcct, famAcct) {
  const st = getCardStatement(card, ym);
  if (!st.closeDate) return { ok: false, msg: '請先設定結帳日' };
  if (st.paid) return { ok: false, msg: '這期已記錄過繳款' };
  if (st.total <= 0) return { ok: false, msg: '這期沒有可繳金額' };

  const note = PAY_TAG + st.ym;            // 用結帳月標記，日期微調不影響已繳判定
  const payDate = st.dueDate || st.closeDate;
  const mineAmt = Math.round(st.mine);
  const famAmt = Math.round(st.total) - mineAmt;
  if (mineAmt > 0 && mineAcct) {
    addTransaction({ type: '轉帳', amount: mineAmt, date: payDate, accountOut: mineAcct, accountIn: card, category: '轉帳', note: note });
  }
  if (famAmt > 0 && famAcct) {
    addTransaction({ type: '轉帳', amount: famAmt, date: payDate, accountOut: famAcct, accountIn: card, category: '轉帳', note: note });
  }
  return { ok: true };
}

/** 刪除某月某一列（前端清單的操作選單用） */
function deleteRow(month, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(month);
  if (sh && row >= 2 && row <= sh.getLastRow()) sh.deleteRow(row);
  return { ok: true };
}

/** 更新已存在的某筆紀錄（前端編輯小視窗用）
 *  oldMonth/row＝原本所在的月份分頁與列號；t＝改好的新內容。
 *  若日期被改到別的月份，會把這筆搬到正確的月份分頁（刪舊、寫新）。
 *  建立時間欄沿用原本的值，不因編輯而改變。 */
function updateRow(oldMonth, row, t) {
  const amt = Number(t.amount);
  if (!t.type) throw new Error('缺少交易類型');
  if (!(amt > 0)) throw new Error('金額需大於 0');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const oldSh = ss.getSheetByName(oldMonth);
  if (!oldSh || row < 2 || row > oldSh.getLastRow()) throw new Error('找不到這筆紀錄');

  const created = oldSh.getRange(row, CREATED_COL).getValue() || new Date();
  const dateObj = parseDate_(t.date);
  const newMonth = Utilities.formatDate(dateObj, TZ, 'yyyy-MM');
  const rowVals = [
    dateObj,
    t.type,
    t.accountOut || '',
    t.accountIn || '',
    amt,
    t.category || (t.type === '轉帳' ? '轉帳' : ''),
    t.note || '',
    created,
    famPct_(t)
  ];

  if (newMonth === oldMonth) {
    oldSh.getRange(row, 1, 1, TXN_HEADERS.length).setValues([rowVals]);
  } else {
    // 日期改到別的月份：搬到新月份分頁，再刪掉舊列
    getOrCreateMonthSheet_(newMonth).appendRow(rowVals);
    oldSh.deleteRow(row);
  }
  return { ok: true, month: newMonth };
}