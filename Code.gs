/**
 * 大藏經對照表：CBETA 目錄 × 如是我聞（rushiwowen.co）介紹
 *
 * 使用方式：
 *   1. 在試算表「擴充功能 > Apps Script」貼上本檔，儲存。
 *   2. 重新整理試算表，上方會出現「大藏經對照」選單。
 *   3. 依序執行選單 1 → 2 → 3 →（可選）4 → 5。
 *      步驤 2、4、5 會自動分段執行，超過時間會自己排下一次，不必守著。
 *   4. 部署 > 新增部署 > 網頁應用程式，執行身分「我」，存取權「任何人」，
 *      取得 /exec 網址，前端就用它讀 JSON。
 */

const CFG = {
  SHEET_MAIN: '目錄',
  SHEET_RW: '如是我聞',
  SHEET_LOG: '日誌',
  RW_API: 'https://rushiwowen.co/api/books',
  RW_LIMIT: 100,
  RW_PAGE: 'https://rushiwowen.co/r/',
  CBETA_ALL: 'https://cbdata.dila.edu.tw/stable/download/all-works.json',
  CBETA_WORK: 'https://cbdata.dila.edu.tw/stable/works?work=',
  CBETA_ONLINE: 'https://cbetaonline.dila.edu.tw/zh-tw/',
  REFERER: 'https://script.google.com/', // CBETA API 要求帶 Referer 供流量統計
  TIME_BUDGET_MS: 4.5 * 60 * 1000,       // GAS 單次上限 6 分鐘，留餘裕
  ENRICH_BATCH: 20,                      // 步驤 4 每批平行查詢筆數
};

// 「目錄」工作表欄位（1-based 欄號）
const COL = {
  WORK: 1,        // A 經號
  TITLE: 2,       // B 經名(CBETA)
  CREATOR: 3,     // C 譯者/作者
  DYNASTY: 4,     // D 朝代
  CATEGORY: 5,    // E 部類
  JUAN: 6,        // F 卷數
  CANON: 7,       // G 藏經
  DESC_S: 8,      // H 介紹(簡)
  DESC_T: 9,      // I 介紹(繁)
  MY_NOTE: 10,    // J 我的註釋（手寫，程式永不覆寫）
  RW_URL: 11,     // K 如是我聞連結
  CBETA_URL: 12,  // L CBETA 連結
  STATUS: 13,     // M 狀態
  UPDATED: 14,    // N 更新時間
};
const MAIN_HEADERS = [
  '經號', '經名(CBETA)', '譯者/作者', '朝代', '部類', '卷數', '藏經',
  '介紹(簡)', '介紹(繁)', '我的註釋', '如是我聞連結', 'CBETA連結', '狀態', '更新時間',
];
const RW_HEADERS = [
  'code', 'title', 'translator', 'author', 'dynasty', 'description',
  'slug', 'id', 'totalPages', 'popularity', 'collection', 'createdAt',
];

// ---------------------------------------------------------------- 選單

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('大藏經對照')
    .addItem('1. 載入 CBETA 全目錄', 'step1_loadCbetaWorks')
    .addItem('2. 抓取如是我聞介紹（自動分段）', 'step2_fetchRushiwowen')
    .addItem('3. 依經號合併', 'step3_merge')
    .addSeparator()
    .addItem('4. 補 CBETA 譯者/朝代/部類（可選，較久）', 'step4_enrichCbeta')
    .addItem('5. 介紹簡轉繁（可選，Google 翻譯）', 'step5_convertTraditional')
    .addSeparator()
    .addItem('重設進度並清除排程', 'resetProgress')
    .addToUi();
}

// ---------------------------------------------------------------- 步驤 1

function step1_loadCbetaWorks() {
  const sh = getOrCreateSheet_(CFG.SHEET_MAIN, MAIN_HEADERS);
  const res = UrlFetchApp.fetch(CFG.CBETA_ALL, {
    headers: { Referer: CFG.REFERER },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) throw new Error('CBETA all-works.json HTTP ' + res.getResponseCode());
  const works = JSON.parse(res.getContentText());

  const existing = readColumnMap_(sh, COL.WORK); // 經號 -> 列號
  const now = new Date();
  const newRows = [];
  works.forEach(w => {
    if (existing[w.work]) return; // 已有就不重複
    const row = new Array(MAIN_HEADERS.length).fill('');
    row[COL.WORK - 1] = w.work;
    row[COL.TITLE - 1] = w.title || '';
    row[COL.JUAN - 1] = (w.juans || []).length;
    row[COL.CANON - 1] = canonOf_(w.work);
    row[COL.CBETA_URL - 1] = CFG.CBETA_ONLINE + w.work;
    row[COL.STATUS - 1] = '尚未合併';
    row[COL.UPDATED - 1] = now;
    newRows.push(row);
  });
  if (newRows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, newRows.length, MAIN_HEADERS.length).setValues(newRows);
  }
  sh.setFrozenRows(1);
  log_('步驤1', `CBETA 目錄共 ${works.length} 部，新增 ${newRows.length} 列`);
  toast_(`CBETA 目錄載入完成：共 ${works.length} 部，新增 ${newRows.length} 列`);
}

// ---------------------------------------------------------------- 步驤 2

function step2_fetchRushiwowen() {
  const start = Date.now();
  const props = PropertiesService.getScriptProperties();
  const sh = getOrCreateSheet_(CFG.SHEET_RW, RW_HEADERS);

  let page = Number(props.getProperty('RW_NEXT_PAGE') || 1);
  if (page === 1) {
    // 從頭抓：清掉舊資料只留表頭
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
  }

  let totalPages = Number(props.getProperty('RW_TOTAL_PAGES') || 0);
  let fetched = 0;

  while (true) {
    if (totalPages && page > totalPages) break;
    if (Date.now() - start > CFG.TIME_BUDGET_MS) {
      props.setProperty('RW_NEXT_PAGE', String(page));
      scheduleContinue_('step2_fetchRushiwowen');
      log_('步驤2', `時間到，已抓到第 ${page - 1} 頁，1 分鐘後續抓`);
      toast_(`如是我聞已抓到第 ${page - 1}/${totalPages} 頁，1 分鐘後自動續抓`);
      return;
    }

    const url = `${CFG.RW_API}?page=${page}&limit=${CFG.RW_LIMIT}`;
    const res = UrlFetchApp.fetch(url, {
      headers: { 'User-Agent': 'Mozilla/5.0 (personal sutra index; GAS)' },
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) {
      log_('步驤2', `第 ${page} 頁 HTTP ${res.getResponseCode()}，稍後重試`);
      props.setProperty('RW_NEXT_PAGE', String(page));
      scheduleContinue_('step2_fetchRushiwowen');
      return;
    }
    const json = JSON.parse(res.getContentText());
    const data = json.data || {};
    const books = data.books || [];
    if (!totalPages && data.pagination) {
      totalPages = Number(data.pagination.totalPages || 0);
      props.setProperty('RW_TOTAL_PAGES', String(totalPages));
    }
    if (!books.length) break;

    const rows = books.map(b => [
      b.code || '',
      b.title || '',
      b.translator || '',
      b.author || '',
      b.dynasty || '',
      b.description || '',
      b.slug || '',
      b.id || '',
      b.totalPages || '',
      b.popularity || '',
      (b.collection && b.collection.code) || '',
      b.createdAt || '',
    ]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, RW_HEADERS.length).setValues(rows);
    fetched += rows.length;
    page++;
    Utilities.sleep(300); // 對人家的伺服器客氣一點
  }

  props.deleteProperty('RW_NEXT_PAGE');
  props.deleteProperty('RW_TOTAL_PAGES');
  clearTriggers_('step2_fetchRushiwowen');
  sh.setFrozenRows(1);
  log_('步驤2', `如是我聞抓取完成，本次寫入 ${fetched} 筆，工作表共 ${sh.getLastRow() - 1} 筆`);
  toast_(`如是我聞抓取完成，共 ${sh.getLastRow() - 1} 筆`);
}

// ---------------------------------------------------------------- 步驤 3

function step3_merge() {
  const main = getOrCreateSheet_(CFG.SHEET_MAIN, MAIN_HEADERS);
  const rw = getOrCreateSheet_(CFG.SHEET_RW, RW_HEADERS);
  if (main.getLastRow() < 2) throw new Error('「目錄」是空的，請先執行步驤 1');
  if (rw.getLastRow() < 2) throw new Error('「如是我聞」是空的，請先執行步驤 2');

  // 如是我聞：code -> 資料
  const rwVals = rw.getRange(2, 1, rw.getLastRow() - 1, RW_HEADERS.length).getValues();
  const rwMap = {};
  rwVals.forEach(r => {
    const code = String(r[0]).trim();
    if (code) rwMap[code] = { title: r[1], translator: r[2], author: r[3], dynasty: r[4], description: r[5], slug: r[6], id: r[7] };
  });

  const n = main.getLastRow() - 1;
  const vals = main.getRange(2, 1, n, MAIN_HEADERS.length).getValues();
  const now = new Date();
  let matched = 0;
  const seen = {};

  vals.forEach(row => {
    const work = String(row[COL.WORK - 1]).trim();
    const hit = rwMap[work];
    if (hit) {
      seen[work] = true;
      matched++;
      row[COL.DESC_S - 1] = hit.description || '';
      row[COL.RW_URL - 1] = CFG.RW_PAGE + (hit.slug || hit.id);
      // 若 CBETA 譯者尚未補上，先借用如是我聞的
      if (!row[COL.CREATOR - 1]) row[COL.CREATOR - 1] = hit.translator || hit.author || '';
      if (!row[COL.DYNASTY - 1]) row[COL.DYNASTY - 1] = hit.dynasty || '';
      row[COL.STATUS - 1] = '已對應';
    } else {
      row[COL.STATUS - 1] = '如是我聞無介紹';
    }
    if (!row[COL.CBETA_URL - 1]) row[COL.CBETA_URL - 1] = CFG.CBETA_ONLINE + work;
    row[COL.UPDATED - 1] = now;
    // 注意：I 介紹(繁)、J 我的註釋 不在這裡動
  });
  main.getRange(2, 1, n, MAIN_HEADERS.length).setValues(vals);

  // 如是我聞有、CBETA 沒有的 code，另列一張表
  const orphan = Object.keys(rwMap).filter(c => !seen[c]).map(c => [c, rwMap[c].title, rwMap[c].translator, rwMap[c].dynasty, CFG.RW_PAGE + (rwMap[c].slug || rwMap[c].id)]);
  const osh = getOrCreateSheet_('未對應', ['code', 'title', 'translator', 'dynasty', '連結']);
  if (osh.getLastRow() > 1) osh.getRange(2, 1, osh.getLastRow() - 1, 5).clearContent();
  if (orphan.length) osh.getRange(2, 1, orphan.length, 5).setValues(orphan);

  log_('步驤3', `合併完成：CBETA ${n} 部，對到介紹 ${matched} 部，如是我聞另有 ${orphan.length} 筆 CBETA 查無`);
  toast_(`合併完成：${matched}/${n} 部有介紹，${orphan.length} 筆對不上（見「未對應」）`);
}

// ---------------------------------------------------------------- 步驤 4（可選）

function step4_enrichCbeta() {
  const start = Date.now();
  const props = PropertiesService.getScriptProperties();
  const sh = getOrCreateSheet_(CFG.SHEET_MAIN, MAIN_HEADERS);
  const n = sh.getLastRow() - 1;
  if (n < 1) throw new Error('「目錄」是空的，請先執行步驤 1');

  let rowIdx = Number(props.getProperty('ENRICH_ROW') || 2);
  const works = sh.getRange(2, COL.WORK, n, 1).getValues().map(r => String(r[0]).trim());
  const cats = sh.getRange(2, COL.CATEGORY, n, 1).getValues().map(r => String(r[0]));
  let done = 0;

  while (rowIdx <= n + 1) {
    if (Date.now() - start > CFG.TIME_BUDGET_MS) {
      props.setProperty('ENRICH_ROW', String(rowIdx));
      scheduleContinue_('step4_enrichCbeta');
      log_('步驤4', `時間到，處理到第 ${rowIdx} 列，1 分鐘後續跑`);
      toast_(`CBETA 補資料進度 ${rowIdx - 2}/${n}，1 分鐘後自動續跑`);
      return;
    }
    // 收集一批尚未有部類的列
    const batch = [];
    while (batch.length < CFG.ENRICH_BATCH && rowIdx <= n + 1) {
      const i = rowIdx - 2;
      if (!cats[i]) batch.push({ row: rowIdx, work: works[i] });
      rowIdx++;
    }
    if (!batch.length) continue;

    const reqs = batch.map(b => ({
      url: CFG.CBETA_WORK + encodeURIComponent(b.work),
      headers: { Referer: CFG.REFERER },
      muteHttpExceptions: true,
    }));
    let resps;
    try { resps = UrlFetchApp.fetchAll(reqs); }
    catch (err) {
      // CBETA API 暫時連不上：記住進度，1 分鐘後續跑
      props.setProperty('ENRICH_ROW', String(batch[0].row));
      scheduleContinue_('step4_enrichCbeta');
      log_('步驤4', `連線失敗（${err.message}），1 分鐘後續跑，進度 ${batch[0].row - 2}/${n}`);
      return;
    }
    resps.forEach((res, k) => {
      const b = batch[k];
      if (res.getResponseCode() !== 200) return;
      let info;
      try { info = JSON.parse(res.getContentText()); } catch (e) { return; }
      const r = (info.results || [])[0];
      if (!r) {
        sh.getRange(b.row, COL.CATEGORY).setValue('(CBETA API 查無)');
        return;
      }
      sh.getRange(b.row, COL.CREATOR, 1, 3).setValues([[
        r.byline || r.creators || '',
        r.time_dynasty || '',
        r.category || r.orig_category || '',
      ]]);
      done++;
    });
    Utilities.sleep(200);
  }

  props.deleteProperty('ENRICH_ROW');
  clearTriggers_('step4_enrichCbeta');
  log_('步驤4', `CBETA 補資料完成，本次更新 ${done} 列`);
  toast_('CBETA 譯者/朝代/部類補齊完成');
}

// ---------------------------------------------------------------- 步驤 5（可選）

/**
 * 用 Google 翻譯做 zh-CN → zh-TW。它是「翻譯」不是純字元轉換，偶爾會改寫用語，
 * 且有每日配額；配額用完會停下並自動排隔天再跑。只填 I 欄為空的列。
 * 若你想要純字元轉換，改在前端用 opencc-js 即可，這一步可以不跑。
 */
function step5_convertTraditional() {
  const start = Date.now();
  const props = PropertiesService.getScriptProperties();
  const sh = getOrCreateSheet_(CFG.SHEET_MAIN, MAIN_HEADERS);
  const n = sh.getLastRow() - 1;
  let rowIdx = Number(props.getProperty('CONV_ROW') || 2);
  const descS = sh.getRange(2, COL.DESC_S, n, 1).getValues();
  const descT = sh.getRange(2, COL.DESC_T, n, 1).getValues();
  let done = 0;

  for (; rowIdx <= n + 1; rowIdx++) {
    if (Date.now() - start > CFG.TIME_BUDGET_MS) {
      props.setProperty('CONV_ROW', String(rowIdx));
      scheduleContinue_('step5_convertTraditional');
      toast_(`簡轉繁進度 ${rowIdx - 2}/${n}，1 分鐘後自動續跑`);
      return;
    }
    const i = rowIdx - 2;
    const s = String(descS[i][0] || '');
    if (!s || descT[i][0]) continue;
    try {
      const t = LanguageApp.translate(s, 'zh-CN', 'zh-TW');
      sh.getRange(rowIdx, COL.DESC_T).setValue(t);
      done++;
    } catch (e) {
      // 多半是配額用完
      props.setProperty('CONV_ROW', String(rowIdx));
      clearTriggers_('step5_convertTraditional');
      ScriptApp.newTrigger('step5_convertTraditional').timeBased().after(24 * 60 * 60 * 1000).create();
      log_('步驤5', `翻譜配額可能用完（${e.message}），已排 24 小時後續跑，進度 ${rowIdx - 2}/${n}`);
      toast_('Google 翻譯配額用完，24 小時後自動續跑');
      return;
    }
  }
  props.deleteProperty('CONV_ROW');
  clearTriggers_('step5_convertTraditional');
  log_('步驤5', `簡轉繁完成，本次轉 ${done} 筆`);
  toast_('簡轉繁完成');
}

// ---------------------------------------------------------------- Web App JSON 輸出

/**
 * 部署為網頁應用程式後：
 *   GET /exec                 → 全部目錄（JSON 陣列）
 *   GET /exec?work=T0915      → 單筆
 *   GET /exec?q=金剛          → 經號或經名包含關鍵字
 *   GET /exec?lite=1          → 不含介紹欄，體積小，適合先載清單
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const sh = ss_().getSheetByName(CFG.SHEET_MAIN);
  const q = (p.q || '').trim();
  const work = (p.work || '').trim();
  const lite = p.lite === '1';
  const out = [];
  if (sh && sh.getLastRow() > 1) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, MAIN_HEADERS.length).getValues();
    vals.forEach(r => {
      const id = String(r[COL.WORK - 1]);
      const title = String(r[COL.TITLE - 1]);
      if (work && id !== work) return;
      if (q && id.indexOf(q) < 0 && title.indexOf(q) < 0) return;
      const o = {
        work: id,
        title: title,
        creator: r[COL.CREATOR - 1],
        dynasty: r[COL.DYNASTY - 1],
        category: r[COL.CATEGORY - 1],
        juan: r[COL.JUAN - 1],
        canon: r[COL.CANON - 1],
        rw_url: r[COL.RW_URL - 1],
        cbeta_url: r[COL.CBETA_URL - 1],
        status: r[COL.STATUS - 1],
      };
      if (!lite) {
        o.desc_s = r[COL.DESC_S - 1];
        o.desc_t = r[COL.DESC_T - 1];
        o.note = r[COL.MY_NOTE - 1];
      }
      out.push(o);
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify(work ? (out[0] || null) : out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------- 工具

function resetProgress() {
  const props = PropertiesService.getScriptProperties();
  ['RW_NEXT_PAGE', 'RW_TOTAL_PAGES', 'ENRICH_ROW', 'CONV_ROW'].forEach(k => props.deleteProperty(k));
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  toast_('進度與排程已重設');
}

function getOrCreateSheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readColumnMap_(sh, col) {
  const map = {};
  const n = sh.getLastRow() - 1;
  if (n < 1) return map;
  sh.getRange(2, col, n, 1).getValues().forEach((r, i) => {
    const k = String(r[0]).trim();
    if (k) map[k] = i + 2;
  });
  return map;
}

function canonOf_(work) {
  const m = String(work).match(/^[A-Za-z]+/);
  return m ? m[0] : '';
}

function scheduleContinue_(fnName) {
  clearTriggers_(fnName);
  ScriptApp.newTrigger(fnName).timeBased().after(60 * 1000).create();
}

function clearTriggers_(fnName) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === fnName)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function log_(step, msg) {
  const sh = getOrCreateSheet_(CFG.SHEET_LOG, ['時間', '步驤', '訊息']);
  sh.appendRow([new Date(), step, msg]);
  console.log(`[${step}] ${msg}`);
}

function toast_(msg) {
  try { ss_().toast(msg, '大藏經對照', 8); } catch (e) { /* 觸發器執行時沒有 UI */ }
}

// ---------------------------------------------------------------- 試算表來源
// 若此專案綁定在試算表上會直接用 getActive()，否則用 ID 開啟
const SHEET_ID = '1kb4sh6Ct2PVFIBTf33VzkbUFBAjcCZV_lhUXbTVwuxA';
function ss_() {
  return SpreadsheetApp.getActive() || SpreadsheetApp.openById(SHEET_ID);
}
