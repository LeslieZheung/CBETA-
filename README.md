# 大藏經對照

**網站：https://lesliezheung.github.io/Tripitaka/**

以 CBETA 經號為鑰，把每一部經的「導讀」對照到 CBETA 原文。
首頁依 CBETA「依據部類」23 個部類逐層瀏覽（部類 → 子分類 → 經典），或直接搜尋經號、經名、譯者。
純靜態網頁（GitHub Pages）＋ Google Apps Script 提供 JSON，沒有任何 AI API。

## 這是什麼

- **CBETA 全目錄**：5,748 部（大正藏 T、卍續藏 X、嘉興藏 J 等），來自 CBETA API 的 `all-works.json`。
- **導讀**：取自 [如是我聞](https://rushiwowen.co/)（rushiwowen.co）每部經的「關於《…》」段落，以 CBETA 經號對應，約 4,460 部有導讀。前端用 opencc-js 即時轉為正體。
- **我的註釋**：試算表 J 欄，可自行撰寫，網站會顯示在導讀下方。
- **對照閱讀**：每部經提供「在 CBETA 讀經文」「如是我聞原頁」「複製經名 → 佛光大藏經 / 星雲大師全集」。

## 使用

- 開啟網站後可搜尋經號、經名、譯者、部類，或依藏經篩選。
- 在 CBETA 看到經號時，在網址後加 `#T0915` 可直接開啟該部經。
- 快捷鍵 `/` 聚焦搜尋。手機版單經頁底部有操作列。

## 架構

```
Google Sheet（目錄 / 如是我聞 / 未對應 / 日誌）
   ↑ Apps Script（Code.gs）
   │   step1  CBETA all-works.json → 目錄
   │   step2  rushiwowen /api/books → 如是我聞
   │   step3  依經號合併
   │   step4  CBETA API 補譯者 / 朝代 / 部類（可選）
   │   step5  簡轉繁（可選）
   │   doGet  輸出 JSON（?work= / ?q= / ?lite=1）
   ↓
index.html（GitHub Pages）── fetch JSON ── 搜尋 / 卡片 / 單經導讀
```

## 檔案

- `index.html`：整個前端，單一檔案，無建置步驤。頂端 `DEFAULT_API` 為 Apps Script 網頁應用程式的 `/exec` 網址。
- `catalog.json`：CBETA「依據部類」目錄樹（23 部類、900 子分類、5,332 部），由 [CBETA 官方部類目錄](https://github.com/heavenchou/cbwork-bin/blob/master/cbreader2X/bulei/bulei.txt) 轉出。
- `Code.gs`：Apps Script 原始碼。若試算表不是綁定專案，會用 `SHEET_ID` 開啟。

## 資料來源與授權

- 目錄、經號、經名：[CBETA 中華電子佛典協會](https://cbetaonline.dila.edu.tw/)，CC BY-NC-SA。
- 導讀文字：[如是我聞](https://rushiwowen.co/)，為 AI 生成內容，僅供學習參考；本站保留出處並回連原頁。
- 本專案為個人研習用途，不作商業使用。
