# PPT 製作流程記錄

> 目的：記錄本次 PARK24 / Times 外部提案簡報的實際製作流程，作為之後使用 Codex 製作商務 PPT 的標準工作法。  
> 核心經驗：不要一開始直接做 PPTX；先整理外部提案邏輯，再用 HTML/CSS 做可視化預覽，確認內容與風格後，再用高解析截圖方式轉成 PPTX，避免跑版。

---

> ### ⚠️ AI 執行規則：本文件只提供工作流，不提供可複用的內容
>
> 本文件記錄的是「製作流程、Pattern 與品質規則」，不是任何品牌的定位、文案或頁面內容範本。
>
> **每個品牌的內容（定位語言、數字、合作訴求、頁面邏輯）必須來自該品牌的 `00_來源資料`，不可從本文件或其他品牌的專案資料夾借用。**
>
> 本文件中出現「PARK24」「城市車旅」「嘟嘟房」等品牌名稱的段落，均為「歷史案例記錄，說明當時如何修正工作流」，不代表這些品牌的用詞或邏輯可以套用到其他品牌。
>
> AI 執行任何新品牌專案時，應：
> 1. 讀本文件 → 了解工作流順序與品質規則
> 2. 讀當前品牌的 `00_來源資料` → 取得該品牌的所有內容素材
> 3. 不得將本文件中任何品牌的具體文案、定位語言、頁面標題直接複製使用

---

## 0. 前置步驟：DD Execution Guide（所有新專案必讀）

任何新外部提案專案在進入 BD 自動化工作流之前，必須先完成此前置步驟。

### Reading Index：案場選點任務必讀

若任務涉及以下內容，AI 必須先讀：

```text
PARKING_SITE_SELECTION_PLAYBOOK.md
```

適用任務：

- 停車場選點
- 案場資料庫
- 候選場站
- TOP10 / TOP3 初選
- 桃園 / 台中場站評估
- 供電分析
- 現勘清單

核心原因：

```text
停車場官網常把案場資料藏在 JavaScript、API 或動態搜尋頁中；
官方案場存在也不等於適合試點。
AI 必須先完成官方來源解析、資料來源分級、事實與推論切分，
再進行麥肯錫式選點評估。
```

### 為什麼需要 DD Execution Guide？

PARK24 和城市車旅的成功流程顯示：

`00_來源資料` 必須先有真實的調查文件，後續的定位分析、OnePager、Coldmail 才能做到品牌客製化。若跳過此步驟直接進入工作流，AI 只能產出通用模板，無法反映對象品牌的實際特性。

### 三層結構（含案場資料庫）

```text
Layer 1  DD Execution Guide
         → 對目標品牌進行商業合作盡職調查（13 個區塊）
         → 可用 AI 執行網路公開情報調查
         → 產出 00_來源資料 的核心文件：
           DD_REPORT.md / BD_STRATEGY.md / PROJECT_MASTER.md / HEADQUARTER_APPROACH.md

Layer 1.5  案場資料庫（新增，必要）
         → 在 DD 完成後、BD 工作流啟動前建立
         → 目標：建立目標品牌在【桃園 / 台中】的候選場站資料庫
         → 必讀：PARKING_SITE_SELECTION_PLAYBOOK.md
         → 先確認官方案場來源、動態頁面 / API / script 資料、資料來源分級
         → 格式：Excel（參考 PARK24_Project_Master.xlsx + CITY_案場資料庫_v1.xlsx）
         → 工作表：公司DD摘要 / 決策鏈 / 台中場站庫 / 桃園場站庫 / 評分模型 /
                   TOP10初選 / TOP3建議試點 / 供電分析 / 風險登錄 / 行動計畫 / Source_Assumptions
         → 若品牌場站主要在其他城市，改為對應城市
         → 產出：{品牌}_案場資料庫_v1.xlsx，放入 00_來源資料/

Layer 2  BD 自動化工作流（Step -0.5 至 8）
         → Step -1：DD 調查建立與來源資料產出（Layer 1 產出）
         → Step -0.5：案場資料庫建立（Layer 1.5 產出）← 新增
         → Step 0：資料盤點
         → Step 1：專案定位分析
         → Step 2-8：OnePager → Coldmail → PPT
         → 每節完成後停下等使用者確認

Layer 3  PPT 提案
         → 有料、有架構、有風格後才進入 HTML → PPTX
```

### 各品牌文件位置

| 品牌 | DD 執行指南 | 案場資料庫 |
|------|------------|-----------|
| PARK24 / Times | `PARK24_Times_提案專案/00_來源資料/PARK24_DD_Execution_Guide_v1.md.txt` | `PARK24_Times_提案專案/00_來源資料/PARK24_Project_Master.xlsx` |
| 城市車旅 | `城市車旅_提案專案/00_來源資料/CITY_PARKING_DD_EXECUTION_GUIDE.txt` | `城市車旅_提案專案/00_來源資料/CITY_案場資料庫_v1.xlsx` |
| 嘟嘟房 | `嘟嘟房_提案專案/00_來源資料/DODO_DD_EXECUTION_GUIDE.md` | `嘟嘟房_提案專案/00_來源資料/DODO_案場資料庫_v1.xlsx` ✅ 已建立 |
| 其他新品牌 | `{品牌}_提案專案/00_來源資料/{品牌}_DD_EXECUTION_GUIDE.md` | `{品牌}_提案專案/00_來源資料/{品牌}_案場資料庫_v1.xlsx` |

### 新品牌完整建立流程

1. 複製既有 DD Execution Guide（建議以 DODO 版為範本）
2. 依目標品牌調整合作核心假設、市場定位角度、比較對象
3. 交給 AI 執行 13 個調查區塊（可用 WebSearch / DeepSeek / ChatGPT）
4. 產出文件放入 `00_來源資料/`
5. **建立案場資料庫 Excel**（桃園 / 台中候選場站，需先讀 `PARKING_SITE_SELECTION_PLAYBOOK.md`）
   - 參考：PARK24_Project_Master.xlsx（供電分析格式）
   - 參考：CITY_案場資料庫_v1.xlsx（評分模型格式）
   - 先檢查官方頁面是否有動態資料、script、API、城市篩選或分頁限制
   - 工作表：公司DD / 決策鏈 / 台中場站庫 / 桃園場站庫 / 評分模型 / TOP10 / TOP3 / 供電分析 / 風險 / 行動計畫 / Source_Assumptions
   - 回報時必須分清楚：官方事實、AI 推論、待現勘確認、不建議優先
6. 確認品質後才進入 BD 自動化工作流 Step 0

---

## 1. 先建立任務定位

開始前要先判斷這份簡報的用途。

> **通用原則：** 開始任何外部提案前，先釐清以下三個問題，再決定要讀什麼資料、產出什麼文件。

```text
這份簡報是要給誰看？
對方看完要做什麼決定？
這是內部研究、外部提案、募資、銷售、還是會議報告？
```

> 📌 **PARK24/Times 案例記錄（參考用，不作為其他品牌定位範本）：**
>
> 當次簡報定位為對外提案，目標對象是 PARK24/Times 總部合作窗口，目標不是內部研究也不是 SOP，而是讓對方認同進入 1 至 3 個試點場站評估。

之後遇到類似案子，要先問：

```text
這份簡報是要給誰看？
對方看完要做什麼決定？
這是內部研究、外部提案、募資、銷售、還是會議報告？
```

如果是外部提案，不能把內部 BD、DD、風險登錄、資料來源、推案路徑直接放上主簡報。

## 2. 先讀來源資料，再做大綱

> **通用原則：** 每個品牌的來源資料都在該品牌的 `00_來源資料/` 資料夾內。AI 執行時，讀當前品牌的 `00_來源資料`，不要讀其他品牌的資料夾。

> 📌 **PARK24/Times 案例記錄（當時讀的檔案，不是通用指引）：**
> - `PARK24_Times_提案專案/00_來源資料/PARK24_Headquarter_Proposal.docx`
> - `PARK24_Times_提案專案/00_來源資料/PARK24_BD_Strategy.docx`
> - `PARK24_Times_提案專案/00_來源資料/PARK24_Project_Master.xlsx`
> - `風格/` 內既有公司簡報

正確做法（通用）：

1. 先抽出 Word 主張。
2. 再讀 Excel 的候選場站、供電、風險、行動計畫。
3. 先產 Markdown 大綱，不直接做 PPT。
4. 每頁先定義：
   - 結論式標題
   - 核心訊息
   - 支撐點
   - 視覺方向
   - 資料來源

本次產出：

- `PARK24_Times_提案專案/PARK24_PPT_GENERATION_PROMPT.md`
- `PARK24_Times_提案專案/PARK24_PPT_OUTLINE.md`

## 3. 大綱要先分清楚內部與外部語氣

初版常見問題（通用）：

- 太像內部推案簡報。
- 出現 BD、DD、Project Master、來源文件等內部語氣。
- 太早談候選場站與怎麼執行。
- 對外部客戶來說，還沒建立「為什麼要合作」。

修正原則（通用）：

- 不要先講旺來想做什麼。
- 要先講「為什麼找這個品牌」。
- 不要把合作說成租地、租格位或設備導入。
- 不要過早討論候選場站、設備、價格或細節。
- 先建立合作必要性。

> ⚠️ 注意：每個品牌的合作定位名稱都不同（「社區型智慧服務節點」是 PARK24/Times 的版本；嘟嘟房用「複合型智慧服務節點」）。合作定位的具體名稱必須來自當前品牌的 `BD_Positioning.md` 或 `OnePager.md`，不可從本文件借用其他品牌的用詞。

外部提案主軸的通用邏輯應該是：

```text
[對方品牌] 如何透過旺來建立新的場站價值。
```

而不是：

```text
旺來想放設備。
```

## 4. 依照外部提案邏輯重組頁序

外部提案通用頁序邏輯：

```text
先建立產業趨勢（對方的市場為什麼現在需要改變）
→ 為什麼找「這個品牌」（具體理由，來自該品牌 DD 報告）
→ 旺來能帶來什麼能力（旺來的優勢）
→ 對方能得到什麼具體價值（用對方視角描述）
→ 為什麼先小規模試點（降低決策門檻）
→ 最後才談執行細節與確認項目
```

> 📌 **PARK24/Times 案例記錄（16 頁實際大綱，僅供了解工作流，不作為其他品牌頁序模板）：**
>
> 1. 封面 / 2. 合作提案摘要 / 3. 停車場產業的下一個競爭 / 4. 為什麼是 PARK24？ / 5. 旺來提供的是什麼能力？ / 6. PARK24 可以獲得什麼？ / 7. 如果合作成功 / 8. 為什麼先試點？ / 9. Times 場域合作優勢 / 10. 合作定位：不是租格位 / 11. 場站設置原則 / 12. 對 Times 的合作價值 / 13. 90 天試點驗證方式 / 14. 試點前共同確認項目 / 15. 會後建議決策方向 / 16. 附錄：首波候選場站方向
>
> ⚠️ 其他品牌（嘟嘟房、城市車旅等）的頁面標題、頁序、重點，必須依照各自 DD_Positioning.md 與 OnePager.md 重新設計，不可參照上面這個 PARK24 大綱。

## 5. 外部提案避免的字眼

本次修正中特別注意：

避免：

- 不要求 PARK24 建立新團隊
- PARK24 不需要做什麼
- 我們不是來推銷設備
- 由 PARK24 判斷是否擴大
- 租地
- 放智取櫃
- 設備導入案

較佳說法：

- 旺來提出可先小規模驗證的合作模式。
- 我們是來推驗證合作模式。
- 雙方共同決定是否擴大、調整或停止。
- 社區型智慧服務節點。
- 低干擾場站加值。
- 1 至 3 個試點場站評估。

原則：

```text
對大型外部合作對象，不要替對方預設內部負擔；
要用合作、驗證、共同評估、可控試點的語氣。
```

## 6. 風格要先參考既有 PPT，不要自創

本次一開始風格錯誤：

- 太像 SaaS 商務卡片。
- 深藍灰、卡片、儀表板感太重。
- 不像旺來既有簡報。

後來改正方法：

1. 讀取 `風格/` 資料夾內既有 PPT。
2. 匯出參考簡報截圖。
3. 抓到旺來風格特徵：
   - 白底
   - 左上大黑標
   - 右上 waGas 橘色 Logo
   - 橘色副標
   - 灰色分隔線
   - 底部品牌精神與頁碼
   - 橘色幾何封面
   - 條列、流程箭頭、表格、對照圖

之後做旺來 PPT，優先套這種風格，不要另起爐灶。

通用素材與風格參考位置：

- `shared_assets/`：放旺來通用品牌素材，例如 waGas Logo、品牌精神圖。
- `wagas_style_reference/`：放從既有旺來 PPT 萃取出的風格截圖與圖片。
- `風格/`：放原始參考 PPT。

子專案產物不要混入通用素材：

- `ppt_export_slides/` 屬於單一 PPT 的逐頁輸出，應保留在各子專案內。
- 客戶專案的來源資料、PDF 指令、HTML、Final PPTX，應保留在該客戶子專案內。

## 7. HTML 預覽先於 PPTX

正確製作方式：

1. 先做 `PARK24_Times_提案專案/PARK24_PPT_PREVIEW_EXTERNAL.html`。
2. HTML 每個 `.slide` 是固定 16:9。
3. 用 CSS 控制版面。
4. 先在 HTML 階段調整文字、風格、插圖、頁序。
5. 等使用者確認 Final，再轉 PPTX。

原因：

- HTML 修改快。
- 容易控制視覺與排版。
- 可快速重組頁序。
- 不會被 PowerPoint 文字框拖慢。

## 8. 插圖與視覺補強方式

本次有效做法：

- 不直接抓網路圖片。
- 參考智慧停車 / parking kiosk / smart parking illustration 的視覺語彙。
- 用 CSS/HTML 自製插圖：
  - 停車格
  - 車道
  - 車輛
  - 智慧服務據點
  - 社區建物
  - 手機 QR
  - 流程箭頭
  - 對照圖

這樣可以避免：

- 版權問題
- 圖片風格不一致
- 外部載入失敗
- 簡報檔案不完整

插圖原則：

```text
外部提案頁不要只放文字和表格。
每 2 到 3 頁至少要有一頁圖像化解釋：
趨勢圖、對照圖、流程圖、場景圖、價值矩陣。
```

## 9. 候選場站要放附錄

本次修正重點：

- 候選場站不要放在主簡報中間太早出現。
- 因為對方還未建立合作共識，太早談場站會讓簡報變成租地或設置提案。
- 候選場站只作為後續評估素材。

正確位置：

```text
附錄：首波候選場站方向
```

正確說法：

```text
候選場站僅作為後續評估素材，不作為本次簡報主軸。
以下為初步討論候選，不代表已確認可設置；
仍需以 Times 權責、供電與現勘結果為準。
```

## 10. PPTX 最後用截圖方式產出，避免跑版

使用者確認 HTML Final 後，才輸出 PPTX。

本次產出方式：

1. 用 Playwright + Edge 開啟 HTML。
2. 每一頁 `.slide` 截成高解析 PNG。
3. 用 `python-pptx` 建立 16:9 PPT。
4. 每張 PNG 滿版放入一頁投影片。

本次工具檔：

- `PARK24_Times_提案專案/build_ppt_from_html.py`

本次輸出：

- `PARK24_Times_提案專案/PARK24_Times_External_Final.pptx`
- `PARK24_Times_提案專案/ppt_export_slides/slide_*.png`

檢查結果：

- PPT 共 16 頁。
- 尺寸為 16:9。
- 每頁圖片為 `2560 x 1440`。
- 圖片滿版放入 PPT，不會因字型或文字框造成跑版。

## 11. 之後做 PPT 的標準流程

之後類似任務請照以下順序：

```text
1. 讀取使用者需求與來源資料
2. 判斷簡報用途：外部提案 / 內部報告 / SOP / 銷售簡報
3. 先產 Markdown 大綱
4. 檢查是否過度內部語氣
5. 重組成外部受眾會在意的敘事
6. 參考既有風格 PPT
7. 製作 HTML/CSS 預覽
8. 使用者確認內容與風格
9. 視覺補強：插圖、流程、表格、對照
10. 使用者確認 Final
11. HTML 截圖成 PNG
12. PNG 滿版組成 PPTX
13. 檢查頁數、16:9、圖片尺寸、是否跑版
```

## 12. 對外提案內容檢查表

每次輸出前檢查：

- 是否先回答「對方為什麼要做」？
- 是否先建立合作必要性？
- 是否避免太早談設備、場站、價格、SOP？
- 是否避免像內部研究或 BD 推案？
- 是否用對方視角描述價值？
- 是否有明確但不強迫的下一步？
- 是否使用共同決策語氣？
- 是否把候選場站放附錄？
- 是否風格符合既有公司簡報？
- 是否最後用截圖轉 PPTX 防跑版？
- ⚠️ **地區名稱是否正確？** 旺來試點城市一律是桃園 / 台中，不是台北。HTML 產出時特別容易殘留台北舊字樣，需逐段核對（見下方 Section 15）。

## 13. 本次最重要的心法（源自 PARK24 / Times 專案）

```text
PPT 不是把資料變成投影片。
PPT 是把對方需要相信的事情，依序變成可以接受的判斷。
```

對 PARK24 / Times 這類大型外部對象，簡報要先讓對方相信：

1. 這是停車場產業未來可能的方向。
2. PARK24 有條件驗證這件事。
3. 旺來有能力提出可驗證合作模式。
4. PARK24 可以取得場站、品牌、會員與學習價值。
5. 試點規模小、風險可控。
6. 是否擴大由雙方共同決定。

## 14. OnePager 與 Coldmail 工作流（源自城市車旅 / 嘟嘟房專案）

停車場開發案不一定一開始就進入 PPT。

若任務目標是開發窗口，建議先依照以下順序：

```text
One Pager 內容稿
→ One Pager DM HTML
→ 使用者確認 HTML
→ 必要時輸出 PNG / PDF
→ Coldmail 文書
→ 寄送與追蹤
→ 取得窗口或 30 分鐘會議
```

Coldmail 的定位：

- 不是完整提案。
- 不是產品型錄。
- 不是公司介紹。
- 目標是讓對方願意打開附件，並安排 30 分鐘會議。

每個品牌子專案若有 `02_Coldmail/`，應在其中建立可迭代的 cold mail 工作檔，例如：

```text
城市車旅_提案專案/02_Coldmail/CITY_Coldmail_Workflow.md
```

Coldmail 工作檔應包含：

- 本階段目標。
- 寄送附件。
- Email 結構。
- 主旨選項。
- 正式版信件。
- 精簡版信件。
- AI 產出 Prompt。
- 寄送前檢查表。
- 迭代紀錄。

Coldmail 內容原則：

- 先說「為什麼找這個品牌」。
- 再說合作構想。
- 明確排除租格位、廣告、傳統設備進駐的誤解。
- 說清楚旺來負責設備、客服、補貨、維運與異常處理。
- 對方協助事項要低負擔，例如場域評估、權責確認、現勘窗口。
- 不一開始要求大量點位。
- 先主張 `1 至 3 個示範點`。
- 收口在 `30 分鐘會議`。

---

## 15. OnePager DM 輸出品質規範（PNG / PDF）

> 源自嘟嘟房專案的卡點修正（2026-06-02）。

### 問題：Chrome headless CLI 不適合 OnePager 輸出

| 問題 | 原因 |
|------|------|
| PNG 含灰色背景 | `--screenshot` 截整個 viewport，包含 `body { background: #d9d9d9 }` |
| PDF 分成兩頁 | `--print-to-pdf` 用印刷引擎分頁，不保證 A4 剛好一頁 |

### 正確做法：Playwright Python

```python
from playwright.sync_api import sync_playwright

with sync_playwright() as pw:
    browser = pw.chromium.launch()

    # PNG：element.screenshot() 精確裁切 .page 元素，不含灰底
    ctx = browser.new_context(viewport={"width": 794, "height": 1123},
                              device_scale_factor=2)
    pg = ctx.new_page()
    pg.goto(url, wait_until="networkidle")
    pg.wait_for_timeout(800)
    pg.locator(".page").screenshot(path="output.png")
    ctx.close()

    # PDF：A4 單頁，彩色背景保留，不分頁
    ctx2 = browser.new_context(viewport={"width": 794, "height": 1123})
    pg2  = ctx2.new_page()
    pg2.goto(url, wait_until="networkidle")
    pg2.wait_for_timeout(800)
    pg2.pdf(path="output.pdf", format="A4", print_background=True,
            margin={"top":"0","right":"0","bottom":"0","left":"0"}, scale=1)
    ctx2.close()
    browser.close()
```

關鍵參數：
- `device_scale_factor=2`：輸出 2x Retina PNG（794×1123 → 1588×2246 px）
- `viewport 794×1123`：A4 at 96dpi，確保 PDF 恰好一頁
- `pg.locator(".page").screenshot()`：只截 `.page` div，不含 body 背景
- `print_background=True`：保留 CSS 背景色（橘色 bar、區塊底色）
- `margin={"top":"0"...}`：無邊距，A4 滿版

### 安裝前提

```powershell
pip install playwright
python -m playwright install chromium
```

注意：用 `python -m playwright install chromium`，不是 `playwright install chromium`（後者找不到 PATH）。

### 可重用腳本位置

```
嘟嘟房_提案專案/99_工作暫存/export_onepager.py
```

新品牌 OnePager 輸出時，修改腳本中的 HTML 路徑與 OUT 路徑即可重用。

---

## 16. BD 自動化工作流 Prompt 範本（新品牌必讀）

> 源自嘟嘟房專案（2026-06-02），為目前最完整版本。

### 範本位置

```
嘟嘟房_提案專案/99_工作暫存/DODO_BD_AUTOMATION_PROMPTS.md
```

此文件包含 Steps -1 至 8 的完整 AI Prompt，每一步都可直接複製給 Claude / Codex 執行。

### 新品牌使用方式

1. 複製 `DODO_BD_AUTOMATION_PROMPTS.md` 到新品牌的 `99_工作暫存/`
2. 把檔名改為 `{品牌}_BD_AUTOMATION_PROMPTS.md`
3. 全文搜尋「嘟嘟房」→ 替換為新品牌名稱
4. 全文搜尋「DODO」→ 替換為新品牌英文代稱
5. 確認各 Step 的輸出路徑與新品牌資料夾結構一致

### 工作流節點總覽

| Step | 任務 | 產出 |
|------|------|------|
| -1 | DD 調查建立與來源資料產出 | DD_REPORT / BD_STRATEGY / PROJECT_MASTER / HEADQUARTER_APPROACH |
| -0.5 | 案場資料庫建立（桃園 / 台中）| {品牌}_案場資料庫_v1.xlsx |
| 0 | 專案資料盤點 | {品牌}_Project_Readiness_Check.md |
| 1 | 專案定位分析 | {品牌}_BD_Positioning.md |
| 2 | OnePager 內容稿 | {品牌}_OnePager.md |
| 3 | OnePager DM HTML | {品牌}_OnePager_DM.html |
| 4 | PNG / PDF 輸出 | export/ 下的 PNG + PDF（用 Playwright，見 Section 15）|
| 5 | Coldmail 工作流與寄送文件 | Coldmail_Workflow / Email.txt / Package / Blueprint |
| 6 | 追蹤話術與取得會議文件 | Followup_Emails / Talk_Track |
| 7 | 簡版 PPT 大綱 | ShortPPT_Outline.md |
| 8 | 正式 PPT 大綱 | FormalPPT_Outline.md |

每節完成後停下來等使用者確認，不要自動進入下一步。

---

## 17. 地區名稱 QA：台北 vs 桃園 / 台中（重要）

> 源自嘟嘟房專案的高頻錯誤（2026-06-02）。

### 背景

旺來停車場 BD 試點城市一律是**桃園與台中**。台北不是試點目標（即使對象品牌的台北場站最多）。

### 為什麼容易出錯

1. DD 調查報告中「台北市 84 站」等數據多次出現，AI 在生成 OnePager 內容時容易跟著說「台北」
2. BD Positioning.md 可能已正確寫桃園/台中，但 OnePager HTML 在後續獨立生成時若沒有明確指示，仍可能殘留「台北」
3. Talk Track 和 FAQ 中「你們想在台北做嗎？」類的問答，必須確認回答是「旺來重心在桃園與台中」

### QA 規則

每份產出文件（OnePager.md、OnePager_DM.html、Coldmail、Talk_Track、PPT_Outline）在完成後，必須：

```text
搜尋「台北」，逐一確認每個出現位置：
- 若是描述對方品牌的現況數據 → 保留（例如「嘟嘟房台北市 84 站」）
- 若是描述旺來試點地區 → 必須改成桃園 / 台中
```

### 錯誤模式說明（抽象，不列實際文案）

此類錯誤的發生模式：

1. DD 報告中出現「[品牌] 在台北有 X 個場站」等數據
2. AI 生成 OnePager 內容時，誤把台北帶入試點地區描述
3. BD Positioning.md 可能已正確寫桃園/台中，但後續 HTML 生成是獨立執行的，若 prompt 沒有明確指定，容易殘留台北

**修正方式**：完成每份文件後，用文字搜尋「台北」，依 QA 規則逐一判斷保留或修改。

> ⚠️ 本段不列出任何品牌的實際文案，以避免跨專案引用。各品牌的修正記錄見各自專案資料夾的版本記錄。
