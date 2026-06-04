# 旺來 PPT 製作專案

這個資料夾用來製作旺來商務簡報，尤其是對外合作提案型 PPT。

本專案目前已完成一份 PARK24 / Times 外部提案簡報，並沉澱出一套可重用的 PPT 製作流程。

## 使用方式

之後只要在這個專案中製作新的 PPT，請先閱讀：

1. `PROJECT_INDEX.md`
   - 本專案最重要的操作索引。
   - 記錄這次 PARK24 / Times 簡報的成功流程。
   - 新的外部提案簡報應優先依照此流程製作。

2. `AI_PPT_GUIDELINE.md`
   - 通用 AI PPT 製作 guideline。
   - 用於建立需求包、大綱、視覺系統、HTML 預覽與輸出流程。

3. `PARKING_SITE_SELECTION_PLAYBOOK.md`
   - 停車場案場選點、TOP10 初選、桃園 / 台中場站評估與供電分析的專用工作法。
   - 只要任務涉及案場資料庫、候選場站、現勘清單或試點選址，就必須先閱讀。
   - 重點是避免只看官網表面資料，並把官方事實、AI 推論與待現勘項目分清楚。

4. 當次任務的專案素材
   - 依使用者指定，讀取本次簡報相關的需求文件、大綱、研究資料、表格、品牌素材或參考簡報。
   - 可能是 `*_OUTLINE.md`、`*_PROMPT.md`、Word、Excel、PDF、TXT、HTML，或各子專案來源資料夾、`風格/`、`shared_assets/`、`wagas_style_reference/` 中的檔案。
   - 不要假設每次都有固定檔名；應先搜尋工作區，再依任務需求判斷該讀哪些檔案。

## 核心製作原則

- 不要一開始直接製作 PPTX。
- 先讀來源資料，整理 Markdown 大綱。
- 先判斷簡報用途：外部提案、內部報告、SOP、銷售簡報或會議簡報。
- 外部提案要先回答「對方為什麼值得做」，不要直接講「我們想做什麼」。
- 避免把內部 BD、DD、執行計畫、風險登錄直接放進主簡報。
- 先製作 HTML/CSS 預覽版。
- 使用者確認 Final 後，再用高解析截圖方式轉成 PPTX，避免跑版。
- 若任務涉及停車場選點，先讀 `PARKING_SITE_SELECTION_PLAYBOOK.md`，再建立或修改案場資料庫。

## 旺來簡報風格

本專案使用的旺來風格特徵：

- 白底
- 左上大黑標
- 右上 waGas 橘色 Logo
- 橘色副標
- 灰色分隔線
- 底部品牌精神與頁碼
- 橘色幾何封面
- 條列、流程箭頭、表格、對照圖

不要任意改成深色 SaaS 風、卡片式儀表板風或過度科技感版面，除非使用者明確要求。

## 共用素材與風格參考

- `shared_assets/`
  - 旺來通用品牌素材，例如 waGas Logo 與品牌精神圖。
  - 新的子專案 HTML 若需引用品牌素材，應使用相對路徑指向此資料夾。

- `wagas_style_reference/`
  - 從既有旺來 PPT 萃取出的風格參考圖片與頁面截圖。
  - 用於判斷版面語氣、橘色幾何封面、標題配置、流程箭頭、表格與品牌視覺。

- `風格/`
  - 原始參考 PPT 檔。
  - 若需重新萃取或比較完整簡報風格，優先從此資料夾讀取。

## 已完成成果

### Final PPT

- `PARK24_Times_提案專案/PARK24_Times_External_Final.pptx`

### Final HTML 預覽

- `PARK24_Times_提案專案/PARK24_PPT_PREVIEW_EXTERNAL.html`

### PPT 逐頁高解析圖片

- `PARK24_Times_提案專案/ppt_export_slides/`

### HTML 轉 PPTX 腳本

- `PARK24_Times_提案專案/build_ppt_from_html.py`

## PARK24 / Times 簡報定位

這份簡報不是：

- 設備提案
- 租地提案
- 內部研究報告
- SOP
- 立即大量展點提案

這份簡報定位為：

```text
PARK24 x 旺來瓦斯 社區型智慧服務節點合作提案
```

目標是：

```text
取得總部同意進入 1 至 3 個試點場站評估。
```

## 下次製作 PPT 的推薦指令

可以直接這樣要求 Codex：

```text
請先閱讀 README.md、PROJECT_INDEX.md 與 AI_PPT_GUIDELINE.md，
再依照本專案的 PPT 製作流程，協助我製作新的外部提案簡報。

請先產出 Markdown 大綱，不要直接製作 PPTX。
確認大綱後，再製作 HTML/CSS 預覽。
使用者確認 Final 後，再用截圖方式轉成 PPTX，避免跑版。
```

若任務是停車場選點或案場資料庫，可以改用：

```text
請先閱讀 README.md、PROJECT_INDEX.md 與 PARKING_SITE_SELECTION_PLAYBOOK.md，
再依照本專案的案場選點流程，協助我建立或修正本品牌的案場資料庫。

請先確認官方案場來源是否完整，檢查官網是否有動態資料、script 或 API，
再分清楚官方事實、AI 推論、待現勘確認與不建議優先項目。
```
