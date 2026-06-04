# 停車場品牌合作開發計畫 Rehydrate Report

> 依據：`PARK24_Times_提案專案/99_工作暫存/00_PROJECT_BOOTSTRAP.md.txt`  
> 目的：接手既有專案，不重做已完成內容，先盤點現況、找出缺口，並以「最短路徑取得品牌總部簡報機會」為最高決策原則。  
> 產出日期：2026-06-03

## 1. Project Inventory

### 已完成

#### 共用層

- `README.md`
  - 已建立專案使用方式、PPT 製作原則、旺來風格說明。
- `PROJECT_INDEX.md`
  - 已建立 PPT 工作流、OnePager / Coldmail 工作流、OnePager 輸出品質規範、BD 自動化工作流範本、地區 QA 規則。
- `AI_PPT_GUIDELINE.md`
  - 已建立通用 AI PPT guideline。
- `One Pager 產出規範/OnePager_Production_Standard.md.txt`
  - 已建立停車場品牌合作 OnePager 標準。
- `shared_assets/`
  - 已有旺來通用 Logo / slogan 素材。
- `wagas_style_reference/`、`風格/`
  - 已有旺來既有簡報風格參考。

#### PARK24 / Times

- `00_來源資料/`
  - `PARK24_Headquarter_Proposal.docx`
  - `PARK24_BD_Strategy.docx`
  - `PARK24_DD_Report.docx`
  - `PARK24_DD_Execution_Guide_v1.md.txt`
  - `PARK24_Project_Master.xlsx`
- `PARK24_PPT_GENERATION_PROMPT.md`
- `PARK24_PPT_OUTLINE.md`
- `PARK24_PPT_PREVIEW_EXTERNAL.html`
- `PARK24_Times_External.pptx`
- `Times_合作報告.pptx`
- `ppt_export_slides/slide_01.png` 至 `slide_16.png`
- `build_ppt_from_html.py`

判斷：PARK24 已完成正式外部提案 PPT 級別產物。

#### 城市車旅

- `00_來源資料/`
  - `CITY_PARKING_DD_EXECUTION_GUIDE.txt`
  - `CITY_PARKING_DD_REPORT.docx`
  - `CITY_PARKING_BD_STRATEGY.docx`
  - `CITY_PARKING_HEADQUARTER_APPROACH.docx`
  - `CITY_PARKING_DD_MASTER.xlsx`
  - `CITY_案場資料庫_v1.xlsx`
  - 其他合作建議與供電評估資料
- `01_OnePager/`
  - `CITY_OnePager.md`
  - `CITY_OnePager_DM.html`
  - `export/CITY_OnePager_DM.png`
  - `export/CITY_OnePager_DM.pdf`
- `02_Coldmail/`
  - `CITY_Coldmail_Workflow.md`
  - `CITY_Coldmail_Email.txt`
  - `CITY_Coldmail_Package.md`
  - `CITY_BD_Framework.md`
  - `CITY_BD_Blueprint.md`

判斷：城市車旅已完成取得窗口前的 OnePager + Coldmail 包。

#### 嘟嘟房

- `00_來源資料/`
  - `DODO_DD_EXECUTION_GUIDE.md`
  - `DODO_DD_REPORT.md`
  - `DODO_BD_STRATEGY.md`
  - `DODO_PROJECT_MASTER.md`
  - `DODO_HEADQUARTER_APPROACH.md`
  - `DODO_案場資料庫_v1.xlsx`
- `99_工作暫存/`
  - `DODO_BD_AUTOMATION_PROMPTS.md`
  - `DODO_Project_Readiness_Check.md`
  - `DODO_BD_Positioning.md`
  - `BD自動化工作流_向上報告.md`
  - `BD自動化工作流_向上報告.docx`
  - `export_onepager.py`
- `01_OnePager/`
  - `DODO_OnePager.md`
  - `DODO_OnePager_DM.html`
  - `export/DODO_OnePager_DM.png`
  - `export/DODO_OnePager_DM.pdf`
- `02_Coldmail/`
  - `DODO_Coldmail_Workflow.md`
  - `DODO_Coldmail_Email.txt`
  - `DODO_Coldmail_Package.md`
  - `DODO_BD_Blueprint.md`
- `03_取得窗口與會議/`
  - `DODO_Followup_Emails.md`
  - `DODO_Talk_Track.md`
- `04_簡版PPT/`
  - `DODO_ShortPPT_Outline.md`
- `05_正式PPT/`
  - `DODO_FormalPPT_Outline.md`

判斷：嘟嘟房已完成從 DD 到正式 PPT 大綱的完整前置鏈，但尚未完成正式 PPT HTML / PPTX。

### 部分完成

- PARK24 的 `01_OnePager`、`02_Coldmail`、`03_取得窗口與會議` 資料夾為空。
  - 但 PARK24 已完成正式提案 PPT，因此不是內容不足，而是 outreach package 未結構化。
- 城市車旅已完成 OnePager / Coldmail，但 `03_取得窗口與會議`、`04_簡版PPT`、`05_正式PPT` 尚未產出。
- 嘟嘟房已有簡版 / 正式 PPT 大綱，但尚未製作 HTML 簡報與 PPTX。
- ViVi PARK、俥亭、台灣聯通目前僅有標準資料夾，尚未進入此輪三品牌主線。

### 尚未完成

以「取得品牌總部簡報機會」為目標，目前缺口如下：

- PARK24
  - 缺少可寄送的 meeting request / coldmail package。
  - 缺少會議邀約話術與追蹤信。
  - 缺少明確的窗口追蹤紀錄。
- 城市車旅
  - 缺少寄送後追蹤話術。
  - 缺少簡版 PPT 大綱與正式 PPT 大綱。
  - 缺少窗口名單與寄送紀錄。
- 嘟嘟房
  - 缺少窗口名單與寄送紀錄。
  - 缺少正式 PPT HTML / PPTX。
  - 缺少正式簡報前的視覺化版本。

### 重複內容

- `PROJECT_INDEX.md` 與 `DODO_BD_AUTOMATION_PROMPTS.md`
  - 前者是專案級規則與記憶，後者是嘟嘟房執行 Prompt。
  - 功能相近但層級不同，不建議合併。
- `DODO_BD_Positioning.md` 與 `DODO_BD_STRATEGY.md`
  - 一個是 AI 產出的 BD 定位分析，一個是來源資料策略稿。
  - 內容可能重疊，但用途不同，不建議刪除。
- `PARK24_Times_External.pptx` 與 `Times_合作報告.pptx`
  - 應確認哪一份是最終寄送版本。
  - 暫不刪除，避免誤刪使用者手動版本。
- 城市車旅 `CITY_Coldmail_Workflow.md` 與 `CITY_Coldmail_Package.md`
  - Workflow 是規則，Package 是寄送版本，保留。

## 2. Project Progress Report

### PARK24 / Times

目前進度：

- DD / BD / Project Master 已完成。
- 外部正式提案 HTML 與 PPTX 已完成。
- 已具備總部簡報內容。
- 缺的是「寄送取得會議」前的開發文書與窗口追蹤。

完成率：

```text
內容準備：90%
取得會議流程：55%
整體推進至總部簡報機會：75%
```

下一步：

- 補 PARK24 meeting request package。
- 確認 final PPT 檔案版本。
- 建立 30 分鐘會議邀約信與追蹤話術。

### 城市車旅

目前進度：

- 來源資料、DD、BD、案場資料庫已完成。
- OnePager DM 已完成並已輸出 PNG / PDF。
- Coldmail package 已完成。
- 尚未產出追蹤話術與 PPT 大綱。

完成率：

```text
內容準備：70%
取得會議流程：65%
整體推進至總部簡報機會：60%
```

下一步：

- 產出追蹤信、電話話術、會議邀約腳本。
- 根據寄信回覆狀況，再製作簡版 PPT 大綱。

### 嘟嘟房

目前進度：

- DD、BD、案場資料庫已完成。
- OnePager / DM / PNG / PDF 已完成。
- Coldmail package 已完成。
- 追蹤信與電話話術已完成。
- 簡版與正式 PPT 大綱已完成。
- 尚未製作 PPT HTML / PPTX。

完成率：

```text
內容準備：85%
取得會議流程：80%
整體推進至總部簡報機會：80%
```

下一步：

- 先執行寄送與追蹤，而不是急著做正式 PPT。
- 若已取得會議，再把簡版 PPT 大綱轉成 HTML 預覽。

## 3. Brand Funnel Stage

| 品牌 | 目前 Stage | 判斷 |
|------|------------|------|
| PARK24 / Times | Stage 8：總部提案 | 正式 PPT 已完成，但會議取得流程文件不足 |
| 城市車旅 | Stage 5：Cold Email | OnePager + Coldmail 完成，尚未建立追蹤與會議腳本 |
| 嘟嘟房 | Stage 7：簡報邀約 | Coldmail、追蹤話術、簡版 PPT 大綱皆已完成，可推進會議邀約 |

漏斗定義：

```text
Stage 1 品牌研究
Stage 2 DD調查
Stage 3 案場研究
Stage 4 One Pager
Stage 5 Cold Email
Stage 6 窗口確認
Stage 7 簡報邀約
Stage 8 總部提案
Stage 9 試點規劃
Stage 10 簽約導入
```

## 4. Gap Analysis：距離「取得總部簡報機會」還缺什麼

### 高優先

- 建立 / 確認各品牌的實際寄送窗口。
- 建立寄送紀錄表：寄送日期、窗口、主旨、附件、回覆狀態、下一次追蹤日期。
- PARK24 補齊 meeting request package，因為它已經有正式 PPT，最不該卡在寄信文件。
- 嘟嘟房進行窗口確認與寄送，不要繼續堆文件。

### 中優先

- 城市車旅補 `03_取得窗口與會議` 文件。
- 城市車旅補簡版 PPT 大綱，作為取得會議後的對焦素材。
- 嘟嘟房簡版 PPT 大綱轉 HTML 預覽，但應等會議機會明確後再做。

### 低優先

- PARK24 OnePager 補檔。
  - 因正式 PPT 已完成，OnePager 不是最短路徑必要項。
- 嘟嘟房正式 PPT HTML / PPTX。
  - 現階段還未取得會議，正式 PPT 不急。
- ViVi PARK、俥亭、台灣聯通啟動 DD。
  - 不應分散目前三品牌推進資源。

## 5. Action Plan：只列未完成項目

| 任務 | 優先級 | 依賴文件 | 預估工時 | 輸出檔案 | 完成條件 |
|------|--------|----------|----------|----------|----------|
| 建立三品牌 BD 寄送追蹤表 | 高 | 現有 Coldmail / PPT package | 0.5 hr | `BD_OUTREACH_TRACKER.md` | 每品牌有窗口、寄送狀態、下一步欄位 |
| 補 PARK24 meeting request package | 高 | `PARK24_PPT_PREVIEW_EXTERNAL.html`、`PARK24_Times_External.pptx` | 1 hr | `PARK24_Meeting_Request_Email.txt`、`PARK24_Meeting_Request_Package.md` | 可直接寄送邀約總部簡報 |
| 確認 PARK24 Final PPT 檔案版本 | 高 | `PARK24_Times_External.pptx`、`Times_合作報告.pptx` | 0.25 hr | 記錄於 tracker | 明確知道寄送哪一份 |
| 嘟嘟房窗口確認與寄送準備 | 高 | `DODO_Coldmail_Email.txt`、`DODO_OnePager_DM.pdf` | 0.75 hr | `DODO_Send_Checklist.md` | 具備寄送前 checklist 與附件確認 |
| 城市車旅追蹤話術 | 中 | `CITY_Coldmail_Email.txt`、`CITY_OnePager_DM.pdf` | 0.75 hr | `CITY_Followup_Email.txt`、`CITY_Call_Talktrack.md` | 可追蹤收件並邀約會議 |
| 城市車旅簡版 PPT 大綱 | 中 | `CITY_OnePager.md`、`CITY_BD_Framework.md`、來源資料 | 1.5 hr | `CITY_ShortPPT_Outline.md` | 取得會議後可快速轉簡報 |
| 嘟嘟房簡版 PPT HTML 預覽 | 低 | `DODO_ShortPPT_Outline.md` | 2 hr | `DODO_ShortPPT_PREVIEW.html` | 會議明確後可視化 |

## 6. Priority Brand Ranking

### 1. 嘟嘟房

理由：

- 已完成 DD、案場資料庫、OnePager、DM、Coldmail、追蹤話術、簡版/正式 PPT 大綱。
- 距離「寄送 → 追蹤 → 取得會議」最近。
- 目前不需要更多內容製作，應轉入 BD 執行。

### 2. PARK24 / Times

理由：

- 正式 PPT 已完成，若取得窗口即可直接進入總部提案。
- 但缺少 meeting request package 與寄送追蹤，因此短路徑是補齊邀約文件。

### 3. 城市車旅

理由：

- OnePager + Coldmail 已完成，可寄送。
- 但相比嘟嘟房，追蹤話術與簡版 PPT 尚未完成，取得會議後的承接資料較弱。

## 7. Recommended Immediate Execution

依最短路徑，下一個執行動作建議為：

```text
建立 BD_OUTREACH_TRACKER.md，
統一管理 PARK24、城市車旅、嘟嘟房的寄送狀態、窗口、附件與下一步。
```

完成 tracker 後，優先處理：

```text
1. 嘟嘟房寄送準備
2. PARK24 meeting request package
3. 城市車旅追蹤話術
```

