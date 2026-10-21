# Changelog

本文件記錄 AIS 32 方位月報工具的重要版本變更。

格式參考 [Keep a Changelog](https://keepachangelog.com/zh-TW/1.1.0/)，版本號採 [Semantic Versioning](https://semver.org/) 概念管理。

## [Unreleased]

尚無未發布變更。

## [1.5.6] - 2026-10-21

正式交付版；直接採用完整驗證的 hardening r2 binary，未為 release 重新 build。下列 local 版本段落保留當時狀態；最終驗證結論以本版與 Release Notes 為準。

- 正式發布tag為`v1.5.6-release1`，產品版本仍1.5.6；原`v1.5.6` tag保留，不重寫歷史。release1僅修兩組synthetic resume fixtures的canonical路徑與發布metadata測試契約，更新正式下載連結；production及validation演算法與已驗收EXE bytes均未改。

- 修正 CLI SourceFileError 逸出到 windowed bootloader 的互動錯誤對話框；統一 CLI／GUI 底層錯誤分類，CLI 保存結構化 FAIL log 並回傳非零。
- 新增逐列 Source Preflight：schema／欄位 mapping、GPS 辨識但不授權轉換、mixed-date、完整日曆與來源日、partial-day、fragment/core overlap、完整多重集合 subset 證據及 immutable SHA。
- 正式處理前阻擋不安全整月；不修改、拆分、去重或自動排除來源。預檢後及正式輸出前再驗來源身分。
- 正式整合 source-days／calendar-days validation，缺日 cache 不可虛構，ledger／helper／detail 不可出現來源資料。新增真實 28/29/30/31 天與缺日 corruption-negative fixtures。
- AUTO、90–100%、3-message、500 NM、Modern／Historical profiles、A/B 與32方向核心維持 frozen，不新增年度輸出功能。
- 修正 preflight numeric channel 的 XML int／float／text 身分邊界，避免漏判 normalized duplicate；r2 已重新完成全部必要回歸。
- 350 tests（0 failure/error/skip）；2026 334 份安全來源、303 logical days、10 港月首輪與重跑，以及2024六個安全月182份來源全部通過。26次港月、156工作簿、312對native geometry／decoded pixels均通過，數值差異0。
- 桌面部署後由正式EXE GUI完整製作HWLN 2026/04、2026/06，各六份成果；cache／selector／spool／工作簿semantic／native Excel與24對圖表pixels均與既有r2 baseline一致。mixed-date CLI自然exit 2與GUI可讀拒絕、取消及正常關閉通過。
- 保留既有選值研究限制；不宣稱AUTO已等同歷史人工判讀。正式套件不含任何AIS來源、reference或QA資料。

## [1.5.5-local] - 2026-10-07

本機候選；完整回歸與安裝驗證尚在進行，尚未建立公開 release。

- 修正 Modern 路徑與五份 Historical 成果路徑相撞時覆寫決策台帳，且成功回報卻只剩五份成果的缺陷。
- 加入跨程序資源鎖，保護共用輸出、cache 和 normalized spool；鎖寫入使用者暫存目錄，不寫入原始 AIS 資料夾。避免同時執行兩個 GUI／CLI 時固定 `.building`／`.tmp` 路徑互相干擾。
- 首次分析的六份成果採同一組建置／提交／rollback；Historical 中途失敗不再留下新 Modern 搭配舊 Historical。
- 覆核五檔與首次六檔共用提交與回復機制；未授權覆寫時使用原子拒絕既有目標操作，保留建置期間晚出現的同名檔。
- 保護實際覆核輸入副本，拒絕其與 Historical／temporary／rollback／cache 路徑碰撞。GUI 來源 preflight 失敗不再留下停用的按鈕；首次輸出不預設授權覆寫。
- source/copy 驗證器補強 manifest／檔案身分、checkpoint 原子寫入、失敗續跑、讀取期間變動檢查及本次 hash／沿用結果計數。
- 全量 runner 保存完整計畫與中斷狀態；validator 不再把空報告或已完成的部分月份判為全量 PASS。獨立稽核保護原始／reference／cache／spool 位置，拒絕被改動或範圍不符的續跑紀錄。
- 補充獨立來源／normalized row 全量稽核，並修正原報告對視覺等同、人工 reference、原重複輸出事件及正式可用性的過度聲稱。
- Native 全系列/point 樣式盤點重現 Historical 標題16pt與人工5份reference的14pt不同；共用renderer改14pt。不同full/small角色的Chart Area geometry差異仍待修復與全量重验，不宣稱視覺完全相同。

## [1.5.4-local] - 2026-10-05

- 修正 HWLN 2026-04-01 至 04-09 的 `YYYYMMd` 檔名未被辨識、流程卻仍回報成功的 silent omission；來源掃描現在支援明確的 7／8 位日期並對無法解析的 AIS-like 檔名 fail closed。
- Historical 圖表依六月正確成果重建：filled radar、30 色 Office palette、60% transparency、reference Chart/Plot Area、尺寸／anchor、KLNG 21 方位視覺順序，以及青色 MAX、橙色 MIN、紅色 2.25 pt AVG line。
- 保留 floor-bin 32 方位與港別 policy：KLNG `klng21`，HWLN／其他站點 `all32`；沒有把花蓮問題修成港別硬編碼。
- 新增全量來源 catalog、月份批次 runner、reference 比對、OOXML／內容驗證、Excel-native chart export 與 deterministic package hash 工具。
- 保留 v1.5.3 exact-overlap fail-closed 行為。NAS 全量 335 檔中的 KLNG 2026-07-24 `_23`／`_24` 會如預期觸發 audit-truncated gate；既有逐列稽核已證明 `_23` 的 708,418 筆正規化資料有 99.999435% 包含於 `_24`，因此正式安全輸入依既有 manifest 排除被取代的 `_23`，保留 `_24`，不做 silent dedupe。

## [1.5.3-local] - 2026-09-18

本機修復與驗證版本；尚未建立 git tag、GitHub Release 或公開下載資產。

### Fixed

- 修正 Period A/B 大 workbook 在 `constant_memory` 模式下回寫較早列失效，導致「總表」AI:AL helper cells 與第二張歷史統計圖空白；維持逐列串流與大檔低記憶體特性，並以完整連續索引寫入 chart cache，避免 Excel 16.0 丟棄 combo chart series。
- 修正 Modern、legacy 與 Historical 總表的空白引用／空集合統計公式；Excel 原生重算後缺值仍為空白，不會被 `IFERROR(blank_reference,"")` 或 `MAX/MIN(empty_range)` 轉成 0。
- Modern、Historical Period A、Historical Period B 與 A/B 整合圖表標題加入 scope，避免把不同訊息 profile／period 定義的圖誤認為同一結果。

### Changed

- 非 byte-identical fragments 若含 exact normalized core duplicates，或 cross-fragment audit 超過 250,000 筆而無法證明安全，會停止本次月份處理且不產生／替換正式檔；不做 silent dedupe。只有互補 fragments、無 exact duplicate 的可完整稽核 overlap，以及 byte-identical 單次處理可繼續。
- `CACHE_VERSION` 升為 7，強制舊 cache 6 失效並重掃原始 fragments，避免既有 cache 繞過新的 overlap safety gate。Normalized spool version 1 與 review schema `AIS_V15_REVIEW_3` 不變。

## [1.5.2] - 2026-09-15

### Fixed

- 修正多港別共用 KLNG 21 方位排除清單。加入可追溯的 `auto`／`klng21`／`all32` 設定；auto 對 KLNG 維持原規則，其他站點不套用 KLNG 地理排除。
- 總表、覆核台帳、每日明細、Period A/B、整合及交付驗證使用同一方位設定，缺資料仍留白。
- 雷達圖加入 `auto`／`line`／`filled` 和獨立的刻度上限；auto 對 KLNG 保留線形、其他站點採填色；不容許指定上限裁掉有效數值。
- 方位設定納入快取及覆核身分；cache 6、review schema `AIS_V15_REVIEW_3`，舊覆核須重新分析，防止以舊 HWLN 21 方位台帳覆蓋 32 方位。
- 修正 XLSX 的 MAXIFS、STDEV.S 公式序列化，加入必要的 `_xlfn.` 前綴，避免 Microsoft Excel 重新計算出現 `#NAME?` 後被總表 IFERROR 隱藏成空白。公式的計算範圍及數學意義不變。
- 填色雷達省略全無有效值的日期系列，月份表格與缺日紀錄仍保留；座標刻度使用整數 NM，避免長刻度文字與方位標籤重疊。
- 未改動 AIS 訊息規則、bearing 分箱、群聚選值、Period 定義或 KLNG 既有數值契約。`all32` 是保留資料的工程設定，不宣稱各站全部方向都是海域。

## [1.5.1] - 2026-09-07

### Fixed

- 32 方位中文名稱依中文維基百科〈羅盤方位〉「中文名」欄修正十個 label，包含 SEbE = 東南微東、SEbS = 東南微南。所有中文顯示與 coastal review 名稱同步更新；英文縮寫不變。
- 舊中文 direction keys 的 cache version 由 4 升為 5。Review schema 升為 `AIS_V15_REVIEW_2`，明確拒絕 v1.5.0 覆核 workbook，要求由原始月份資料重新分析以避免 decision direction 錯置；不自動 migration。

### Tests

- 固定完整 index／abbreviation／中文名稱 mapping，涵蓋 32 方位、SEbE／SEbS、WSW coastal review、floor-bin、cache 失效與舊 review schema 拒絕。
- 六月 numeric regression 以固定 index／abbreviation 比較，避免歷史中文 label 造成假 regression。

### Unchanged

- 11.25° floor-bin、32 index、21 海向、500 NM、10%／3 筆 selector、Modern／Historical profile、Period A/B 及 decision value contract 均不變。
- Numeric binary normalized spool version 維持 1；五份正式成果的 detail MAX、small=big、integrated MAX 契約不變。

## [1.5.0] - 2026-09-07

### Added

- 同港同日支援任意數量來源 fragments；檔名 suffix 改為純追溯資訊，不再以 mtime 選一檔而遺失同日資料。
- 加入 profile-neutral logical-day normalized spool，保存 msg_type、時間、MMSI、channel、fragment、sheet、source row 及超過 selection cap 的有效 East rows。
- Modern workbook 新增 `logical_day`、`period_a`、`period_b` 候選與唯一 decision ledger；Period A/B 僅接受真實 Candidate ID 或留白。
- 首次分析預設產生 Modern workbook 與五份正式 Historical delivery；覆核後可從 GUI 或 `--finalize-from` 驗證並整組重生五檔。
- 加入 time-range、occupied-second、正常同秒訊息、possible duplicate event、normalized-core 相同紀錄及 byte-identical file 的分層診斷。
- 加入六月外部 Golden contract gate、代表日 full-search checksum regression、五檔公式／圖表／數值契約與故障復原測試。

### Changed

- Modern logical-day 固定先 union 當日全部 fragments 再選值；Historical delivery 依 row timestamp 切成 Period A（00:00–11:59:59）與 Period B（12:00–23:59:59）。
- `top_candidates` 只影響 workbook 顯示，selector 一律完整搜尋；輸出 renderer 不再改變研究值。
- Modern 預設 profile 保持 `{1,2,3,18,19}`，Historical delivery 固定 `{1,3,4,18,19}`，兩者設定與語意互相隔離。
- 五份成果均由同一 decision snapshot 建立；統計與圖表由本次 grid 重算，整合表嚴格使用 `MAX(Period A, Period B)`。
- 每日 worker 改以 logical day 計數；CLI 新增 `--max-days` 與 `--delivery-dir`，`--max-files`、`--legacy-output` 僅作 deprecated alias。
- 經六月 6/1、6/5、6/20 benchmark 後採 simple per-day compact in-memory sorting；external runs 延後到實際單日資料超出目前約 288 MiB peak RSS 基準時再評估。

### Fixed

- 修正 v1.4.0 對同港同日多檔只保留 mtime winner，造成六月 `_11` 30 份資料全部未處理的問題。
- 修正 modern-only 只搜尋顯示候選、啟用舊格式才完整搜尋所造成的輸出選項影響研究值。
- 正式五檔改為先完整建立再整組替換；失敗會復原舊 snapshot，不留下新舊混用成果。
- Excel 明細超過 1,048,575 筆時於發布前停止，絕不靜默截斷。

### Contract

- `Period detail MAX = Period final`。
- `Period 小總表 = 對應大 workbook 總表`。
- `Integrated = MAX(Period A final, Period B final)`。
- 非 byte-identical fragments 在沒有領域 dedupe key 前只診斷、不自動去重。

## [1.4.0] - 2026-09-03

### Added

- 支援從 `D&TMOK <PORT>_YYYYMMDD_*.xlsx` 自動辨識港別，港別代碼不限制為預先列出的清單。
- 圖形介面加入港別下拉選單；切換港別時只顯示該港別可用月份。
- CLI 新增可選的 `--port`；唯一港別仍可沿用既有免指定方式，多港資料夾則明確要求選擇。
- Multi-port 測試涵蓋 KLNG、HWLN、港別/月分組、CLI 歧義、輸出命名與研究數值回歸。

### Changed

- 新版與原格式相容版的檔名、工作簿標題、處理紀錄、log 與錯誤報告改用實際港別代碼。
- cache 與 legacy spool 以港別、年份及月份隔離；舊格式快取會自動失效。
- 版本提升為 `1.4.0`；研究演算法、32 方位、21 海向、500 NM、10%／至少 3 筆群聚、人工覆核與 AIS 去重規則均未變更。

### Fixed

- 防止同一來源資料夾內不同港別的每日資料被合併進同一份月報。

## [1.3.0] - 2026-08-26

### Added

- 影片中的逐日篩選、異常高值排除、32 方位選值、逐日表與總表流程正式列為一鍵全自動作業。
- 圖形介面加入由固定檔名建立的唯讀月份清單；多月份資料夾會預選最新月份並允許安全改選。

### Changed

- 年份、月份改以 `D&TMOK KLNG_YYYYMMDD_*.xlsx` 檔名為唯一依據，開始前會重新掃描，不再接受自由輸入。
- 命令列只需 `--input` 與 `--output` 即可從檔名推定月份；舊有 `--year`、`--month` 仍可成對使用。

### Fixed

- 避免切換每月來源資料夾後忘記同步修改月份，造成讀不到檔案或產出錯誤月份名稱。

## [1.2.2] - 2026-07-28

### Changed

- CLI 預設平行工作數會依電腦自動選擇 1–3，不再固定為 1。
- 補上上述情境的回歸測試與 Python 3.13、3.14 套件中繼資料。

### Fixed

- 可攜版預設輸出資料夾改為 EXE 所在位置旁的 `output\AIS月報`，不再指向 PyInstaller 暫存目錄。
- 來源活頁簿第一頁若為封面或說明，會自動尋找 `AIS` 或含必要欄位的資料工作表。

## [0.0.1] - 2026-07-19

### Added

- 一頁式 Windows 圖形介面。
- 自動偵測來源年月與月份天數。
- 完整數值化的 32 方位篩選及群聚選值。
- 同次執行產生新版自動分析與原格式相容版。
- 每日斷點快取及原格式二進位暫存。
- 中文錯誤訊息與詳細錯誤報告。
- Excel 圖表、待複核清單與人工覆核欄。
- PyInstaller 單一 EXE 建置腳本及自動測試。

[Unreleased]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.5.6-release1...HEAD
[1.5.6]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.5.2...v1.5.6-release1
[1.5.2]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.5.1...v1.5.2
[1.5.1]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.5.0...v1.5.1
[1.5.0]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.4.0...v1.5.0
[1.4.0]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.3.0...v1.4.0
[1.3.0]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v0.0.1...v1.3.0
[1.2.2]: https://github.com/tomoya034/ais-32-direction-monthly-report/commits/main
[0.0.1]: https://github.com/tomoya034/ais-32-direction-monthly-report/releases/tag/v0.0.1
