# Changelog

本文件記錄 AIS 32 方位月報工具的重要版本變更。

格式參考 [Keep a Changelog](https://keepachangelog.com/zh-TW/1.1.0/)，版本號採 [Semantic Versioning](https://semver.org/) 概念管理。

## [Unreleased]

尚無未發布變更。

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

[Unreleased]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.5.0...HEAD
[1.5.0]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.4.0...v1.5.0
[1.4.0]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.3.0...v1.4.0
[1.3.0]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v0.0.1...v1.3.0
[1.2.2]: https://github.com/tomoya034/ais-32-direction-monthly-report/commits/main
[0.0.1]: https://github.com/tomoya034/ais-32-direction-monthly-report/releases/tag/v0.0.1
