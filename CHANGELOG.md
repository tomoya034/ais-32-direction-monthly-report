# Changelog

本文件記錄 AIS 32 方位月報工具的重要版本變更。

格式參考 [Keep a Changelog](https://keepachangelog.com/zh-TW/1.1.0/)，版本號採 [Semantic Versioning](https://semver.org/) 概念管理。

## [Unreleased]

尚無未發布變更。

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

[Unreleased]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.4.0...HEAD
[1.4.0]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v1.3.0...v1.4.0
[1.3.0]: https://github.com/tomoya034/ais-32-direction-monthly-report/compare/v0.0.1...v1.3.0
[1.2.2]: https://github.com/tomoya034/ais-32-direction-monthly-report/commits/main
[0.0.1]: https://github.com/tomoya034/ais-32-direction-monthly-report/releases/tag/v0.0.1
