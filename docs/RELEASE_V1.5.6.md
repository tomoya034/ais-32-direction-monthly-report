# AIS 32 方位月報工具 v1.5.6

來源預檢與錯誤處理強化版：直接交付已完成完整回歸的 r2 EXE，沒有重新 build，也沒有改變研究選值規則。

## 📥 下載

Windows 請下載 `AIS_32_Direction_Monthly_Report_v1.5.6.zip`，解壓縮後執行 `AIS_32方位月報工具.exe`，不需要 Python。核對套件及 EXE 的 `SHA256SUMS.txt`；`Source code` 不是 Windows 執行版。

## ✨ 本版重點

- 新增逐列 Source Preflight：以名稱辨識 schema／欄位，列出來源日、完整日曆、缺日、部分日、fragment overlap／subset 與來源身分。
- mixed-date 在數值處理前阻擋整月；不改日期、不拆檔、不默默去重或排除來源。
- GPS_* schema 明確偵測並阻擋，暫不自動轉換；缺欄、未知／歧義 schema 亦阻擋。
- 完整重複／normalized overlap／subset 保留證據。SUBSET_CANDIDATE 不是自動排除授權，須另有明確 manifest 的安全來源 view。
- 修正 source days／calendar days 驗證；閏年缺日可保留空白日曆列，但不虛構 cache、ledger、detail 或 helper 來源。
- 改善 CLI SourceFileError：結構化持久錯誤紀錄、可用時寫 stderr，來源／設定錯誤自然 exit 2，不再交給 windowed bootloader 顯示互動錯誤；GUI 保留可理解訊息。
- 整合先前本機候選的日期辨識、整組六檔安全提交／rollback、資源鎖、空值與原生圖表修復。
- AUTO selector、90–100% cluster、至少3筆、500 NM、Modern／Historical profiles、A/B、11.25° floor-bin與港別方向政策未因本輪 hardening／release 更改。

## ✅ 驗證

- 完整本機 suite：350 tests，0 failure、0 error、0 skip（含外部契約 fixtures）。
- 2026：334份安全來源、303 logical days、10港月，完整首輪及重跑，120 XLSX。
- 2024安全月：1、2、5、8、9、12月，182份來源，36 XLSX。另完成全年336份raw來源預檢及正式負向測試；不是宣稱全部raw可安全處理。
- hardening合計26次港月、156工作簿、312對原生geometry／decoded RGBA pixels，數值差異0。
- Windows部署EXE／GUI：HWLN 2026/04及2026/06完整正常流程，各1 Modern＋5 Historical；共12工作簿、24對原生圖表，與既有r2 baseline數值、selector／cache／spool、全部工作簿semantic、公式、blank／0、source／chart ranges、series／missing handling及pixels均一致。
- 部署負向smoke：真實mixed-date由CLI自然exit 2及GUI明確拒絕，零正式XLSX／cache；來源選取、取消、正常關閉通過。
- GitHub Actions與本機外部fixture驗證分開記錄，遠端不附私人AIS／Golden；實際CI結論見本版本commit的Tests run，不以CI skips代替外部資料驗證。

## ⚠️ 已知限制

- AUTO遵循既定規則已獲工程驗證，但尚未證明等同歷史人工／研究判讀。
- GPS時間欄語義仍UNVERIFIED，因此沒有adapter；混日／subset不自動修復。
- 缺日／partial-day警告不是完整接收保證；無來源日保持空白。
- 來源預檢與身分核對增加掃描時間；完整證據保留在工程evidence root，不隨使用者套件發布。
- 沒有年度總表新功能。未做商業程式碼簽章，請核對官方來源與SHA。

## ⬆️ 從舊版升級

先正常關閉舊工具，保留EXE、說明與既有成果，再使用新版。桌面舊v1.5.4-local已有可回復備份；回復方式見套件的「部署與回復說明.txt」。不同版本重製成果請選新輸出位置，不要刪改原始AIS或直接修改五份Historical。cache version 7、normalized spool version 1、review schema AIS_V15_REVIEW_3維持原契約；更舊cache會失效重掃，不自動轉換舊覆核。

## 📋 完整變更

見 `CHANGELOG.md`；binary／source freeze身分見 `docs/releases/v1.5.6.json`。此版使用已驗證candidate原bytes，正式文件更新不表示重新編譯。
