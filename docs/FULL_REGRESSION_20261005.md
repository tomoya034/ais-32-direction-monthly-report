# AIS 32 方位結案後全量回歸紀錄（2026-10-05）

## 結論狀態

本文件記錄 v1.5.4-local 對 NAS 既有資料的盤點、修復與重現方法。335 份原始檔已先以原樣輸入完成安全門測試，再以有逐列證據的 334 檔安全 view 完成兩輪 10/10 全量回歸；60 份 workbook 的內部內容／圖表契約、116,244,596 筆 retained detail rows、Excel-native 60/60 與 packaged EXE 兩港 smoke 均通過。

2026-10-07 補充稽核重現輸出路徑碰撞、跨程序寫入互相干擾，以及 Modern／Historical 提交不一致三项 production 缺陷。因此不得將本文件的 v1.5.4 單程序回歸 PASS 延伸為無條件正式驗收。另人工 selector、部分 legacy reference 來源及全量逐圖視覺等同性仍有 `UNVERIFIED`。修復與新版本回歸將另行留存；本文件保留 v1.5.4 歷史結果並修正過度聲稱。

## 資料邊界與不可變性

- 唯一原始來源：`\\DS923plus\NAS\陳玄元\專案\AIS資料確認\2026\AIS資料\原始檔案`
- 實際處理使用隔離副本：`B:\AIS\work\ais32_full_regression_20261005\source_copy`
- 原始檔共 335 份 XLSX、21,269,682,462 bytes；來源與副本 335/335 SHA-256 相同。
- 原始資料與人工 reference 沒有被修改；所有產物位於 `B:\AIS\work\ais32_full_regression_20261005`。
- 來源 catalog：`source_catalog.json`；逐檔 copy hash：`copy_verification.json`。
- 處理後完整性重查已完成：`source_copy_integrity_final.json` 為 source 335 + copy 335，共 670/670 MATCH、0 mismatch、0 missing/extra。這是該次讀取的歷史證據；不代表每次沿用 checkpoint 都重新 hash 670 檔。

## 全量輸入盤點

| 港口 | 月份 | 實體 XLSX | logical days | 備註 |
|---|---:|---:|---:|---|
| HWLN | 2026-01 | 31 | 31 | 每日一檔 |
| HWLN | 2026-02 | 28 | 28 | 每日一檔 |
| HWLN | 2026-03 | 31 | 31 | 每日一檔 |
| HWLN | 2026-04 | 30 | 30 | 1–9 日為 `YYYYMMd` 七位日期 |
| HWLN | 2026-05 | 31 | 31 | 每日一檔 |
| HWLN | 2026-06 | 30 | 30 | 每日一檔 |
| KLNG | 2026-04 | 30 | 30 | 每日一檔 |
| KLNG | 2026-05 | 31 | 31 | 每日一檔 |
| KLNG | 2026-06 | 60 | 30 | 每日兩個互補 fragments |
| KLNG | 2026-07 | 33 | 31 | 7/8 是互補 fragments；7/24 是新舊高度重疊匯出檔 |

總計 335 份實體檔、303 個 logical days、10 個 port-month jobs。全部 workbook 均可開啟，資料表欄位索引一致；兩種 schema 的差異只在站點基點常數：HWLN `(121.6085958, 23.98358688)`，KLNG `(121.777586666667, 25.1499266666667)`。必要欄位是 Year/Month/Day/Hour/Minute/Second、channel、msg_type、mmsi、LONGITUDE_DESC、bearing、distance in nautical miles。

其中 KLNG 2026-07-24 的 `_23`／`_24` 不是可直接 union 的互補 fragments。兩檔 SHA-256 與 2026-09-18 已稽核版本完全相同：`_23 = 67FA00...19FC`、`_24 = 7916B9...8173A`。既有逐列 multiset 稽核證明 `_23` 的 708,418 筆 normalized rows 有 99.999435% 包含於 `_24`；四筆表面差異只是約 4e-16 NM 的浮點序列化差，`_24` 另有 22,586 筆真實新增資料。故 335 檔原樣直接輸入必須被 overlap safety gate 阻擋；正式回歸使用不修改來源的 334 檔安全 view，依既有 `safe_source_manifest.json` 排除被 `_24` 取代的 `_23`。這是來源適用性判定，不是程式在背後去重。

## 問題、根因與修復

### 1. HWLN 四月靜默漏掉九天

舊 parser 只接受 `YYYYMMDD`。HWLN 2026-04-01 至 04-09 的真實檔名是 `YYYYMMd`，因此舊版把 9 份來源當作「不適用檔」跳過，仍以 21/30 天完成且回報成功。這是本輪最嚴重的 silent omission，並非 bearing 分箱錯誤。

v1.5.4 明確接受七位或八位日期；七位格式固定解析為四位年、兩位月、一位日。任何看似 `D&TMOK ... .xlsx` 但日期或格式無法解析的檔名改為 fail closed，避免再次無聲漏檔。實際全量工作必須顯示 HWLN 2026-04 為 30/30 logical days。

### 2. 港別方位不是同一套遮罩

bearing 的共同 numeric mapping 為 `floor((bearing % 360) / 11.25)`，index 0–31 順時針。人工六月大 workbook 的兩個 period 各讀取前 100,000 筆時，floor mapping 均為 100,000/100,000；centered-bin 只匹配 44,292 與 48,137 筆。`reference_direction_mapping.json` 是這兩個片段的佐證，並不是人工 reference 全明細獨立比對完成的證據。

港別差異在後續 policy，而不是分箱：

- `auto + KLNG` → `klng21`，保留 index `0..10 + 22..31`。
- `auto + HWLN/其他` → `all32`，保留 index `0..31`。

歷史花蓮問題的根因是把 KLNG 的 21 方位地理遮罩共用到 HWLN；現有 policy 已把 mapping 與港別遮罩分離。`all32` 只表示不套 KLNG 遮罩，不表示已證明所有方向都是海域。

### 3. 「兩個表格」當次原因未重現

本輪單程序全量 pipeline 每個月份呼叫一次 Modern writer 與一次 Historical delivery writer，沒有在這兩輪重現額外輸出。這不能還原使用者先前那次異常；其根因仍為 `UNVERIFIED (original duplicate-output incident)`。Historical writer 的正常契約是五檔：

1. Period A 00:00–11:59 明細大檔。
2. Period B 12:00–23:59 明細大檔。
3. Period A 小總表。
4. Period B 小總表。
5. A/B 每格較大值整合總表。

再加一份 Modern logical-day 分析，因此每個 port-month 正常為 6 份 XLSX。不能只因正常契約有 A/B，就認定使用者當時看到的兩份檔案屬於正常輸出。回歸 validator 嚴格要求每月正好 6 份、無額外檔與缺檔。

2026-10-07 以隔離合成資料另重現三項輸出安全缺陷：Modern 路徑與整合總表相同時台帳被覆蓋；跨兩 writer 交錯時成功回報的成果可能混入不同 snapshot；Historical 失敗後 Modern 已更新而五份 Historical 保留舊批次。證據：`duplicate_risk_synthetic_20261007/findings.json`。這是新重現的同類風險，不能直接當成原事件的確切根因。

### 4. 圖表格式漂移

舊 renderer 使用通用六色循環、line radar、65% alpha、不同圖表尺寸／anchor，以及藍橙綠 MAX/MIN/AVG；人工六月成果則是 filled radar、30 色 Office palette、60% transparency、青色 MAX、橙色 MIN、紅色 2.25 pt 無 marker 平均線。KLNG 統計圖的視覺順序還是 `WSW..NbW, N..ESE`，不能直接以內部 `N..ESE, WSW..NbW` 順序作圖。

v1.5.4 將視覺契約集中在 Historical renderer：

- filled radar、30 日不重複 palette、`a:alpha=40000`。
- B:AG 可見表仍維持固定 32 方位順序；KLNG 另用正確 mapping 的 hidden helper 旋轉 21 方位圖表順序，不複製 reference 的舊錯位 range。
- 明確設定 Chart Area、Plot Area、座標軸、gridline、title、legend、scope-specific size/anchor。
- 統計圖為 MAX `#00B0F0`、MIN `#ED7D31`、平均 `#FF0000` / 2.25 pt / marker none。
- Excel 16.0 build 20430 原生開啟、FullRebuild、PNG 匯出後，六月三份小總表資料逐格 0 差異，6/6 generated/reference workbook 無公式錯誤且 read-only hash 不變。`visual_candidate_final_report.json` 證明三份總表的資料比對；`visual_comparison_com_final/excel_native_qa.json` 與其 `charts` 證明可呈現／匯出。這些結果不等同全部 120 張圖已逐張與 reference 外觀完全一致。

### 5. 7/24 重複來源不能合併

先前某次探索性執行曾在 cache 留下「移除 708,418 筆重複計入」的診斷；那不是目前正式程式允許的行為。正式 v1.5.4 仍採 fail closed：任何 exact normalized duplicate，或超過 250,000 筆而無法完整證明安全的 cross-fragment audit，都停止整月輸出，不產生或替換正式檔。

本次用目前 NAS 副本直接重跑，先保留安全門拒絕證據；再依 2026-09-18 的逐列稽核與既有 31-day `daily_final` 選擇 `_24`，建立 334 檔安全 view 從頭重跑全部 10 個月份。這同時避免兩種錯誤：把兩個近乎相同匯出檔重複計數，或在沒有證據時任意挑一檔。

### 6. 回歸驗證器的二月誤判

第一輪全檔 validator 最初回報 HWLN 二月整合總表兩個 chart anchor 不符。輸出 OOXML 的 anchor 實際上是整合總表的正確值；問題在驗證器用 `"_2" in filename` 判定 Period A，因而把檔名中的 `HWLN_2月` 誤當成 `_2` 輸出角色。已改為只接受檔尾 `_2` 或 `_2 總表`。修正後同一批 60 份檔重新掃描 116,244,596 筆 detail rows，0 issue；沒有放寬圖表 expected，也沒有修改產物。

## 人工六月成果與自動 selector 的邊界

人工／學長成果不是單一 tolerance 規則的直接輸出。0.5%–15% tolerance 掃描中最佳為 9.5%，只匹配 1,101/1,260（87.38%）；10% 匹配 1,099/1,260。不得為了讓測試通過修改 expected result 或反推一個假的通用 selector。

但五份人工成果的 1,260 個 period final 全部可追溯到同日、同 period、同方位的 normalized raw candidate，且 1,260/1,260 都在工具呈現的候選清單。將這些 Candidate ID 放回正式 decision ledger 後，重新由 writer 產生五檔，五個 30×32 grid 全部逐格 0 差異。證據：`golden_candidate_verification.json`、`reviewed_june_reproduction_report.json`。

因此工具可重現已覆核成果；但沒有人工 decision ledger 的其他月份，其研究判讀仍需人工覆核，不能把自動 selector 輸出宣稱為最終 domain truth。

## 驗證契約

`scripts/validate_full_regression.py` 對每輪全部月份檢查：

- run report 是 10/10 PASS，parser logical days 等於 catalog logical days。
- 每月正好 1 Modern + 5 Historical；沒有 `.building`、rollback 或額外 XLSX。
- B:AG 32 欄名稱、順序、日列、A/B 大小總表一致、Integrated = MAX(A,B)。
- KLNG 排除 11 方位保持空白；HWLN 32 方位均為 decision scope。
- 每一筆 retained detail row 的 bearing 重新計算 floor-bin；每方位 detail MAX 等於 period final。此版 validator 重用 production 方位 function，證明輸出內部一致性，不能單獨當成獨立原始資料 oracle。
- 圖表 source range、series cache、21/32 方位順序、類型、palette、alpha、顏色、線寬、marker、anchor。
- 第二輪每份 XLSX 除 `docProps/core.xml` 時戳 metadata 外的 semantic package hash 與第一輪一致。第二輪沿用同來源／同版本 cache，證明完整流程與 cache 重跑可再現，並非第二次全部重新解析 raw。

`scripts/excel_native_full_qa.ps1` 對最終全檔以已安裝 Excel read-only 開啟、`CalculateFullRebuild()`、搜尋公式錯誤、匯出所有圖表 PNG，並確認開啟前後 SHA-256 不變。

## 重現命令

在 repository root 與專案 `.venv` 下：

```powershell
python -m unittest discover -v
python -m scripts.audit_source_catalog <isolated-source-copy> --output <source_catalog.json>
python -m scripts.run_full_regression <isolated-source-copy> <run1-root> --workers 3
python -m scripts.validate_full_regression <run1-root> --output <run1-validation.json>
python -m scripts.run_full_regression <isolated-source-copy> <run2-root> --workers 3
python -m scripts.validate_full_regression <run2-root> --compare-root <run1-root> --output <run2-validation.json>
pwsh -File scripts/excel_native_full_qa.ps1 -Root <run1-root> -OutputDirectory <excel-native-qa-root>
```

Golden candidate、reference mapping、grid 重建與 tolerance 分析分別由：

- `scripts/verify_golden_candidates.py`
- `scripts/verify_reference_direction_mapping.py`
- `scripts/reproduce_june_reference.py`
- `scripts/compare_reference_grids.py`
- `scripts/analyze_reference_tolerance.py`

## 最終全量與發布狀態

### 實跑結果

- 原樣 335 檔 gate run：10 個月份中 9 PASS、KLNG 2026-07 預期 FAIL。失敗原因是 7/24 `_23`／`_24` cross-fragment audit 超過 250,000 筆，無法在執行中證明安全；該月 0 份正式 XLSX、0 個 `.building` 殘留。證據：`raw335_overlap_gate_run_report.json`。
- 安全 334 檔 run 1：10/10 PASS、303 logical days、60 份 XLSX。第一輪 validator：60/60、116,244,596 retained detail rows、0 issue。先前文字加總為304有誤；逐月run report一直是正確的303。
- 安全 334 檔 run 2：10/10 PASS、60 份 XLSX；與 run 1 的 60 份 semantic package hash 逐檔一致，0 mismatch。證據：`safe334_rerun_validation.json`。
- 測試：69/69 PASS、0 SKIP，含外部六月 golden/contract/baseline；執行時間 240.357 秒。證據：`tests_final.stderr.log`。
- Excel-native：Excel 16.0 build 20430 對 60/60 workbook 唯讀開啟並 `CalculateFullRebuild()`；0 formula error、120/120 chart PNG 匯出成功、60/60 開啟前後 SHA-256 不變。證據：`excel_native_full_qa_final/excel_native_qa.json`。
- packaged EXE：v1.5.4 對 HWLN 2026-04（30/30）及 KLNG 2026-07（31/31；沿用同來源同版本逐日 cache，以隔離 package/writer 差異）各產生 1 Modern + 5 Historical。五份正式交付的所有 package parts 除 `docProps/core.xml` 外逐一相同；Modern 只有輸出路徑、重新解析時的逐日耗時與 core metadata 可變，資料／圖表 parts 未出現其他差異。證據：`packaged_smoke_hwln_202604_audit.json`、`packaged_smoke_klng_202607_audit.json`。

### Reference 交叉比對解讀

- KLNG 2026-04 prior accepted automated set：5 檔、4,800 cells，逐格 0 差異。
- KLNG 2026-05 partial legacy reference：2 檔、1,984 cells、836 differences，其中 752 個雙方有值 cells 不同、84 個 reference 空白／本輪有值。此組來源角色及 selector 等同性尚未查明，保留 `UNVERIFIED (May legacy provenance/selector)`，不得省略或列為 PASS。
- HWLN 2026-01 reference 只含單一來源日；雙方都有值的 75 cells 為 75/75 相同，其餘差異全是新全月輸出有值而 reference 空白，不能把 31 日結果誤判為失敗。
- KLNG 2026-06 人工／學長 reference：自動 selector 仍有 427/3,150 個雙方有值 cells 不同；這是人工 decision ledger 邊界，不是 mapping 錯位。正式 decision ledger 回放的五個 30×32 grid 已另行驗證逐格 0 差異。
- KLNG 2026-07 prior accepted set：16/4,960 cells 不同，全部集中於 7/24；本輪使用逐列證據選出的 `_24`，舊 reference 對應不同 7/24 來源狀態，不應回退使用 `_23` 湊相同。
- 完整明細：`final_reference_comparison.json`、`golden_candidate_verification.json`、`reviewed_june_reproduction_report.json`。

### 可用性判定與限制

v1.5.4-local 的既有單程序全量結果已通過本文件列出的內容一致性、OOXML 契約及 Excel 開啟／重算／圖表匯出檢查。2026-10-07 新重現的三項輸出安全缺陷仍需修復並重新全量回歸；目前不做無條件正式可用判定。

仍保留以下明確邊界：

1. `UNVERIFIED (domain decision)`：未附人工覆核 decision ledger 的月份，自動 selector 是可追溯候選結果，不等同研究人員最終判讀。
2. `UNVERIFIED (public release provenance)`：本輪尚未建立 Git commit、tag、遠端 release，且 Windows EXE 未簽章；因此是已驗證本機版本，不是可追溯的公開正式 release。
3. `UNVERIFIED (original duplicate-output incident)`：未取得當次兩份檔案及執行紀錄，不能確定其根因。
4. `UNVERIFIED (all-output visual reference equality)`：全部圖表可匯出，不等同全部圖表逐張 reference／pixel 等同；部分月份沒有同來源同判讀的可比 reference。
5. `UNVERIFIED (May legacy provenance/selector)`：五月 partial legacy reference 的 836 格差異尚未有足夠來源證據解釋。
6. 人工 reference 全明細的獨立方位比對尚未完成；原先兩段各 100,000 筆不得標為全量驗證。
