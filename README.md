# AIS 32 方位月報一鍵製作

[![Tests](https://github.com/tomoya034/ais-32-direction-monthly-report/actions/workflows/tests.yml/badge.svg)](https://github.com/tomoya034/ais-32-direction-monthly-report/actions/workflows/tests.yml)
[![Release](https://img.shields.io/badge/release-v1.5.0-blue)](https://github.com/tomoya034/ais-32-direction-monthly-report/releases/tag/v1.5.0)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

離線讀取 `D&TMOK <PORT>_YYYYMMDD_*.xlsx`，把同港別、同日期的任意數量檔案視為同一個 logical day，再一次產生 Modern research 分析及五份正式 Historical delivery 工作簿。檔名尾碼只供追溯，不代表固定來源或時段；程式不使用雲端 API，也不會上傳 AIS 資料。

> 目前版本為 `1.5.0`。程式會自動完成目前月份工作流程，但 AIS 異常的研究判斷仍需由領域人員在唯一的 decision ledger 覆核。

![AIS 32 方位月報工具視窗](docs/images/app-window.png)

## 下載與標準流程

1. 從 [Releases](https://github.com/tomoya034/ais-32-direction-monthly-report/releases) 下載 `AIS_32_Direction_Monthly_Report_v1.5.0.zip`。
2. 解壓縮後雙擊 `AIS_32方位月報工具.exe`，不必安裝 Python。
3. 選擇原始月份資料夾；工具會從檔名辨識港別與年月。
4. 按「開始全自動製作」。第一次執行會產生一份 Modern 分析與五份正式成果。
5. 在 Modern workbook 的「決策台帳」覆核 `logical_day`、`period_a`、`period_b`，然後儲存該 workbook。
6. 回到程式按「讀取已覆核分析並重生五份成果」。程式會驗證來源 manifest 與決策，再由同一份 decision snapshot 以 rollback 保護整組替換五份成果。

不要直接修改五份交付 workbook；再次 finalization 會以已儲存的 Modern decision ledger 為唯一決策來源。

Windows 可能因 EXE 尚未做商業程式碼簽章而顯示未知發行者。請只從本專案 Release 下載，並核對 Release 提供的 SHA-256。

## 六份輸出及其語意

Modern workbook：

- `<PORT>_YYYY年MM月_32方位數值_新版自動分析.xlsx`
  - `logical_day` 使用當日所有 fragments 先合併、再完整搜尋群聚。
  - 內含總表、每日分析、候選 provenance、待複核、決策台帳、來源 manifest 與圖表。

五份正式成果預設置於 `正式交付/<PORT>_YYYYMM/`：

- `<PORT> M月 32方位數值_2.xlsx`：Period A 完整工作簿。
- `<PORT> M月 32方位數值.xlsx`：Period B 完整工作簿。
- `<PORT> M月 32方位數值_2 總表.xlsx`：Period A 小總表。
- `<PORT> M月 32方位數值 總表.xlsx`：Period B 小總表。
- `<PORT>_M月_32方位_每日較大值整合總表.xlsx`：逐格整合總表。

Period 是依每列實際時間切分，與 `_11`、`_23` 或其他尾碼無關：

- Period A：`00:00:00 <= time < 12:00:00`。
- Period B：`12:00:00 <= time < 24:00:00`。

五檔固定維持以下可驗證契約：

- `Period detail MAX = Period final`（21 個海向；period final 留白時不保留該海向明細）。
- `Period 小總表 = 對應大 workbook 的總表`。
- `Integrated = MAX(Period A final, Period B final)`。
- 統計與兩張圖一律由本次 grid 重算，不複製舊 workbook 的過期統計值。

Modern logical-day final 與 Historical integrated 是兩種不同語意；因為群聚選值是非線性運算，前者不保證等於後者。

## 兩套 profile 與覆核契約

- Modern research profile 預設訊息類型為 `1, 2, 3, 18, 19`，可在 GUI 或 `--message-types` 修改。
- Historical delivery profile 固定為 `1, 3, 4, 18, 19`；修改 Modern profile 不會改變五份正式成果。
- `logical_day` 可選擇同 scope 的 Candidate ID、留白，或輸入非負有限 numeric override。
- `period_a`／`period_b` 只能選擇同日、同 period、同方位、Historical profile 的真實 Candidate ID，或留白；不得輸入任意數字。
- `top_candidates` 只控制 workbook 顯示筆數。selector 永遠搜尋完整合格資料，不能因顯示數量或是否輸出五檔而改變研究值。

候選保存 timestamp、MMSI、channel、來源 fragment、工作表與來源列號，供人工判斷同船、同秒或跨 fragment 群聚風險。完整規則見 [docs/ALGORITHM.md](docs/ALGORITHM.md)。

## 多 fragment、重複與異常資料

- 同港同日的所有 fragments 都會納入；排序不依賴 mtime、尾碼或檔案列舉順序。
- byte-identical 檔案只處理一次，其他路徑記為 alias。
- time-range overlap、occupied-second overlap、正常同秒不同訊息、possible duplicate event 與 normalized-core 相同紀錄分開報告。
- timestamp 或 occupied second 相同不等於 duplicate；非 byte-identical fragments 在缺乏領域 dedupe key 前不會被靜默去重。
- 損壞 XLSX、缺少正式 provenance 欄位、列內日期不符檔名 logical day 或 spool checksum 不符會阻擋正式輸出。

程式不建立合併後的巨大來源 Excel。每個 logical day 會建立 profile-neutral normalized spool，保留 East、核心欄位有效的必要資料、`msg_type`、時間與 provenance，也保留超過 500 NM 的列；500 NM 只在 selector／renderer 階段套用。

## 快取、容量與失敗保護

- 每完成一個 logical day 就保存帶來源簽章與 checksum 的結果及 normalized spool；來源 metadata、分析設定或 cache/spool schema 變更時，對應日會安全重算。
- 目前採每個 logical day 串流讀取後，以 bounded worker 數進行記憶體內 compact sorting；實測六月最高量日仍在可接受資源內，因此 v1.5.0 不加入 external merge 複雜度。
- 五份成果先全部寫入 `.building.xlsx`，成功後才整組替換；替換失敗會復原原有五檔。
- 每張 Excel 明細最多 `1,048,575` 筆資料列；超限會在發布前停止，絕不截斷。v1.5.0 不提供 split 模式。
- 失敗時會保留已驗證的每日快取，並在輸出資料夾建立 `AIS月報_錯誤報告_*.txt`。

## 命令列

首次分析並產生全部六份成果：

```powershell
python .\ais_monthly_app.py `
  --input "D:\AIS\2026_06" `
  --output "D:\AIS\KLNG_2026年06月_32方位數值_新版自動分析.xlsx" `
  --delivery-dir "D:\AIS\正式交付\KLNG_202606" `
  --port KLNG --workers 2 --overwrite
```

儲存人工覆核後，重新產生五份正式成果：

```powershell
python .\ais_monthly_app.py `
  --finalize-from "D:\AIS\KLNG_2026年06月_32方位數值_新版自動分析.xlsx" `
  --delivery-dir "D:\AIS\正式交付\KLNG_202606" `
  --overwrite
```

`--max-days N` 只供測試，永遠保留前 N 個完整 logical days；`--max-files` 與 `--legacy-output` 僅保留為 deprecated 相容參數。來源含多港別時必須指定 `--port`；`--year` 與 `--month` 只在同一港別含多月份時成對使用。

## 從原始碼執行與測試

需求：Windows、Python 3.11 以上。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python .\ais_monthly_app.py
python -m unittest discover -v
```

一般測試只建立去識別化小型 fixtures，不需要真實 AIS。六月 external regression 由本機環境變數提供資料，不會進 Git：

- `AIS_JUNE_INPUT_DIR`
- `AIS_JUNE_GOLDEN_DIR`
- `AIS_JUNE_CONTRACT_GATE_JSON`
- `AIS_JUNE_BASELINE_JSON`
- `AIS_JUNE_BENCHMARK_JSON`

## 建置 Windows EXE

```powershell
.\scripts\build_windows.ps1
.\scripts\package_release.ps1 -SkipBuild
```

建置結果位於 `dist\AIS_32方位月報工具.exe`；Release ZIP 與 `SHA256SUMS.txt` 也在 `dist`。

## 資料保護

- `.gitignore` 排除 Excel、CSV、cache、spool、`.building`、`.rollback`、錯誤報告與建置產物。
- 不要在 Git、Issue、Pull Request 或 Release 附上真實 AIS、含 MMSI 的分析檔或 Golden workbook。
- 程式完全離線運作，不需要 API Key 或 Token。

## 授權

[MIT License](LICENSE)
