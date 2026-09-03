# AIS 32 方位月報一鍵製作

[![Tests](https://github.com/tomoya034/ais-32-direction-monthly-report/actions/workflows/tests.yml/badge.svg)](https://github.com/tomoya034/ais-32-direction-monthly-report/actions/workflows/tests.yml)
[![Release](https://img.shields.io/badge/release-v1.4.0-blue)](https://github.com/tomoya034/ais-32-direction-monthly-report/releases/tag/v1.4.0)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

離線讀取每日 `D&TMOK <PORT>_YYYYMMDD_*.xlsx`，自動辨識港別與年月，一次產生新版自動分析與原格式相容版兩份 Excel 月報。程式不使用雲端 API、不需要 Token，也不會上傳原始 AIS 資料。

> `0.0.1` 是首個可分享測試版本。數值規則已自動化，但航向與岸向異常仍應由領域人員覆核。
>
> 目前版本為 `1.4.0`；航向與岸向異常仍需由領域人員完成人工複核。

![AIS 32 方位月報工具視窗](docs/images/app-window.png)

## 下載與使用

1. 前往 [Releases](https://github.com/tomoya034/ais-32-direction-monthly-report/releases) 下載 `AIS_32_Direction_Monthly_Report_v1.4.0.zip`。
2. 解壓縮後雙擊 `AIS_32方位月報工具.exe`，不必安裝 Python。
3. 選擇含 `D&TMOK <PORT>_YYYYMMDD_*.xlsx` 的 AIS 資料夾。
4. 工具會從固定檔名辨識港別與年月；多港資料夾可選港別，多月份可選月份。
5. 保持勾選「同時產生原格式相容版」，按「開始全自動製作」。
6. 完成後先查看新版工作簿的「待複核」工作表。

Windows 可能因程式尚未做商業程式碼簽章而顯示未知發行者。請只從本專案 Release 下載，並核對 Release 提供的 SHA-256。

## 輸出

- `*_新版自動分析.xlsx`
  - 操作說明、每日候選、數值化判定理由、人工覆核欄、總表、待複核、處理紀錄與圖表。
- `*_原格式相容版.xlsx`
  - 每日 A:C 完整排序資料、H:AM 第 2 列 32 方位結果、舊式總表、圖表及隱藏的「工作」對照表。

## 影片手工作業已自動化

- 不必先人工建立每月活頁簿或逐日工作表。
- 不必逐方向套篩選、尋找異常高值、刪除整列、複製公式或重新排序。
- 程式會只讀原始每日檔，完成訊息類型與東經篩選、32 方位換算、500 NM 上限、群聚選值、逐日表、總表及待複核清單。
- 原始檔不會修改；排除原因、候選值與來源列號保留在新版分析檔供覆核。

## 港別與月份自動判定

- 港別、年份、月份以 `D&TMOK <PORT>_YYYYMMDD_*.xlsx` 的檔名為唯一依據；例如 `D&TMOK KLNG_20260701_23.xlsx` 與 `D&TMOK HWLN_20260101_23.xlsx`。
- 港別代碼採英數格式、大小寫不敏感並統一顯示為大寫；沒有把支援範圍寫死成 KLNG/HWLN 清單。
- 只有一個港別時會自動選定；含多個港別時可由唯讀清單選擇，月份清單會隨港別同步更新。
- 同一港別只有一個月份時會直接選定；若含多個月份則預選最新月份並允許改選。
- 開始製作前會再次掃描港別、年月與每日檔案；同一份月報只會處理一個港別。

## 預設數值規則

- AIS 船舶位置訊息類型：`1, 2, 3, 18, 19`。
- `LONGITUDE_DESC = East`。
- 自動值距離上限：500 NM；超過者保留統計但不採用。
- 每個方向由高至低尋找第一組「至少 3 筆、落在最高值 10% 範圍內」的群聚。
- 僅產出 21 個海向方位。
- 西南西、西微南、西高於 10 NM 時列入待複核。
- 每方向在新版保留前 50 筆候選與來源列號。

完整演算法與相容格式說明見 [docs/ALGORITHM.md](docs/ALGORITHM.md)。

## 快取與錯誤處理

- 每完成一天即保存快取；中斷後以相同來源與設定重跑會接續未完成日期。
- 原格式版使用逐日二進位暫存，避免為了第二份工作簿再次讀取大型來源檔。
- 會檢查缺少欄位、檔案損壞、月份不符、Excel 鎖檔、磁碟不足、記憶體不足、列數超限與平行程序失敗。
- 失敗時會在輸出資料夾建立 `AIS月報_錯誤報告_*.txt`。

## 從原始碼執行

需求：Windows、Python 3.11 以上。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python .\ais_monthly_app.py
```

命令列範例：

```powershell
python .\ais_monthly_app.py `
  --input "D:\AIS\2026_01" `
  --output "D:\AIS\HWLN_2026年01月_32方位數值_新版自動分析.xlsx" `
  --legacy-output "D:\AIS\HWLN_2026年01月_32方位數值_原格式相容版.xlsx" `
  --port HWLN `
  --workers 2 --overwrite
```

命令列會從檔名自動判定港別與年月。來源只有一個港別時不需要 `--port`，因此 v1.3.0 的 `--input`／`--output` 用法仍相容；來源含多個港別時必須以 `--port` 選擇。`--year` 與 `--month` 保留給同一港別含多個月份時使用，且必須成對指定。

## 測試

```powershell
python -m unittest -v test_ais_monthly_app.py
```

測試會在暫存目錄建立小型假資料，不需要真實 AIS 檔案。

## 建置 Windows EXE

```powershell
.\scripts\build_windows.ps1
```

建置結果位於 `dist\AIS_32方位月報工具.exe`。腳本會處理某些 Python 安裝路徑造成的 Tcl/Tk 封裝問題。

## 資料保護

- `.gitignore` 預設排除所有 Excel、CSV、二進位快取、錯誤報告與建置產物。
- 不要在 Issue、Pull Request 或 Release 附上真實 AIS 原始檔或含船舶識別資訊的工作簿。
- 程式完全離線運作，不需要 OpenAI API Key 或任何 Token。

## 授權

[MIT License](LICENSE)
