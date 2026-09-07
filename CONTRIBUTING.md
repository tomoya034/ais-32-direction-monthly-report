# Contributing

1. 建立功能分支。
2. 不要提交真實 AIS、Excel 月報、快取或含船舶識別資訊的資料。
3. 修改數值規則時，同步更新 `docs/ALGORITHM.md` 與測試。
4. 執行 `python -m unittest discover -v`。
5. 修改 fragment、selector、decision ledger 或 renderer 時，必須同時證明 `detail MAX = period final`、`small total = big workbook total` 與 `integrated = MAX(A,B)`。
6. 真實六月 regression 只能透過 `AIS_JUNE_*` 環境變數引用 repository 外資料；不得提交原始 AIS、Golden workbook、baseline JSON、cache 或 spool。
7. Pull Request 說明應包含變更原因、對 Modern／Historical 語意與輸出格式的影響，以及 unit、external regression、EXE smoke 的驗證方式。
