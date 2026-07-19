# Contributing

1. 建立功能分支。
2. 不要提交真實 AIS、Excel 月報、快取或含船舶識別資訊的資料。
3. 修改數值規則時，同步更新 `docs/ALGORITHM.md` 與測試。
4. 執行 `python -m unittest -v test_ais_monthly_app.py`。
5. Pull Request 說明應包含變更原因、對輸出格式的影響及驗證方式。
