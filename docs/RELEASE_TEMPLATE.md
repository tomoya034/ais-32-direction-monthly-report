# GitHub Release 模板

本文件定義 AIS 32 方位月報工具的固定 Release 格式。未來版本原則上沿用此結構，避免每次發布重新設計版面。

# AIS 32 方位月報工具 vX.Y.Z

一句話說明本版本最重要的變更與定位。

## 📥 下載

Windows 一般使用者請下載正式 Windows 發布包：

`AIS_32_Direction_Monthly_Report_vX.Y.Z.zip`

解壓縮後執行：

`AIS_32方位月報工具.exe`

不需要安裝 Python。

> `Source code (zip)` 與 `Source code (tar.gz)` 是 GitHub 自動產生的原始碼封裝，不是一般 Windows 使用者的執行版。

## ✨ 本版重點

- 變更重點 1
- 變更重點 2
- 變更重點 3

## ✅ 驗證

- 自動測試：`X/X` 通過
- Source smoke test：通過／未執行
- Windows EXE smoke test：通過／未執行
- GitHub Actions：通過／未執行
- 真實資料 smoke test：通過／未執行／不適用

只填寫實際完成的驗證，不要推測或補寫。

## ⚠️ 已知限制

- 列出仍需要人工處理或尚未解決的限制。
- 不要把既有研究 UNKNOWN 描述成已解決。

## 🔐 SHA-256

ZIP：

`<ZIP_SHA256>`

EXE：

`<EXE_SHA256>`

## ⬆️ 從舊版升級

說明此版本是否可直接下載新版使用，以及 cache、輸入資料或其他相容性注意事項。

## 📋 完整變更

完整版本變更請見 [`CHANGELOG.md`](../CHANGELOG.md)。

---

## Release Asset 規範

正式 Windows Release 原則上固定提供：

- `AIS_32_Direction_Monthly_Report_vX.Y.Z.zip`
- `SHA256SUMS.txt`
- GitHub 自動產生的 `Source code (zip)`
- GitHub 自動產生的 `Source code (tar.gz)`

正式 ZIP 內至少包含：

- `AIS_32方位月報工具.exe`
- `使用說明.txt`

除非有明確理由，避免不同版本交替使用「單獨 EXE」、「ZIP」、「不同命名規則」等不同發布方式。既有歷史 Release 不為了排版一致而重寫 tag 或 commit；資產形式不同時應如實保留。

## 每次發布檢查清單

1. 更新 `VERSION`、`APP_VERSION`、`pyproject.toml`、`version_info.txt` 等版本資訊。
2. 將本版變更從 `CHANGELOG.md` 的 `[Unreleased]` 整理到正式版本區塊。
3. README 只有在功能或操作方式改變時才更新，不堆疊歷史版本更新紀錄。
4. 執行完整 tests。
5. Build Windows EXE。
6. 執行 Source / EXE smoke test。
7. 建立正式 ZIP。
8. 產生 `SHA256SUMS.txt`。
9. Commit 並確認 working tree clean。
10. Push `main`。
11. 建立並 push annotated tag。
12. 使用本模板建立 GitHub Release。
13. 上傳正式 ZIP 與 `SHA256SUMS.txt`。
14. 從 GitHub Release 重新下載正式發布檔。
15. 重新驗證 SHA-256 與可執行性。
