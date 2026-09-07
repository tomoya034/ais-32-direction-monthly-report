# v1.5.1 數值、覆核與工作簿契約

## 1. Source fragment 與 logical day

來源檔名格式為 `D&TMOK <PORT>_YYYYMMDD_<opaque suffix>.xlsx`。識別鍵只有港別與日期；同一鍵下的 1、2、3 或更多檔案都是同一個 logical day 的 fragments。

- suffix 只保留供 provenance 顯示，不解讀為來源、班次或固定片數。
- 所有 fragments 以內容 hash 建立穩定身分；mtime、suffix 與檔案列舉順序不影響 selector。
- SHA-256 完全相同的檔案只處理一份，其餘路徑記為 aliases。
- 每列儲存日期必須等於檔名日期；錯日、無效時間、損壞 workbook 或缺少必要 schema 會阻擋該月正式輸出。

## 2. 單次讀取與 normalized spool

每個 canonical fragment 以 openpyxl read-only 串流讀取一次。必要欄位為 Year、Month、Day、Hour、Minute、Second、channel、msg_type、mmsi、LONGITUDE_DESC、bearing 與 distance in nautical miles。

profile-neutral spool 保存以下 normalized rows：

1. `msg_type` 是可表示的非負整數。
2. `LONGITUDE_DESC` 去除空白並忽略大小寫後等於 `East`。
3. bearing、distance 是有限數值，且 distance 非負。
4. 同時保存 msg_type、timestamp、MMSI、channel、fragment index、sheet、source row 與穩定 source key。

spool 不先套 Modern 或 Historical message profile，也不先刪除超過 500 NM 的列。500 NM 是 selection／delivery policy，不是 parser policy；因此同一份 spool 至少是 `{1,2,3,4,18,19}` 的超集合，也能支援來源追溯與日後明確調高距離上限。

程式不建立合併後來源 XLSX。v1.5.1 以每個 logical day 的 compact rows 在記憶體內按方向、距離降冪與穩定來源順序排序，worker 數限制為 1–8。

### 架構 benchmark 決策

以六月的 6/1、6/5、6/20 比較 simple in-memory 與 bounded external runs；兩者 selector checksum 完全一致：

| 日期 | simple wall / peak RSS | external wall / peak RSS / temp disk |
|---|---:|---:|
| 6/1 | 60.49 s / 173.9 MiB | 58.84 s / 97.7 MiB / 12.1 MiB |
| 6/5 | 87.65 s / 214.6 MiB | 99.52 s / 104.7 MiB / 17.1 MiB |
| 6/20 | 132.08 s / 287.7 MiB | 131.43 s / 121.5 MiB / 24.8 MiB |

最高量日 simple peak RSS 約 288 MiB，時間未劣於 external runs，且沒有額外 run lifecycle／磁碟故障面。因此 external runs 不列入 v1.5.1；若未來單日資料量明顯超過本 benchmark，再以相同 checksum contract 加入，不改 workbook 或 decision schema。

## 3. Overlap 與 duplicate 診斷

不同 fragments 的診斷分層如下：

1. time-range overlap：兩檔觀測起訖範圍相交。
2. occupied-second overlap：兩檔實際有資料的秒集合相交。
3. normal same-second different AIS messages：同秒內存在不同 msg_type／MMSI 身分的訊息。
4. possible duplicate events：同秒、msg_type、MMSI 相同，但 normalized core 不完全相同。
5. exact normalized core records：同秒、msg_type、MMSI、channel、bearing、distance 均相同。
6. byte-identical duplicate files：整檔 SHA-256 相同。

時間範圍重疊不表示實際資料重疊；若 occupied-second 集合沒有交集，就安全短路，不做 cross-fragment row audit。相同 timestamp 也可能是合法的不同 AIS 訊息。v1.5.1 只對 byte-identical files 合併處理；其他情形只診斷、不 silent dedupe。

## 4. 方位與兩套 profile

bearing 正規化至 `[0, 360)`，以每 11.25 度一格、floor-bin 映射 32 方位。0–11.249… 度是北，11.25–22.499… 度是北微東，依序順時針。

中文方向名稱來源為中文維基百科[〈羅盤方位〉](https://zh.wikipedia.org/zh-tw/羅盤方位)表的「中文名」欄，不採「又名」欄。v1.5.1 只修正十個中文 label，包括 SEbE = 東南微東、SEbS = 東南微南；32 個 index、英文縮寫、bearing 分箱與研究數值不變。

北至東南偏東、西南偏西至北微西，共 21 方位；`OPEN_SEA_INDEXES` 固定為 index `0..10 + 22..31`，中間 11 個陸向的 final grid 保持空白。

### Modern research

- scope：`logical_day`，先 union 同日所有 fragments，再選值。
- 預設 profile：`{1,2,3,18,19}`；可由使用者明確修改。
- Type 4 不因 Golden 歷史流程而混入 Modern 預設研究值。

### Historical delivery

- scope：`period_a` 與 `period_b` 各自選值。
- 固定 profile：`{1,3,4,18,19}`，不受 Modern 設定影響。
- Period A：row timestamp 的 `00:00:00 <= time < 12:00:00`。
- Period B：row timestamp 的 `12:00:00 <= time < 24:00:00`。
- period 與實體 fragment 或 `_11`／`_23` suffix 無關；單一 fragment 可以含兩個 periods，多個 fragments 也可以屬同一 period。

## 5. 完整群聚 selector

對每個 scope 與海向：

1. 套用該 scope 的 message profile 與設定的距離上限（預設 500 NM）。
2. 距離由高至低排列，穩定來源順序只用於同距離 tie-break。
3. 從最高順位開始，以該值的 90% 為下限搜尋 rolling window。
4. 第一個至少含 3 筆的 window，其最高值即自動值。
5. 若找不到群聚，保留最高有效值並標示待複核；完全無合格列則留白。

西南偏西、西微南、西的值高於 10 NM 時另列風險。`top_candidates` 只截取顯示清單；完整 selector 永遠搜尋全部合格列，且是否 renderer 五檔不得改變 selected、rank 或 cluster count。

## 6. Candidate provenance 與 decision ledger

Candidate ID 由實際 fragment content identity 與 normalized source record 建立，並可回查 timestamp、MMSI、channel、fragment、sheet 與 source row。Decision key 是 `(scope, day, direction)`，三種 scope 共用一份 snapshot：

- `logical_day`：可填同 scope 的 Candidate ID、留白，或非負有限 forced numeric override。
- `period_a`／`period_b`：只能填同日、同 period、同方位、Historical profile 且不超過 selection cap 的實際 Candidate ID，或留白；numeric override 是錯誤。

finalization 會驗證 workbook schema、公式、decision row count、港別／月份、完整 key set、目前來源 fragment manifest、cache/spool checksum 與 Candidate ID。Modern workbook 可另存新檔，但必須仍能指向建立它的原始分析路徑與 caches。

v1.5.1 將 `CACHE_VERSION` 從 4 升至 5，使舊中文 direction keys 的 cache 安全失效。`NORMALIZED_SPOOL_VERSION` 維持 1，binary 中的 numeric index 語意不變。`REVIEW_WORKBOOK_SCHEMA` 升至 `AIS_V15_REVIEW_2`；兩個覆核讀取入口都拒絕 v1.5.0 的 `AIS_V15_REVIEW_1`，並要求由原始月份資料重新分析、覆核，避免 decision direction 錯置。不實作舊 workbook 自動 migration。

對 21 個可覆核海向，blank period decision 表示該 period final 留白，renderer 不保留該方位 detail rows，以維持 detail MAX 契約。11 個陸向是結構性空白，其明細不是人工 decision scope。

## 7. 五份正式 Historical delivery

兩個 period 各生成一份含每日明細的大 workbook 與一份小總表；第五份為整合總表。每次首次分析與覆核 finalization 都由同一個 immutable decision snapshot 重建全部五份。

固定契約：

1. 21 海向的 `MAX(retained detail rows) = period final`。
2. blank period final 不保留該海向 rows。
3. 小總表 B:AG 日值 grid 等於對應大 workbook 的總表。
4. 第五份每格 `Integrated = MAX(Period A final, Period B final)`；兩者皆空才空白。
5. MAX、MIN、AVERAGE、STDEV、雷達圖及 MAX/MIN/AVERAGE 統計圖全部由當次 grid 重算。

Modern logical-day union-first selection 與 `MAX(period selections)` 並不具有數學等價性，因此兩者在 schema、GUI 與文件均使用不同名稱，不宣稱共用一個模糊 canonical value。

寫檔時先完成五個 `.building.xlsx`，再備份舊五檔並整組 promote。任一步驟失敗會清除新檔並復原原 snapshot；若程式崩潰留下新舊並存的 rollback 狀態，下次執行會停止要求人工確認，避免混合月份成果。

每張工作表最多 1,048,576 列，扣除標題後明細上限為 1,048,575。超限在正式發布前報錯；v1.5.1 不截斷，也不自動 split。

## 8. Golden regression 邊界

真實 AIS 與 Golden workbooks 均在 repository 外部。六月 regression 分層驗證：

- 30 logical days、60 fragments 與來源 schema／時間完整性。
- 數值比較以 direction index／英文 abbreviation 識別，先將舊中文 label 的 baseline 轉為固定 index；名稱修正不得造成假 regression。
- production union-first Modern selector 對既有 full-search baseline 的 selected、rank、cluster count 與 checksum。
- 1,260 個 Golden 非空 period finals 全部可表示為 Historical profile 的真實 Candidate ID。
- 五檔的 detail MAX、small=big、integrated MAX、公式與圖表契約。
- Golden 舊統計區已知 9 個 stale results 不作新輸出的預期值；統計必須由正確日值 grid 重算。

歷史人工異常刪除規則沒有足夠證據可還原，因此 Golden period grid 是覆核後真值，不用來反向改寫自動 selector 或把 Type 4 加入 Modern profile。

### 完整六月回歸的重現方式

`scripts/june_regression.py` 接受外部路徑，重新處理 30 天／60 fragments、套用既有 Golden candidate evidence，寫出 Modern 與五份正式成果，驗證數值、公式、統計及圖表。`--v150-cache-dir` 僅讀取 v1.5.0 已完成分析的 cache 作數值 baseline，不作新執行的快取；輸出目錄應使用新的空資料夾，確保全月重讀。

```powershell
python scripts/june_regression.py --repo . `
  --input $env:AIS_JUNE_INPUT_DIR --golden $env:AIS_JUNE_GOLDEN_DIR `
  --gate $env:AIS_JUNE_CONTRACT_GATE_JSON --baseline $env:AIS_JUNE_BASELINE_JSON `
  --v150-cache-dir $env:AIS_V150_JUNE_CACHE_DIR `
  --output-root $env:AIS_JUNE_REGRESSION_OUTPUT --workers 3
python scripts/audit_detail_max.py --repo . `
  --delivery "$env:AIS_JUNE_REGRESSION_OUTPUT/正式交付/KLNG_202606" `
  --report "$env:AIS_JUNE_REGRESSION_OUTPUT/detail_max_audit.json"
```

獨立 detail audit 使用 `lxml` 串流回讀正式輸出；未安裝時需先安裝此測試工具依賴。數值比較以固定 index 為 key，v1.5.0 cache 中的舊 label 只用於讀取時的 index 解碼。30 份 normalized spool 的 SHA-256 必須與 v1.5.0 完全一致。完整 raw、Golden、cache、evidence JSON、regression output 都只保留在外部路徑。
