# v1.5.6 來源安全、數值、覆核與工作簿契約

## 0. Source Preflight（不改選值）

正式分析／覆核重生前先完整串流檢查來源，以名稱辨識必要欄位並保存實際欄序、header fingerprint、來源 SHA、逐列日期分布及跨檔 normalized core 多重集合證據。GPS_* 只作鑑識式日期讀取，沒有研究語義等價授權，必須阻擋；不得當作標準時間 alias。

Mixed-date、未知／缺欄／重複必要欄位、損壞 workbook、exact core 重疊或 subset，以及既有 overlap audit 上限未通過，均阻擋整月。缺日／部分日是揭露的警告，非補值依據。SAFE 只表示本次來源身分符合契約；處理完成後再驗來源身分，變更時不得建立或替換正式成果。

完整重複／subset 由 preflight 先阻擋，沒有自動 dedupe 或排除。`SUBSET_CANDIDATE` 明確表示只證明 normalized required-field 多重集合包含（含重複次數），不是所有未使用欄位相同；排除須另留 manifest／SHA／完整證據並建立安全 view。高 occupied-second 重疊只作描述性指標，本身不等同重複訊息；不同 AIS 事件可互補，完全相同 normalized core 則不會默默加總。

JSON 與人讀摘要不寫進原始來源資料夾。Source days 與 calendar days 分開：cache 只對應真正來源日，完整日曆缺日的表格、ledger、helper 與 retained detail 必須保持無來源的空白／零列。本輪 selector／分箱／profiles／period／cap 均未修改。

## 1. Source fragment 與 logical day

來源檔名格式為 `D&TMOK <PORT>_YYYYMMDD_<opaque suffix>.xlsx`；為相容 HWLN 2026-04-01 至 04-09 的既有命名，也接受個位日期 `YYYYMMd`。7 位日期一律解析為四位年、兩位月、一位日，不猜測其他切法。識別鍵只有港別與日期；同一鍵下的 1、2、3 或更多檔案都是同一個 logical day 的 fragments。看似 `D&TMOK` AIS 檔卻無法解析時停止，不再靜默跳過。

- suffix 只保留供 provenance 顯示，不解讀為來源、班次或固定片數。
- 所有 fragments 以內容 hash 建立穩定身分；mtime、suffix 與檔案列舉順序不影響 selector。
- 正式月份入口先由 preflight 阻擋 SHA-256 完全相同的來源檔，要求明確的安全 view；底層相容 reader 的 aliases 表示法保留，但不作正常流程的自動排除授權。
- 每列儲存日期必須等於檔名日期；錯日、無效時間、損壞 workbook 或缺少必要 schema 會阻擋該月正式輸出。

## 2. 正規化處理與 normalized spool

預檢與正式正規化是不同階段，各自完整串流讀取；預檢不產生供 selector 使用的 spool。正式正規化中，每個 canonical fragment 以 openpyxl read-only 串流讀取一次。必要欄位為 Year、Month、Day、Hour、Minute、Second、channel、msg_type、mmsi、LONGITUDE_DESC、bearing 與 distance in nautical miles。

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

時間範圍重疊不表示實際資料重疊；若 occupied-second 集合沒有交集，就安全短路，不做 cross-fragment row audit。相同 timestamp 也可能是合法的不同 AIS 訊息。byte-identical files 只處理一次；可完整稽核且 exact normalized core records 為 0 的非同檔 fragments 照常合併，不對 possible duplicate events 做 silent dedupe。若 exact normalized core records 大於 0，或 audit 超過 250,000 筆上限而無法證明安全，整個月份會在正式檔建立／替換前停止，並要求移除重複檔或改用只含已確認互補 fragments 的安全來源資料夾。

## 4. 方位與兩套 profile

bearing 正規化至 `[0, 360)`，以每 11.25 度一格、floor-bin 映射 32 方位。0–11.249… 度是北，11.25–22.499… 度是北微東，依序順時針。

中文方向名稱來源為中文維基百科[〈羅盤方位〉](https://zh.wikipedia.org/zh-tw/羅盤方位)表的「中文名」欄，不採「又名」欄。v1.5.1 只修正十個中文 label，包括 SEbE = 東南微東、SEbS = 東南微南；32 個 index、英文縮寫、bearing 分箱與研究數值不變。

方位政策獨立於訊息 profile。`direction_policy=auto` 對 KLNG 解析為 `klng21`（index `0..10 + 22..31`），其他站點解析為 `all32`（index `0..31`）。可由 CLI／GUI 明確指定。

`klng21` 延續 KLNG 歷史排除；不適用的 11 方位 final grid 保持空白。`all32` 不套用地理方向排除，32 方位均可覆核及交付，但無有效候選仍留白。不能把 `all32` 說成已證明 HWLN 或任意新站全部方向都屬海域，也不能把舊固定空白解釋為零 AIS。

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

對每個 scope 與選定方位：

1. 套用該 scope 的 message profile 與設定的距離上限（預設 500 NM）。
2. 距離由高至低排列，穩定來源順序只用於同距離 tie-break。
3. 從最高順位開始，以該值的 90% 為下限搜尋 rolling window。
4. 第一個至少含 3 筆的 window，其最高值即自動值。
5. 若找不到群聚，保留最高有效值並標示待複核；完全無合格列則留白。

`klng21` 的西南偏西、西微南、西值高於 10 NM 時另列原有沿岸風險；此站點警示不套用 `all32`。`top_candidates` 只截取顯示清單；完整 selector 永遠搜尋全部合格列，且是否 renderer 五檔不得改變 selected、rank 或 cluster count。

## 6. Candidate provenance 與 decision ledger

Candidate ID 由實際 fragment content identity 與 normalized source record 建立，並可回查 timestamp、MMSI、channel、fragment、sheet 與 source row。Decision key 是 `(scope, day, direction)`，三種 scope 共用一份 snapshot：

- `logical_day`：可填同 scope 的 Candidate ID、留白，或非負有限 forced numeric override。
- `period_a`／`period_b`：只能填同日、同 period、同方位、Historical profile 且不超過 selection cap 的實際 Candidate ID，或留白；numeric override 是錯誤。

finalization 會驗證 workbook schema、公式、decision row count、港別／月份、完整 key set、目前來源 fragment manifest、cache/spool checksum 與 Candidate ID。Modern workbook 可另存新檔，但必須仍能指向建立它的原始分析路徑與 caches。

v1.5.1 將 `CACHE_VERSION` 從 4 升至 5，使舊中文 direction keys 的 cache 安全失效。`NORMALIZED_SPOOL_VERSION` 維持 1，binary 中的 numeric index 語意不變。`REVIEW_WORKBOOK_SCHEMA` 升至 `AIS_V15_REVIEW_2`；兩個覆核讀取入口都拒絕 v1.5.0 的 `AIS_V15_REVIEW_1`，並要求由原始月份資料重新分析、覆核，避免 decision direction 錯置。不實作舊 workbook 自動 migration。

v1.5.2 為 `CACHE_VERSION=6`、`REVIEW_WORKBOOK_SCHEMA=AIS_V15_REVIEW_3`。方位政策納入設定身分及覆核驗證，舊覆核檔須重新分析，避免 HWLN 的舊 21 方位台帳被默認為已覆核 32 方位。Normalized spool 的分箱語意維持不變。

v1.5.3 將 `CACHE_VERSION` 升為 7，讓 cache 6 必定失效並從原始 fragments 重掃，避免舊 DayResult 繞過 exact-duplicate／audit-truncated gate。v1.5.4 的檔名與 renderer 修復不改 normalized numeric schema；`CACHE_VERSION=7`、`NORMALIZED_SPOOL_VERSION=1` 與 `REVIEW_WORKBOOK_SCHEMA=AIS_V15_REVIEW_3` 不變。

對政策選定的 21 或 32 方位，blank period decision 表示該 period final 留白，renderer 不保留該方位 detail rows，以維持 detail MAX 契約。只有 `klng21` 排除的 11 方位屬結構性空白；`all32` 的全部方向均為 decision scope。

## 7. 五份正式 Historical delivery

兩個 period 各生成一份含每日明細的大 workbook 與一份小總表；第五份為整合總表。每次首次分析與覆核 finalization 都由同一個 immutable decision snapshot 重建全部五份。

固定契約：

1. 政策選定方位的 `MAX(retained detail rows) = period final`。
2. blank period final 不保留該方位 rows。
3. 小總表 B:AG 日值 grid 等於對應大 workbook 的總表。
4. 第五份每格 `Integrated = MAX(Period A final, Period B final)`；兩者皆空才空白。
5. MAX、MIN、AVERAGE、STDEV、雷達圖及 MAX/MIN/AVERAGE 統計圖全部由當次 grid 重算。

Modern logical-day union-first selection 與 `MAX(period selections)` 並不具有數學等價性，因此兩者在 schema、GUI 與文件均使用不同名稱，不宣稱共用一個模糊 canonical value。

覆核 finalization 先完成五個 `.building.xlsx`，再備份舊五檔並整組 promote。v1.5.5 首次完整分析則把 Modern 與五份 Historical 全部先寫入同磁碟的隔離 staging，六份完成後才整組 promote；一般錯誤會復原原 snapshot。final／temporary／rollback 路徑碰撞在寫檔前拒絕，同一台電腦上的 GUI／CLI 程序共用輸出與 cache／spool 資源鎖。若程序崩潰留下 lock 或新舊並存的 rollback，保留現況並停止要求人工確認；此鎖不宣稱能協調不同電腦同時寫入同一 NAS 位置。

每張工作表最多 1,048,576 列，扣除標題後明細上限為 1,048,575。超限在正式發布前報錯；v1.5.1 不截斷，也不自動 split。

### 雷達圖與選值上限

全部雷達均引用 32 欄 B:AG；方位政策決定哪幾欄可以有值，缺值不補 0。`radar_style=auto` 依六月確認成果使用 filled radar；可明確指定 line／filled。Historical renderer 使用 reference 的 30 色 palette、60% transparency、圖表尺寸／位置與 MAX／MIN／AVG 配色，但保留修正後的實際資料 range，不複製 reference 的舊錯位 range。

`radar_max` 只影響圖表刻度，預設留白使用自動尺度；指定時不得裁掉該圖要呈現的有效數值。它不篩選資料。`max_distance` 是 selector 上限（預設 500 NM），不是雷達軸上限。參考照片的 150、300 不構成所有月份必須固定該刻度的研究規則。

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
