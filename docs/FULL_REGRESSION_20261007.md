# AIS 32 方位追加修復與全量驗證（2026-10-07）

## 最新續跑結果：2026-10-19，工程全量證據完成，正式放行仍NOT_GRANTED

同一1E2／AE528候選的兩輪334檔／303日／10港月／各60成果、raw303、selector10、
saved/default20、native20、兩種pairs各10及260張可見圖像已由actual lineage完整綁定。
來源／副本NEW generation已670/670 MATCH，10/19補hash10＋目前身分核對續用660，
兩側各335起訖清單無變、manifest未變、0mismatches／events；不是10/19全新hash670。
另完整掃六月30日23,426,967列，427個AUTO/manual差異格去重為161個A/B選點：
112人工捨棄合格高群、49人工選不足3筆支持高點。不能用加嚴同船/同秒/去重通殺。
本次只修停止競態與lineage/診斷support，production/reference/expected未改。
正式判定仍NOT_GRANTED，安裝版仍v1.5.4未換；無commit/tag/release。
最新完整交接：`B:\AIS\work\ais32_full_regression_20261005\REGRESSION_AUDIT_HANDOFF_20261019.md`；
actual彙整：同目錄`final_evidence_lineage_20261019.json`。
下方10/15及更早段落均保留為歷史快照，不取代本節。

## 最新續跑結果：2026-10-15 20:05，剩餘5%依約暫停，仍不能結案

同一1E2／AE528候選已完成actual EXE兩輪各10港月／60檔，raw303、selector10、
saved/default20、native20、兩種pairs各10，以及260張可見內容全量檢查。
但NEW來源／副本fresh generation只保存660/670match，ending inventory未完成，
actual最終lineage collector未跑。正常AUTO與人工June主表仍有427數值格差異；
不是approved-ledger replay就代表AUTO相同。正式可用判定仍NOT_GRANTED。
原始／reference／expected未改，安裝版未換，未commit/tag/release。
最新停止／續跑／限制紀錄：`B:\AIS\work\ais32_full_regression_20261005\PAUSED_20261015_2005.md`。
逐項根因與修法：同目錄`ISSUE_CATALOG_FOR_FINAL_20261015_1718.md`。
下方保留各時間點的歷史狀態，不以舊partial/PASS覆蓋本節結論。

## 最新狀態：2026-10-14 新發現Historical公式空字串被Excel畫成0，修復與新全量雙跑中

使用者明示恢復並維持任一Codex視窗剩<=5%停止；初讀99%／100%。最新source為1E2B0A98831E4768118C87954BBE0D2F65EBEB1FB600022C5BBA1577993376E4，實際EXE為AE528CFF84A4C49496599061B6717AA5B882B4F1F8401D873054CAB4BE55498A（新建置15.530s）。11:51:07/08於全新empty first/repeat目錄重新跑ALL334檔／303日／10港月，無舊cache復用，normal六產物流程不變。

新原生Jan31日全5Historical逐點觀察證實fullA397/fullB402/integrated348缺值被畫成Double0／實際雷達中心；兩summary真空白則null／plot外。人工June5檔逐點對照證明原參考缺值皆null／plot外，不是合法零，支持共通根因修正。現按source-kind判斷是否為formula-backed，這三種來源使用隱藏AO:BT精確NA() helper；主表空白與真0、數值、summary直連、stathelpers與人工外觀仍保留。88 focusedtests18.737sOK0skip；新全套suite、實際component native/reference pixel證明及全量雙回歸尚需完成，UNVERIFIED。製作／追查與證據索引 `HISTORICAL_NATIVE_MISSING_REPAIR_20261014.md`。

舊39C／7B producer及稽核作業已按exact身份停止，所有舊成果與失敗紀錄保留，不再冒充最新版驗收。舊307tests329.550sOK只屬39C快照。10/14重新確認335NAS＋335副本清單、670metadata及334safe-view samefile零差異；10/8完整670content-hash PASS仍保留，但沒有宣稱今天重新全部content-hash。原始/reference/正確numeric expected未改，安裝版仍C07E0E65...v1.5.4，未部署發布。

另已修 regression wrapper 暫停競態：`--pause-flag`在每個月啟動前／active子程序返回後檢查，遇quota中止記PAUSED而非錯誤資料FAIL，不再趁停止掃描間啟下一月；25 testsPASS。20個exact dead-owner lock與oldflag可恢復移存；一、二月曾source/cache/spool/final完整重新hash續跑，但新發現producer缺陷後這些仍只能保留歷史元件證據，最新版必須全部重做。

## 最新狀態：2026-10-08 13時繼續，39C951 候選全量雙跑中

使用者再次明示恢復並維持任一適用Codex視窗剩<=5%暫停。初讀5hr剩99%／weekly77%，舊quota flag可恢復移存，未使用reset credit。正式安裝版實讀仍為v1.5.4／C07E0E65...493610，未部署、commit、tag或發布。

原0D8的Modern統計圖缺值已改正，但新的原生Excel全31天雷達追查證實407個每日空值全部仍被轉Double0，屬另一個共通producer缺陷，不能把統計圖修正冒充兩圖正確。最新39C95153202A18E41AC0AF02BE91E70BD1ED5CB31E51E1A04708DDAEBD5C07F0增加Modern隱藏AO:BT圖表專用每日32方位helper，逐格明確NA()，主表／統計仍空白、真0仍0，雷達引用此範圍並啟用hidden data。Historical renderer未改。正確reference／原始檔未改。

新實際EXE為7B77532D41ED368C7762DBEAD5A69F90496206B15CC0C4500D94746B75D188B6，獨立建置9.850s。13:00:07 fresh/repeat皆於全新空root起跑全部334適用檔／303日／10港月，不共用舊cache、不限天數／rows／時間。月份execution PASS仍不代表內容正確；之後必須完成same-root raw3、selector3、全60份Excel／明細／引用／native／reference及全量重跑一致性。

新候選31天March隔離writer比對4,410,100 formula/value cells、79,000候選Excel rows、1,088 chart points，除2,151個嚴格限定的新增helper／NA語意／自身路徑變更外0issues。此為序列化與scope preservation證據，尚非新全量raw run驗收。NA例外必須綁定独立missing matrix、精確公式與来源，不能忽略所有#N/A或改正確expected讓PASS。Native兩圖全點missing=null／true0=0的實測及完整suite仍在完成。

Source/copy v3同一checkpoint續跑，670個MATCH已到齊，但最後ending inventory／metadata／manifest檢查與final report尚需完成；單有checkpoint不能宣告完整性PASS。可重現交接與exact-owned stop helpers：`CONTINUED_20261008_1253_CHART_NA.md`。下方0D8／D881／57BF及更早狀態均為保留歷史，非最新驗收。

## 最新狀態：2026-10-08 本輪再查出 Modern 缺值圖表根因（本節優先）

依使用者指示恢复後讀到5hr剩97%/weekly84%。先修Windows續跑保護把自身venv啟動器誤判live writer(22testsPASS)，recoverably archive55個已核對死owner的鎖；兩舊run的Jan–Apr各24final來源/cache/spool/final重新hash確認、完整Maypartial與originalreport搬往sibling archive。第一舊run續跑May成功，但嚴格saved-content gate實際捕獲March Modern最後缺值point被自動chartcache補0，故57BF/DB285E不再當final acceptance；所有原證據保留。

Modern共通writer先explicit None chartcache與合法all-empty blank radar fallback，D881/BAEB新EXE建置成功並啟完全empty雙root全334files/303days/10portmonths。但native Excel實證formula `""`仍在重算時變0、平均線落baseline，所以再次FAIL，root停止12個own processes，兩新run皆Januarypartial、0complete月份，未偽稱PASS。

最新source `0D8E28E963A0AC8A811CE892BA7AE78D4B61DC5BA19DF038169580520AB488DA`僅Modern圖表專用helper改COUNT0→NA，主表仍blank、真0保留；尚未建置與全面驗證，UNVERIFIED。嚴格native/formula/default gates沒有放寬成忽略NA。必須以精確helper公式/sourceCOUNT與獨立缺值證據證明intentionalNA後再更新gate，所有其他錯誤仍FAIL。

Source670與raw/candidate/retained-detail/native驗證器補強各有獨立記錄；670未完成不得以舊670代替。本輪新技術索引 `MODERN_MISSING_CHART_REPAIR_20261008.md`、`WINDOWS_RESUME_OWNERSHIP_20261008.md`、`NATIVE_MODERN_MISSING_D881_20261008.md`。所有原始/reference/expected未改，安裝v1.5.4未換、未commit/tag/release、未用reset。下列全部為歷史快照。

## 最新狀態：2026-10-08 03:33依剩5%門檻暫停（優先於全部歷史快照）

Live primary讀到剩5%後，root於03:32:59.540核對PID/建立時間/command停止9個自己的程序；再次census自身pipeline/selector為0。兩個最終actualEXE run實際各4/10 complete（HWLN1–4月120日、各24final檔），active5月first31cache/repeat10cache皆partial非PASS。Source agent另停自身hash/rawtrees；visual/renderer自有程序已退出。沒有reset credit、自動恢復或部署，所有partial/stage/rollback/original report保留。最新交接 `PAUSED_20261008_0333_USAGE_GATE.md`，明列確證、UNVERIFIED與安全續跑前置。最新Janselector亦已31/31PASS、6,525,488rows、2,976selectors、17,856fields；source fresh v3僅450/670 last-observed，未以歷史670代填。

## 2026-10-08 恢復狀態（優先於歷史暫停紀錄）

使用者明示恢復；02:53 live quota兩視窗皆剩100%，02:58再次確認primary剩84%、weekly剩98%。仍遵守任一適用視窗剩<=5%保存checkpoint並暫停，未使用reset credit。候選source9B7D995...、R2 EXE3421E580...、安裝版v1.5.4 C07E0E...已重新hash確認，沒有上一輪稽核工作殘留程序。02:56只將dead owner PID38204/48644、resource位於兩個明確舊run root的40個孤兒lock與暫停旗標可恢復移至 `resume_20261008_0255_orphan_locks`；完整manifest與可重現腳本保留，未刪stage/cache/rollback/raw/reference。

Integrated Normal/font0/theme繼承正在生成副本上隔離實驗；不得把未證明的字型假說當成根因。NAS現況重盤與全670項fresh雙端hash使用新report/checkpoint，不覆蓋10/7歷史PASS。最終候選尚未凍結、完整两輪334/303/10及60產物驗證仍UNVERIFIED，安裝版不變。

### 03:08已凍結與全量啟動（更新上述「尚未凍結」快照）

隔離實驗已證明Integrated文字owners的theme token在生成檔native解析成新細明體；只把既有文字owner的三script typeface明確指定Calibri，不改任何數字／公式／series／引用／幾何，即兩圖decoded PNG各0changedpixels、EXACT_EQUAL。Normal scheme-only、scheme+family-only及先前major/minor-only均無改善，不能將missing chartStyle或某一theme元素斷稱唯一機理根因；空chart parent原生Font仍新細不是可見文字錯誤，未擴大改parent。完整因果證據 `integrated_font_causal_evidence_20261008.md`。

凍結production source SHA256 `57BF060E507CB7BEEB826D5A1C83B5342D78354FC148BA8F7148F532DBF15428`，只Integrated角色literal Calibri，A/B theme triplets保留；67 focused tests PASS(14.111s)。03:07:35另起最新全部unittest，尚待完整結果，不能把舊225追認。新EXE建置完成v1.5.5、14,790,795bytes、SHA256 `DB285E656E6F8B0FC0707FB8A4E1941F632BECB7A698F6027960E4DE42488BBF`，位置 `dist_candidate_v155_renderer_final_20261008`。未部署，已安裝v1.5.4不變。

03:08:20啟動actualEXE全部334／303／10於absent fresh root `safe334_packaged_v155_final_fresh_20261008`，workers3、無only/daycap/processdeadline，保留正常CLI六產物流程；launcher53300/actual57108，詳 `final_packaged_fresh_launch_20261008.json`。所有完整same-root港月另由independent raw全月稽核，root待完整303日後獨立selector／default content／全60明細圖表引用及native、第二輪全部資料一致性驗證。execution PASS不等於正確性PASS。

稽核器亦修補「錯誤引用但快取與錯誤來源自洽」漏檢：新增Modern/Historical預定series/date/category/value exactbindings、Modern每日final→台帳→總表與統計／helper COUNTguard公式、Historical以獨立matrix重算helper統計／KLNG旋轉方位順序。11 focusedtests PASS，另對舊R2January全31日六檔6,528chartpoints實測0issues，報告 `default_content_R2_Jan_bindings_20261008.json`；只是舊完成月補充，不能冒充本候選全量。製作方法與負例 `DEFAULT_CONTENT_BINDINGS_20261008.md`。正確manual/golden expected未改。

### 03:30已驗證與仍待完成

- 最終57BF fresh五角色reference ledger replay：恰五份final、4,800日/方位cells0diff；native5/5 PASS、3,570 formulas0errors、10decodedPNG；全十圖與hash已核對reference全部EXACT_EQUAL/changed_pixels0，無resize/crop。Detailed native5本/10圖/5,115逐點styles、0gettererrors；raw COM theme/RGB/mixed/invisible屬性非逐字段相同，沒有把pixel一致誇大成所有XML/API欄位相同。這是renderer fidelity，不是default自動selector與人工領域決策等同性。證據 `visual_57bf_all5_fullqa_png_compare_20261008_0324.json`、`visual_57bf_all5_captured_style_compare_20261008_0327.json`。
- 所有helper凍結後full unittest於03:22另起，249 tests／264.997s／OK／0SKIP；`v155_complete_helpers_suite_20261008.stderr.log`。後續新增wrapper續跑保護另有tests，不能把此249快照冒充尚未再跑的新增helper完整suite。
- First actualEXE HWLN January default內容/公式/預定引用6本6,528chartpoints0issues；native第一次5Hist PASS而Modern台帳保護使SpecialCells拒絕，原FAIL/UNVERIFIED保留。已修唯讀ISFORMULA/ISERROR mask，不取消保護／不保存；負例fixture能正確抓出1formula error且排除constant error。新的January native6/6 PASS、10,784 formulas0errors、12decodedPNG、全部hash不變；方法 `PROTECTED_LEDGER_NATIVE_QA_20261008.md`。
- 獨立raw same-final-root HWLNJan31/31及Feb28/28 PASS／seal/stamps/census重檢0issues；Jan scanned7,250,517 normalized6,525,488，Feb scanned6,730,738 normalized6,060,405。共59日不是全303。最新Jan獨立selector已03:30:16啟動完整31日/所有spool rows、32方位×3scope×6field，尚待結果。不能把舊oracle或cache自洽替代本候選完整303日獨立證據。
- 重跑比對器原來只忽略core.xml的hash並非充分semantic compare：Modern自身路徑及處理紀錄perf_counter必然不同。已採狹義、逐field驗證canonicalization，只放行有效timestamps、自身路徑與有完整header/date/finite-nonnegative檢查的daily elapsed；source/數值/公式/chart/其他metadata/幾何皆仍比較。實際兩EXEJanuary Modern語意hash相同0856C943...DE3F41B，35provenance字段差異另記，不更改兩份原檔；16 dedicated與combined104tests PASS，詳 `SEMANTIC_REPEAT_PROVENANCE_20261008.md`。
- 兩個actualEXE run分別新empty root；first03:08、repeat03:11啟動（為縮短等待可安全並行唯讀來源，絕不共用cache/output）。到03:30兩輪各3/10 execution complete、HWLN April active；全60內容/明細/引用/native、303selector、ending670sourcehash及完整兩輪一致性仍待完成。人工領域選擇及原歷史「兩表」事故根因仍UNVERIFIED，未換安裝版。

## 最新狀態：23:42用量5%暫停（優先於下方歷史快照）

`PAUSED_BY_USER_USAGE_GATE`。live primary剩5%時停止；root清理完成23:43:07.615+08，自己的48972/39896/41884/48644/21088/34248/23132按command/建立時間核對後停止，remaining_own=0；visual agent自有作業亦均已退出。未用resetcredit、未自動恢復。quota_pause.flag保存停止條件，所有stage/cache/rollback與原IN_PROGRESS report保留，殘留dead-owner locks下次核對後可恢復移存，不盲刪。

最新source9B7D995...／EXE3421E580...未部署。最新全部unittest225 tests、234.412秒、OK、0SKIP。R2 actualEXE全334/303/10計畫保留；停止後report實際1/10完整港月：HWLN January已六檔完成，active February partial（3日cache），不是0/10。剩餘9港月與全部60的內容、同root raw/selector/default/引用/native/視覺及第二輪全部資料一致性尚未完成，UNVERIFIED。31日cache完整前嘗試啟動selector被guard拒絕（當時29日），故本輪未啟動該oracle，不假報PASS。

Source/copy v2已670/670 MATCH（fresh80/resumed590、endinginventory與manifest無變，詳下）；latest9B7人工ledger五份有值replay完成4800日/方位cells0diff，A/B小總表四圖pixel EXACT_EQUAL；Integrated仍不一致。minor/majorFont兩次隔離實驗皆未實際修好，不能稱根因已確證。Latest全部五份native/style及60文件仍UNVERIFIED。追查Normal/font0/style繼承是未驗假說，5%後沒有再做。優先交接 `visual_revalidation_checkpoint_20261007_2339.md`、`RENDERER_FROZEN_20261007_2340.md`、本文件及R2實際run_report。

恢復：先live quota與使用者指示；驗證候選/來源hash與own-process0，保存恢復自己的dead locks/flag；先解Integrated native font（不能改raw/ref/expected），若source變動須重新build/fullsuite及全部334新empty root。完整驗證兩輪all334/303/10，不以January或舊版本快照代填。安裝版仍v1.5.4，尚不能正式重新驗收。

狀態：`RESUMED_BY_USER_20261007_2305`。23:05依使用者再次明示恢復，live primary剩99%、weekly剩84%；繼續遵守任一適用視窗剩<=5%時停止。17:23先读剩6%，下一次讀值越過5%至3%的歷史暫停證據仍保留（不能聲稱精確停在5%）；交接為 `B:\AIS\work\ais32_full_regression_20261005\PAUSED_20261007_1725_USAGE_GATE.md`。原始 AIS 與正確 reference 保持唯讀；所有運行均使用隔離副本。最新候選原始碼926EE920...尚未build及全量驗收，現用EXE仍v1.5.4。沒有commit/tag/public release。

23:11核對沒有本工作殘留程序。停止後實際report證明E498候選HWLN January六檔已完成（1/10），active February partial；原交接0/10誤讀已更正。所有partial/stage/rollback保留，不充當最新版驗收。此次先完成最新字體／刻度原生比對與source/copy剩80項，再freeze/build新版並重跑全部334適用檔、303日、10港月。

23:15僅將resource位於上述舊run HWLN February且dead owner PID71752的20個lock及quota_pause.flag可恢復移至 `resume_20261007_2315_orphan_locks`。沒有刪除原始／reference／partial／rollback。Source verifier使用的是OS advisory lock，程序停止後自動釋放，不需刪JSON resource lock。

### 23時新增實際證據與稽核補強

- 5份人工reference的統計barChart皆為gapWidth219、overlap-27；既有writer未指定而用default150/0。這會影響柱寬／群距，屬共通renderer缺陷，不是港口特殊case。
- 最新926的整合圖XML雖寫0.0000/sourceLinked1，native仍讀取源格0.000或General。驗證需觀察Excel實際解析值；修正source/helper number formats與theme/lang的候選仍需native重驗。
- 新 `scripts/audit_default_workbook_content.py` / 7負例及正例tests PASS：逐個default month將cache（必須另有same-root raw/selector oracle證據）連接Modern總表、每日自動/Final、完整三scope決策台帳、5Historical summaries、獨立sample統計與所有chart numeric-value caches。這是序列化／決策傳遞驗證，不以cache自己冒充独立selector expected。對partial scope或cache日期缺漏fail closed。
- 新gate於舊E498 HWLN January全部31日／6workbooks實測：Modern/ledger/全部5日值與統計一致，但A full/small chart2 NW的MAX/MIN/平均cache為0.0，而AJ30/AK30/AL30實際皆blank。6項blank-to-zero差異證據保留 `default_content_E498_HWLN_202601_20261007.json`，修producer資料快取根因，不改正確expected。雷達category文字以c:numRef且無numCache是XlsxWriter序列化慣例，未誤報成數值cache錯。
- EXE全量wrapper追加拒絕非空output root，避免重用cache或覆寫舊run_report却標成fresh。5tests PASS，保留正常正式工具CLI/GUI使用流程。

### 最新凍結候選與23:28全量啟動

`ais_monthly_app.py` SHA256 `599A09E50B8459E4B6D0DB9A765BD53FDD00D34DAFE759B217F10B907FE71A68`。本輪已修上述cache/間距/theme/lang/原生numberformat及reference表頭格式；修正一個synthetic測試原來把無資料當0的錯誤期待，未修改任何人工／golden數字expected。完整unittest已完成223 tests、230.367秒、OK、0 SKIP，log `v155_final_renderer_suite_20261007_2325.stderr.log`。此數量包含新default內容與root保護tests，不合併舊局部測試虛增數量。

新EXE `dist_candidate_v155_final_renderer_20261007\AIS_32方位月報工具.exe`，v1.5.5、14,789,366 bytes、SHA256 `17A20855D22D0AF6A950394820D6FF8BFD276287241838FDE848D211E3D07912`，由凍結599A建置。仍是候選，未部署，安裝v1.5.4不動。建置使用既有dependencies、所有absolute paths，log `v155_final_renderer_build_20261007_2325.stderr.log`。

最新actualEXE全量fresh已啟動於新不存在root `safe334_packaged_v155_final_renderer_fresh_20261007`，全部334／303／10、workers3，不含 --only／daycap／process timeout／舊cache。精確launcher與時間在即時工具輸出；stdout/stderr為 `v155_final_renderer_packaged_fresh_20261007.*.log`。完整計畫及進度保留run_report.json；即使execution PASS也須same-root raw census、independent selector、default propagation、60資料/引用/明細、native+visual、完整再跑一致性各自通過，未完成保持UNVERIFIED。

23:32:04因新的native視覺FAIL證據停止此599A EXE run（0/10完整港月、HWLN January partial），核對創建時間／command後僅停自己的launcher8036/actual46160及EXE49796/38204/42368/48380/40836；全部partial與原IN_PROGRESS checkpoint保留。這不是運行時間限制，也未跳過資料：599A的Integrated minorFont仍有29 supplemental scriptfaces，其Hant覆寫Calibri，而reference只有三個primaryfaces，native字型及尺寸仍不同；統計xaxis的default outside tick/BFBFBF與ref none/D9D9D9也未修。須新build與新empty root重新全部334。599A A/B radar已actual pixel EXACT_EQUAL，不能擴張為全部chart視覺PASS；原born全suite223證據保留不追認下一source。

Cross-review又重現default cache gate可把兩份空chart XML判為([],0)無issues；已補每個chart所有series必須有cell-backed numRef、至少一series，以及stats2bar+1line契約。缺圖／literal-only不能以零檢查數當PASS，新增負例。所有正式驗收仍須獨立引用範圍／series計畫及native檢查，不以單一gate冒充全pipeline PASS。

### 23:34以後 R2：未完成項目仍UNVERIFIED

最新凍結source `9B7D99508F975F0EE2A068EC5AD283FD843B2CBFF1ED07F7A790DB7BF56C40CE`；補Integrated三primaryfaces/no supplemental、stats兩axes ticknone/D9D9D9 baseline0.75pt、A統計圖reference手動title精確位置。53 focused tests PASS；完整suite於23:34:28另起 `v155_renderer_r2_suite_20261007_2335.*.log`，結果以末尾實際數量為準，不能用前599A的223追認。default propagation helper最新8tests PASS（舊7+空圖負例）。

R2 EXE v1.5.5、14,788,858 bytes、SHA256 `3421E580A15AA14B5A2C8B4E6C4D4D27E8B870D7D09B72405B9AD6F0D4BFDB53`，`dist_candidate_v155_renderer_r2_20261007`；build成功，不部署。已從新empty root `safe334_packaged_v155_renderer_r2_fresh_20261007`啟動latest actualEXE全部334/303/10，完整suite同步收尾，不作final PASS。舊599A/R2進度不得混合；全部最新證明及第二次完整run仍待完成。

R2 launcher48972於23:36:45.598+08啟動；完整suite launcher17172於23:34:28.155+08啟動。R2三小總表actual native最新證據：A/B radar與stats四圖皆EXACT_PIXEL_EQUAL；Integrated仍radar13.18%／stats6.681%不同，native仍新細而refCalibri。故「minorFont supplementals是唯一根因／刪除已修好」必須撤回：XML已匹配primaryfaces卻未改native結果，根因仍UNVERIFIED，需續查chart-specific themeOverride/relationships與字型繼承，或隔離副本實測literal Calibri方式。不再單憑XML判定native font PASS。全5角色actual renderer replay以9B7另起generation，旧599A全5僅保留歷史證據。下一次source修改後仍須全部334重新run，不能追認本輪R2為新source驗收。

source/copy v2終檢已完成PASS（23:33:15）：670/670 MATCH、source/copy各335、起始/結尾inventory無缺多/異動、manifest不變。此恢復本次fresh80、有效metadata-resumed590；81hash attempts包含1筆NAS WinError59完整重試後MATCH，原錯誤與舊retry history保留。不是每項都本次fresh，也不是NAS原子snapshot。報告 `source_copy_integrity_v2_fresh_20261007.json` SHA256 `33A845ADB76796762B65735834F9B9B8E795B34E9D440043FC94154E10851D32`；方法/界線 `SOURCE_INTEGRITY_RESUME_20261007_2313.md`。verifier自有程序已自然退出。

恢復時確認原工作程序均不在運行。僅將PID 77556／65500、resource位於本稽核evidence root且owner已不存在的55個孤兒lock移至 `resume_20261007_1655_orphan_locks`，連同原quota_pause.flag保存，可追溯恢復；沒有刪除stage／rollback／raw／reference。圖表artifact-role與native geometry是本輪優先修復項目，前一B21F候選仍不可部署。

## 16:52 恢復後的即時證據（另於5%門檻保存新快照）

- 修正暫停紀錄中的進度誤讀：old Python resident snapshot實際只有HWLN Jan–Jun六個完整execution PASS；KLNG April因13:52:33 quota停止導致 `BrokenProcessPool`，該月24/30 cache，不是第七個完整PASS。保留原report及錯誤，沒有跳過該月；最新source freeze後全10港月必須重新完整處理。
- April獨立raw census已從19/30續完30/30（19個seal/stamps有效沿用＋11個本次fresh SAX），0issues PASS。Jan–Apr全120日完成；HWLN May全31日另起新audit，不設day cap。
- 新增 `scripts/independent_selector_audit.py`：不import production，對全部normalized rows以ascending sort＋binary-search range counts獨立重算所有32方向／logical_day、period_a、period_b的selected、bearing、rank、完整cluster count、over-cap count/max；六字段精確比對cache。先由獨立raw/spool census建立來源等同性，再由此oracle證明自動規則，不把cache自己當expected、不聲稱manual domain truth。完整輸出grid/公式/visual仍另外驗證。
- 花蓮January全31/31日該oracle PASS（2,976個selectors，17,856個數值字段，全部spool rows），報告 `independent_selector_HWLN_202601_20261007.json`。7個isolated tests PASS，含200組exhaustive-definition交叉比對、epsilon邊界、blank/zero、deep cluster、tie、錯誤配置與tampered numeric field。全303日尚待新版統一runroot。
- Chart geometry根因進一步確認：B full與B small的OOXML anchors相同，但B full自訂欄寬、small預設欄寬，使same twoCellAnchor呈現868.3499pt與675.7499pt；XlsxWriter的Calibri11定位metric亦不同於ExcelTW解析reference的metric。需保存role-specific cols/rows及exact EMU anchors，不以像素湊尺寸。其他新確認差異包含PlotArea auto/manual、legend9pt vs7.5pt、文字#595959 vs#666666，generic五artifact roles修復中。
- source/copy v2依相同checkpoint續跑；420個歷史content hash只在metadata/stamp有效時resume，250個尚待本次hash，ending inventory必須另外完成。NAS metadata讀取目前等待中，沒有改SMB、縮減檔案或用歷史670 PASS代填本輪。

## 17:12 新 renderer freeze 與新增漏檢修復

Production `ais_monthly_app.py` SHA-256為 `E498291F5BF2E3B8E6914C9A994A6F4C1CD35EB1D3CDAE31969E4FAB22C501C6`。不改numeric parser／selector。五種角色明確傳遞，保留所有有效series/range/cache/formulas/detail；圖表產生後以1MiB ZIP串流僅局部套用人工reference的cols/rows/EMU及字體。28/29/31日僅平移統計圖上下頂點；不依港別特殊補丁。五角色十圖原生位置與尺寸已獨立驗證差異0，但空30日fixture不證明有值图視覺等同。

雷達放射線根因：XlsxWriter預設會寫catAx/majorGridlines；人工A full/small有explicit此格線，B full/small與整合圖沒有。依角色A保留、B/Int禁用，全部valAx line依reference設noFill；不是把所有港口/角色硬改同一款。8 layout tests PASS，最後21 layout/HWLN tests PASS/0 SKIP。

`scripts/reproduce_june_reference.py` 修正稽核helper兩項漏檢：full與summary reference pairing原來反指（同日值grid相同而遮蔽錯誤）；compare_grid以前zip會截短比較，現在必須完整30×32且blank不等於0。2新增負例tests PASS；reference值未改。

Excel Chart.Export回傳true不證明有效PNG：前兩reference COM報告中Bfull radar實際0-byte，舊「10PNG」只代表10次API成功呼叫，不是10張有效圖片。原報告/零檔保存，QA已補PNG signature與實際解碼/正尺寸檢查、readonly retry，不以存在檔案當PASS。fullQA也移除SpecialCells blanketcatch，COM讀值或計數失敗列UNVERIFIED/FAIL，output path guards先於mkdir；3個人工小總表pilot為3/3 native PASS、6有效解碼PNG、hash不變。所有60產物仍待新全量驗證。

完整suite與candidate build於17:12:34並行啟動：tests launcher53168、build launcher75196；各python actual/children需在暫停時重新核對identity。logs `v155_geometry_all_tests.*.log`、`v155_geometry_build.*.log`；candidate新dist為 `dist_candidate_v155_geometry_20261007`。舊B21F不部署，已安裝v1.5.4不變。

首次build因version_info相對於specpath解析而FileNotFound，保留完整failure log；改用所有absolute paths，17:13:52 launcher73780重新build成功。新candidate EXE SHA-256 `A59323E440EF776E25BEBB1EFB1AC9813D4E864C335B98B5D17100360280FA2E`、FileVersion1.5.5、14,787,111bytes；未部署。

17:15:55由新EXE開始全部334適用檔／303日／10港月fresh正式CLI regression，root `safe334_packaged_v155_geometry_fresh_20261007` 啟動前不存在，不含舊cache、不用 --only／day cap／process timeout。launcher71560，logs `v155_geometry_packaged_fresh.*.log`。任一月份失敗必留report/error及追查，execution PASS與數值/Excel/視覺分開。

新validator獨立從五copied references建立role contract，不import producer HISTORICAL_LAYOUTS；37 scope/role/negative tests PASS。除anchor/grid/fonts外亦核對drawing→chart relationships，避免正確位置卻引用錯圖。完整suite開始後才追加這些tests，故原190附近suite即使PASS也必另跑最終全部suite。

修正後fullQA以全部5人工reference副本實測：1,892個公式全部native error-count0、10有效解碼PNG、5/5 SHA前後不變；UsedRange.HasFormula mixed的DBNull診斷及前失敗原報告保留。真值3份小總表用latest renderer重產，3×960日值grid0diff（1260人工候選回放），`visual_candidate_geometry_20261007_report.json`；native/像素正在核對，不是全60成果驗收。

## 使用者用量停止條件

使用者要求：「修正，並且在我 codex 用量剩餘 5% 時暫停」。以 Codex account usage tool 即時讀數為準，使用 `remaining = 100 - usedPercent`；任一適用、可取得的視窗剩餘 <= 5% 時保存目前狀態、停止 subagents／新的工作與運行，等待使用者恢復。未知／null 不當作 0 或 100%。不使用 reset credit 或自動恢復工作。

2026-10-07 開始讀數：5 小時剩餘 99%，每週剩餘 100%。各關鍵階段重新查核；到門檻前不得把未完成工作標 PASS。

## 既有證據更新

- `source_copy_integrity_final.json` 已完成 670/670 MATCH：NAS source 335、isolated copy 335，0 mismatch、0 missing/extra。是該次 hash 的歷史證據，保留原樣。
- v1.5.4 的 safe334 兩輪 10/10、60 XLSX、116,244,596 detail rows、69 tests 與 Excel-native 60/60、120 PNG 匯出結果保留，不能擴張成新的並行／提交／獨立原始資料驗收。
- v1.5.4 原技術報告的過度聲稱已修正：原兩表事件仍未查明；120 PNG 匯出不是全量視覺等同；人工 reference 方位舊檢查只讀兩段各100k；第二輪含 cache；五月 legacy 836 格差異不可漏報。

## 新重現的三項正式工具缺陷

重現證據：`B:\AIS\work\ais32_full_regression_20261005\duplicate_risk_synthetic_20261007\findings.json`。

1. Modern 與整合總表最終路徑同名：validate 接受，成功後只剩五個 unique paths，Modern decision ledger 被 Historical 覆蓋。
2. 兩個 delivery writer 交錯：固定 `.building.xlsx` 路徑被另一程序覆寫／提升；A 可回報成功，卻混入 B 的 full workbook 和 A 的 summary。
3. 首次 pipeline 的 Historical 建置失敗：Modern 已換新，而五份 Historical 保持旧批次；六份成果不是同一決策 snapshot。

共同修正：路徑碰撞前置拒絕、跨程序輸出／cache／spool 資源鎖、六檔先完成 staging 再整組提交／rollback。單獨 finalization 的五檔契約與正常 GUI／CLI 操作維持。

原本使用者那次「兩個表格」仍 `UNVERIFIED (original incident)`：新重現的風險不能直接認定為當次原因。

### 交叉審查追加重現

4. standalone 五檔 finalization 在 staging 後沒有再次檢查 `overwrite=False`，晚建立的同名成果仍被替換。必須共用六檔提交 helper 的終檢，不能只有開始時查一次。
5. 實際 reviewed workbook 輸入副本若改名成 Historical final／temporary／rollback，可被當成果覆寫或刪除；metadata 內 Analysis workbook path 並不是實際輸入 path。實際輸入也需納入保護與資源鎖。
6. GUI 確認等待期間來源異動，`build_processing_job` 在 try 外報錯且按鈕已 disable，沒有 worker/event 恢復；需在改變 UI 狀態前安全完成來源 preflight。

三項 before 負例均已重現：`followup_safety_before_20261007.log`。修正中，不以目前 113 tests 的舊快照宣稱追加缺口通過。

追加 GUI 授權缺口也已重現並修正：首次 start／finalize 沒有既有成果、未跳出確認時原本仍 `overwrite=True`；現預設 false，只有偵測到既有成果且同意後才 true。兩項 GUI callback 測試含六個 subtest 情境。production 最終 SHA-256 `1856093F3D079B767FC757BC0F06FE4F29BC2679BE2915CA1EC76033000487FD`；90 integration tests 有 87 PASS／3 external SKIP，外部環境補齊的完整 suite 另外執行中。正常 CLI 明示 `--overwrite` 未改。

## Source/copy verifier 可靠性修正

舊 checkpoint 只綁 manifest 路徑與scope、用null表示成功，會誤用同路徑新manifest／改動檔案，且舊failure永久略過。另直接覆寫JSON可能中斷損毀，讀取前一次stat無法確認hash期間穩定，`hashed_file_count`也把沿用結果計成本次hash。

修正契約：manifest內容與task fingerprint、成功的actual hash／size／file signature／time、failure重新嘗試、原子checkpoint與report、handle及path前後穩定性、開始／結尾inventory、fresh／resume分開計數及 `--fresh` 全檔新hash。舊 null checkpoint 不当作新版本成功證據。

交叉審查追加修正已由 31 個隔離測試確認：report／checkpoint／lock 不可寫入 source/copy root 或覆寫輸入；ending inventory 與 manifest 讀取後再次檢查檔案狀態；both scope 的 expected size/hash 兩側必須一致。這不是 NAS 原子 snapshot，不能保證 final metadata 檢查後仍沒有外部異動。

2026-10-07 13:23 啟動 `source_copy_integrity_v2_fresh_20261007.json` 全量 `--scope both --fresh`，335 source + 335 copy，不沿用歷史成功。初始 launcher PID 64324、實際 Python PID 80080；log `source_copy_integrity_v2_fresh.stdout.log`／`.stderr.log`。尚未完成，不得以歷史 670 PASS 代填本輪結果。

## 獨立來源稽核

新增獨立 ZIP/OOXML streaming reader，不import production parser或方位函式。全部334適用檔逐列重算資料清理counter、時間與period、訊息規則、32方位，再比對完整normalized spool的筆數、方向／period／message分布及逐列multiset fingerprint。此處不替人工selector造expected。

人工六月 reference 方向比對已完成全明細，不設 100k 上限。Period A 11,766,628 筆、Period B 11,144,440 筆，共 22,911,068 筆，全部 matched，0 blank／unrecognized／mapping mismatch；不代表自動 selector 等同人工判讀。證據為 `independent_manual_june_mapping_full.json`，副本 source/copy SHA-256 紀錄為 `manual_june_reference_copy_20261007/reference_copy_manifest.json`。

## 目前驗證進度

- 2026-10-07 完整 unittest：113 tests，203.167 秒，全部 OK、0 SKIP。已設定外部 June golden／contract／baseline 環境變數；完整 log 為 `v155_tests.stderr.log`。此結果是後續 verifier 路徑保護追加修正前的測試快照，修正後需再跑。
- 正式 production source 凍結 SHA-256：`E3899B6519A163108B36473C4B3E412E7ED753CBC7DEBE696F14759632873677`。fresh safe334 runner 已啟動，root 為 `safe334_run_v155_fresh_20261007`，不使用 v1.5.4 cache。
- verifier 交叉審查合成重現新問題：report 指向 task XLSX 時會在 hash PASS 後覆寫原檔；本次實際 report 在 source/copy root 外，未觸發。正在新增 root／task／manifest 路徑守衛與 bytes 保全負例。
- full scope validator 的空 jobs／部分已完成月份誤判 PASS 已修；runner 先寫完整 planned_jobs 與 IN_PROGRESS，全部完成才 COMPLETE，interrupt 保留 INTERRUPTED；validator 必須核對全計畫，舊 runner 報告需 `--source-root`。22 scope tests PASS。
- first fresh 全量程式在追加 GUI／finalization 修正前啟動，故其輸出是前一個記憶體 snapshot；數值 parser 未變，但最終 freeze 後仍須完整 10/10 重跑，不能只以 source 檔案目前 hash 追認前一輪。
- 獨立 raw 稽核逐個完整港月即時進行以與正式流程並行，最終仍須覆蓋全部 334 files／303 days／10 port-months，再核對 global plan、sealed entries 和 input stamps；沒有抽樣。首個 HWLN_202601 audit launcher PID 43488，報告 `independent_raw_v155_HWLN_202601.json`。
- 最終候選 EXE 已建置，版本資訊 1.5.5，檔案為 `dist_candidate_v155_20261007/AIS_32方位月報工具.exe`，SHA-256 `B21FECE0C6561B4092BEE0C6A27D55F1478E4492684F1731616C41DB1E9852A3`。建置沿用既有 PyInstaller 6.21.0／XlsxWriter 3.2.9／openpyxl 3.1.5，未 pip upgrade。
- 13:33 啟動實際 EXE 的完整 fresh 回歸，launcher PID 46060，root `safe334_packaged_v155_fresh_20261007`，log `v155_packaged_fresh.stdout.log`／`.stderr.log`。工具 `scripts/run_packaged_full_regression.py` 每月透過正式 CLI 參數啟動 EXE、核對 return code／六檔／實際 cache day count，先保存全部計畫。這只是 execution gate，仍需完整內容／獨立 census／Excel-native／視覺驗證。工具隔離 guard 的 3 tests PASS。
- 人工 June reference 已再由補強路徑保護後的 v2 工具重跑全 22,911,068 筆，`independent_manual_june_mapping_v2_full.json` 仍 0 mismatch／全部 PASS；不能擴張成 selector 等同。
- Native detailed style 工具正在 3 份既有 June replay 小總表與 2 份人工 reference 副本做 COM API smoke，並非全量驗收。自有 PowerShell PID 72236（13:33:02）、Excel PID 72020（13:33:03，automation），exec session 39877。暫停時先 Ctrl+C 使 finally 關閉不儲存；必要時只對這兩個核對過建立時間／command 的 PID 停止，不能 kill 使用者其他 Excel。
- `v155_final_tests.stderr.log`：154 tests 全部 OK、0 SKIP，231.501 秒。其後追加 runner／validator root 寫入保護 5 tests、month collector 10 tests、packaged wrapper 3 tests；全部測試再於 13:38 啟動，log `v155_frozen_all_tests.stderr.log`，launcher PID 72340，結果待完成。
- HWLN_202601 raw independent 已完成全部 31/31 日，0 issue，報告 PASS。HWLN_202602 raw independent 已啟動，launcher PID 49012；global collector launcher PID 22252，報告 `independent_raw_v155_ALL.json`，目前 31/303 日，IN_PROGRESS，不能全量 PASS。collector 只讀月份 reports，不會自動啟動其他月份；剩餘各月需逐一執行 raw audit。prefix 固定 `independent_raw_v155_`。

## 用量暫停時的運行與恢復備忘

### Native 比對新發現：不能追認最終圖表通過

5 本 COM smoke 完成：5/5 read-only 及 hash 不變、10 PNG、5,115 個 points、0 fatal。原100個 UNVERIFIED 是不存在的 PlotArea.Font probes；已保留原始報告及 `native_style_smoke_20261007_interpretation.md`。此外 Excel 可能回傳 Mixed／inheritance sentinel（例如 -2147483600 的 line weight/transparency），不能把這種 API 讀值當成實際負線寬或已解析視覺透明度。

實際 role-aware 比對發現追加視覺缺陷：

- 人工5份 reference 的全部10個圖表 title 在 OOXML 均為 `sz=1400`，native 回傳14 pt；生成 replay 均16 pt。已修共用 Historical renderer 的兩個 title font 為14，新增全五檔／十圖 title test，單項PASS（0.722秒）。新source SHA為 `11C3362D635F5BFA984D83D0F76125D3487F2FF0E6C2F4263C2A7E6DE9A5AFE5`。
- Period B **full** reference 的 statistics chart寬868.3499 pt，Period B **small summary** reference675.7499 pt；現共同renderer生成的小總表675.2999 pt，full也共用相同geometry，因此不能把一種scope layout推廣成所有artifact role。此role geometry缺陷尚未修復與重驗，`UNVERIFIED / OPEN`。
- 人工Period A full/small statistics widths646.9500 pt，generated small646.8000 pt；整合人工radar475.9500×450.7500 pt，generated476.7000×450.7500 pt。Period B人工radar height448.1251 pt，generated448.5000 pt。需保留這些非零差異，追查anchor/column metrics與role模板，不假定像素等同。

3份人工小總表已另copy-first/hash-match且native read-only後hash不變，目錄 `manual_june_small_reference_copy_20261007`。SHA：A `AECE6095BB0C925BD6A20919486EB70C61C4F3A91B846C504A728808AACAFB1D`；B `C849FD5A98591C83546CFC8867803C28A8DCC029A193F185103680545B1D801A`；整合 `7F7BC1CBBD51E417777F55A526DE13D69C98864AA5B69ABB706FFCC32F3074D7`。原始與copy均未變。

**先前「final freeze」與B21F候選EXE現為前一snapshot，不含14pt修正。**172 tests全部PASS/0SKIP（249.086s）的結果也是title追加修正前；新source只有新增title test驗證，仍需全部tests、build、全10月重新生成/內容/native驗證。正在跑的兩份fresh pipeline均為舊resident snapshot，不可以source檔目前hash追認最新版本。不要部署B21F候選，也不要以原本172 PASS關閉新視覺缺陷。

March raw audit已啟動，launcher PID65332，`independent_raw_v155_HWLN_202603.json`；numeric parser未被title變更修改。其結果只用於來源/census，不能證明最新renderer。

尚未到門檻；以下是停止／恢復所需的已識別程序，不是已停止宣告。

| 工作 | launcher PID | 已識別 actual/child | log / checkpoint |
| --- | ---: | --- | --- |
| 首輪 Python fresh 全10月 | 69896 | 77556 及其日期 workers | `v155_fresh_run.*.log` / run root report/cache |
| 最終 EXE fresh 全10月 | 46060 | 80800 / EXE 81456、33628（13:33）及其 workers | `v155_packaged_fresh.*.log` / packaged root |
| 新 source/copy fresh hash | 64324 | 80080 | `source_copy_integrity_v2_fresh_20261007.progress.json` |
| Feb raw audit | 49012 | 啟動後重新由 parent tree 解析 | 每日原子 `independent_raw_v155_HWLN_202602.json` |
| 全月 collector | 22252 | 重新查 parent tree | `independent_raw_v155_ALL.json` / `quota_pause.flag` |
| 全部 frozen tests | 72340 | 重新查 parent tree | `v155_frozen_all_tests.*.log` |
| COM smoke | 72236 | Excel 72020 | session 39877 / native smoke report |

5% 時需即時重新核對 PID 建立時間與 command，防 PID reuse；停止本工作程序及 descendants，不能以 process name blanket kill。保留 isolated stage／cache／spool／rollback 和自己建立的 lock，不把 partial workbook 計為 final。不刪原始或 reference。續跑須先確認停止與 lock 歸屬，保留殘留狀態後處理自己的 lock／rollback；不可直接跳過 incomplete gate 或把舊 null checkpoint 匯成新成功。

source/copy v2 恢復可不加 `--fresh` 以 metadata 驗證沿用實際成功 hash，報告必須分列 resumed 與本次 fresh；如果要求新完整 hash，另用新 report 與 `--fresh`。raw audit 恢復加 `--resume`，只接受相同 source/cache/spool stamps 與 sealed entries；global collector 用 `--expect-files 334 --expect-days 303 --expect-months 10`。整體 EXE 流程重開仍跑全部10月，已存在 cache 的部分不是新 raw parsing，必須在報告明示；最終 determinism 需要保留上一份60workbook快照再做全量比較。

13:50 後 checkpoint 查核：Python 第一snapshot已完成HWLN Jan–Jun 6個港月、0 execution fail；B21F EXE已完成HWLN Jan–Mar 3個港月、0 execution fail（不含新title修正）。獨立 census Jan–Mar全部90/303日PASS，April 30日audit已啟動、launcher PID75300，部分day已PASS，不算完成月份。fresh hash checkpoint為205 source MATCH＋205 copy MATCH，共410穩定hash，0 error/mismatch；還差260，final inventory gate尚未完成。這些數字為進度快照，暫停時需重新讀實際狀態。

## 待完成／接手順序

1. 完成三項輸出安全與verifier測試，讀回程式diff。
2. 配置既有June golden／contract／baseline，跑全部測試，外部check不以SKIP當PASS。
3. 新版對safe334從新輸出目錄重新完整分析，全部10港月、303日，不設max-days。
4. 全60檔內容／OOXML驗證；另做独立全來源與normalized對照。
5. 再跑完整流程檢查重現性，明示cache／fresh parsing的實際範圍。
6. Excel-native 全60重算／全部chart匯出、視覺與reference檢查；缺同來源reference的項目仍UNVERIFIED。
7. build v1.5.5 EXE並做封裝／安裝驗證；先備份已部署v1.5.4，不覆寫歷史release。
8. 完成最終技術報告、版本checksum、ZIP及可用性判定。

若用量達停止門檻，記錄已修改檔案、測試結果、運行PID／log、下一步；不以候選code取代已驗證EXE、不擅自重開後續工作。

## 必須保留的 UNVERIFIED

- 未有人工覆核decision ledger的domain判讀。六月auto仍427/3150雙有值cell不同，ledger回放等同不等於auto判讀等同。
- 五月partial legacy來源／selector：2檔1984cells、836差異（752雙有值差異、84空→值）。
- 全產物與人工reference逐圖視覺等同，需依實際可比reference及驗證結果更新。
- 原使用者重複兩表事件。
- 公開release provenance與EXE簽章。
