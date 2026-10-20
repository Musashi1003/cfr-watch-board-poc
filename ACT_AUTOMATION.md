# 每週 ACT 自動更新

本機 Outlook 的 Gaming / PC 同週 CFR 附件，交由 `act_sync.py` 使用與 Streamlit 相同的 `act_snapshot.py` 規則計算。現有 CFR 達標率郵件的程式、內容、收件者與寄送狀態保持獨立。

## 自動執行

沿用 Windows 排程 `CFR Weekly Mail Watcher`：工作日 08:00 起，每 10 分鐘檢查，持續 12 小時。原有排程只在使用者登入、電源條件允許時執行。啟動器先執行原寄信程式，再執行 ACT worker；郵件已寄出或處於暫停期時，ACT 仍可獨立重試。

* 本機需開機、登入且 Outlook 可正常讀信；GitHub 使用現有 Git Credential Manager 登入，不需要把 Token 放入程式。
* 兩種來源必須具有相同週次；檔名所辨識的類別、最新 raw data 週次與信件週次需相符。
* 補跑時依 GitHub 最新 ACT 週次，從仍可找到的共同週依序處理，每次最多一週；缺少前週時停止，不跳週。
* 既有 ACT 與重新計算值不同、累計值倒退、缺少必要欄位或資料鍵重複時，停止更新並保留原始資料。
* Git 使用隔離的暫存 checkout，只提交 `ACT table.xlsx`；不會提交工作目錄的其他變更。推送遇到遠端衝突時，下次排程重新讀取最新版本。
* 每次成功都重新 fetch GitHub ACT 檔，核對整份檔案 SHA-256。成功是 GitHub 持久化證明，不代表已驗證 Streamlit Cloud 部署或登入後畫面。
* 同週來源信件與規則未變且已驗證時，略過 Excel 解析與 GitHub 操作。

## 獨立通知

只寄給 `wu.james@inventec.com`，無副本、密件副本或原 CFR 郵件收件者。

* 成功：`[ACT 更新成功] WKxxxx`，附筆數、時間、GitHub 版本與讀回核對結果。同週相同資料不重複寄送。
* 連續三次失敗：`[ACT 更新失敗] WKxxxx`。每週失敗事件只提醒一次，後續每次排程仍會重試。
* 失敗通知寄出後恢復：`[ACT 補更新成功] WKxxxx`。
* 通知寄送失敗會保留待寄狀態，下次重試，不會取消已完成的 ACT 更新。Outlook `Send` 與本機狀態寫入無法形成同一交易，若電腦恰好在寄出後、記錄前中斷，極少數情況可能重寄通知。
* 關機或整個排程未啟動時，本機無法主動寄失敗通知；下一次排程啟動後再補跑。不包含雲端逾期監控。

## 查看結果與手動補跑

雙擊 `Run_ACT_Update.cmd` 可手動執行，或執行 `run_act_sync.ps1`。此操作會依正常規則更新 GitHub 並寄出獨立 ACT 通知。

本機狀態位於 `.act_sync/`（不提交 GitHub）：

* `status.json`：最近狀態、最後成功週次、檔案雜湊、GitHub 版本、通知時間或待寄錯誤。
* `history.jsonl`：每次完成核對的紀錄。
* `act_sync.log`：執行與重試紀錄，採輪替儲存。

試跑命令 `py act_sync.py --dry-run --week 2641` 只讀取 Outlook / GitHub 並試算，不發布、不寄信。增加 `--baseline "ACT table.xlsx"` 可用本機 ACT 表比對。兩份附件與暫存 Git checkout 在執行後清除；保留來源檔名與 SHA-256 供追查。

## 停用與維護

只要從既有 `run_cfr_mailer.ps1` 移除尾端 ACT 區塊，即可停用 ACT 同步而保留 CFR 寄信。安裝前啟動器備份放在原寄信資料夾，檔名 `run_cfr_mailer.ps1.before_act_sync_20261020.bak`。

ACT 計算規則由網站及 worker 共用；修改規則時需執行 `tests/test_act_sync.py`、`tests/test_act_table_workbook.py`，以真實來源試跑比對，並更新 `RULE_VERSION`。遇到同週來源修訂或新開案年度初始化，先確認差異再維護，程式不會直接覆寫正式 ACT。
