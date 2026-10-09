# gunicorn 啟動時會自動讀取工作目錄下的 gunicorn.conf.py。
#
# Render 不讀 Procfile（那是 Heroku 的東西）——實際執行的是 Render 後台
# Settings → Start Command。若那裡只寫 `gunicorn app:app`，Procfile 裡的
# --workers/--max-requests/--timeout 全部不會生效。把設定放在這個檔案，
# 不論 Start Command 怎麼寫都會套用（命令列有明確給的參數才會覆蓋這裡）。
#
# 免費方案 512MB：只開 1 個 worker（多開一個就多一份 pandas+openpyxl 基線 ~100MB）。

workers = 1
threads = 2                 # 同一個 worker 內的執行緒共用記憶體；PubChem 等待是 IO，不佔 CPU
timeout = 180               # 單一請求上限（含 PubChem 查詢）；預設 30 秒會 WORKER TIMEOUT
graceful_timeout = 30
max_requests = 50           # 處理 50 個請求後重啟 worker，清掉長期累積的碎片
max_requests_jitter = 10
