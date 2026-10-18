# GC-MS FID 分類網站 — Jack 回來看這裡

## 你離開前的進度

Claude Code 已經裝好了（你自己跑完的）。
網站所有程式碼也已經寫好，放在這個資料夾。

---

## 你回來要做的步驟

### Step 1：先在自己電腦測試網站能不能跑

開命令提示字元，進到這個資料夾：
```
cd C:\Users\User\Desktop\ai\claude\saf-research\FID分類\gcms-web
```

安裝 Python 套件（只需要做一次）：
```
pip install -r requirements.txt
```

啟動網站：
```
python app.py
```

打開瀏覽器，進入：http://localhost:5000

丟 FMK-ZSM-5_R1.xlsx 進去測試，應該會看到分類結果和驗證報告。

---

### Step 2：部署到 Render（公開網站）

**先申請帳號：**
- GitHub：https://github.com（如果還沒有的話）
- Render：https://render.com（用 GitHub 帳號登入）

**用 Claude Code 推上 GitHub：**

開命令提示字元，進到專案資料夾：
```
cd C:\Users\User\Desktop\ai\claude\saf-research\FID分類\gcms-web
claude
```

然後叫 Claude Code：
```
幫我把這個專案推到 GitHub，建立一個新的 repository 叫 gcms-fid-classifier
```

**在 Render 部署：**
1. 登入 render.com
2. 點「New +」→「Web Service」
3. 連結你的 GitHub，選 gcms-fid-classifier
4. 設定：
   - Environment: Python 3
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn app:app`
5. 點「Create Web Service」
6. 等 2~3 分鐘，拿到公開網址（格式像 https://gcms-fid-classifier.onrender.com）

---

## 檔案說明

```
gcms-web/
├── app.py              ← 後端主程式（Flask + 分類邏輯）
├── templates/
│   └── index.html      ← 前端網頁（拖曳上傳介面）
├── requirements.txt    ← Python 套件清單
├── Procfile            ← Render 部署設定
└── .gitignore          ← Git 忽略清單
```

---

## 如果遇到問題

直接告訴 Claude（這個對話框），說「遇到 XXX 錯誤」，我來處理。

---

## 工作區（2026-10-18 起）與學長五個工具的維護

- `/` 是分頁工作區（`static/suite/shell.html`），`/fid` 是 FID 分類頁；學長的五個工具是 `static/suite/*.html`，不是手寫的，是從他的單檔 `index.html` 用 `tools/build_suite.py` 產生的。
- **使用成本為零 token**：站上沒有任何 AI／Claude 呼叫，伺服器沒有 API 金鑰。`tools/audit_suite.py` 把關這件事，每次重建會自動跑，也可單獨跑 `python tools/audit_suite.py`；它會擋 AI API 字樣、`claude.<其他能力>`、不在白名單的外部網域、多餘的網路呼叫。
- 學長出新版時：`python tools/build_suite.py "<新的 index.html 路徑>"`。我們對他的程式只有四個小 patch（寫在 build_suite.py 的 PATCHES），對不上會直接報錯，代表他改到那段了，要人看。
- 三個函式庫（SheetJS、ExcelJS、Plotly）放在 `static/vendor/`，版本與雜湊記在 `VERSIONS.txt`，不再連外部 CDN。升級要換檔案並更新 VERSIONS.txt。
- 五個工具的數據存在使用者自己的瀏覽器（IndexedDB），不在伺服器；總覽頁有備份／還原，並顯示上次備份距今幾天。
- 還原整件事：`git revert 8fc788b` 之後的提交即可回到只有 FID 分類的版本（`/` 會變回分類頁）。
