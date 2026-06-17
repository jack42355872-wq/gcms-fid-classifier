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
