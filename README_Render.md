# MOPS 金融股月自結損益追蹤系統 — Render 雲端部署指南

本專案已完整配置 Render 雲端主機所需的部署設定檔（包含 `render.yaml`、`app.py`、`requirements.txt`、`mops_self_profit.db` 與靜態資源），支援 100% 免費部署至 Render (render.com)。

---

## 🚀 部署方式一：標準 Web Service 部署（最推薦）

### 步驟 1：將程式碼推送至您的 GitHub 儲存庫
本專案已在本地完成 `git init` 與初始 Commit（`main` 分支）。請在您的 GitHub 建立一個新儲存庫（Repository，公開或私有皆可），並執行下列指令將專案推送到 GitHub：

```bash
# 1. 綁定您的 GitHub 遠端倉庫（請將 URL 替換為您的倉庫網址）
git remote add origin https://github.com/<您的GitHub帳號>/<儲存庫名稱>.git

# 2. 推送至 main 分支
git push -u origin main
```

---

### 步驟 2：在 Render 建立 Web Service
1. 登入 [Render 控制台 (dashboard.render.com)](https://dashboard.render.com/)。
2. 點擊右上角 **「New +」**，選擇 **「Web Service」**。
3. 選擇 **「Build and deploy from a Git repository」**，並點擊 **「Connect」** 連結剛剛建立的 GitHub 倉庫。
4. 填寫服務設定（或系統會自動代入）：
   - **Name**：`mops-financial-tracker`（或任意名稱）
   - **Region**：`Singapore`（新加坡）或 `Oregon (US West)`
   - **Branch**：`main`
   - **Root Directory**：留空（即根目錄）
   - **Runtime**：`Python 3`
   - **Build Command**：`pip install -r requirements.txt`
   - **Start Command**：`gunicorn app:app --workers 2 --threads 4 --timeout 120 --bind 0.0.0.0:$PORT`
   - **Instance Type**：選擇 **Free**（完全免費）
5. 點擊下方 **「Create Web Service」**。
6. 等候約 1~2 分鐘建置完成，Render 將提供一組專屬 HTTPS 網址（例如：`https://mops-financial-tracker.onrender.com`）。

---

## ⚡ 部署方式二：使用 Blueprint 藍圖一鍵部署

Render 支援直接讀取專案根目錄的 `render.yaml` 自動解析所有設定：
1. 將專案推送到 GitHub 後。
2. 於 Render Dashboard 點擊 **「New +」** -> **「Blueprint」**。
3. 選擇您的 GitHub 儲存庫，Render 會自動讀取 `render.yaml`。
4. 點擊 **「Apply」** 即自動完成所有建置與啟動！

---

## 📋 專案部署檔案說明

| 檔案 | 作用 |
| :--- | :--- |
| **`app.py`** | 雲端 WSGI 核心程式（相容 Gunicorn / wsgiref，內建容錯自動偵測 DB 與 HTML 路徑） |
| **`requirements.txt`** | Python 套件清單（`gunicorn`, `pandas`, `openpyxl`, `beautifulsoup4`） |
| **`render.yaml`** | Render Blueprint 自動部署規格設定檔 |
| **`index.html`** | 繁體中文前端完整單頁應用（含投資表、自結清單、快易查、資料來源） |
| **`mops_self_profit.db`**| 預載 SQLite3 快取資料庫（隨附 25 檔金融股歷年損益與指標資料） |
| **`assets/`** | 包含證交所公告快易查截圖等靜態資源 |

---

## 💡 Render 免費版 (Free Tier) 特性注意事項
- **休眠喚醒機制**：Render 免費方案在持續 15 分鐘無人連線造訪時，會自動進入睡眠待機狀態以節省資源；當有新使用者連線造訪時，約需 30~50 秒進行冷啟動喚醒，隨後即恢復全速運作。
- **HTTPS 安全加密**：Render 自動提供免費且永久續期的 SSL/TLS 證書，直接支援 HTTPS 安全連線。
