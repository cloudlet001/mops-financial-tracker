---
title: MOPS 金融股月自結損益彙整系統
emoji: 📊
colorFrom: blue
colorTo: indigo
sdk: docker
app_port: 7860
pinned: false
---

# 📊 MOPS 金融股月自結損益彙整系統 (雲端部署包)

本目錄包含完整的雲端部署設定，可直接免費部署至 **Hugging Face Spaces**、**Render.com** 或 **Koyeb**。

## 📦 包含檔案清單
- `web_app.py`：後端伺服器（已支援 `0.0.0.0` 與動態 `PORT` 環境變數）
- `index.html`：前端互動式網頁介面
- `mops_self_profit.db`：SQLite3 預設種子資料庫（隨專案上傳即可免等待秒讀歷史資料）
- `requirements.txt`：Python 套件清單
- `Dockerfile`：Docker 容器設定（相容 Hugging Face Spaces `7860` port 與 Render/Koyeb）
- `render.yaml`：Render.com 一鍵部署設定檔
- `Procfile`：Koyeb / Railway / Zeabur 啟動設定檔

---

## 🚀 方法一：免費部署至 Hugging Face Spaces（最推薦：免信用卡、2 vCPU + 16GB RAM）

1. 前往 [Hugging Face Spaces](https://huggingface.co/spaces) 登入帳號，點擊右上角 **「New Space」**。
2. 填寫設定：
   - **Space name**：例如 `mops-financial-profit`
   - **Select the Space SDK**：選擇 **Docker**（選 Blank 模板）
   - **Space hardware**：選擇 `CPU basic · 2 vCPU · 16 GB · FREE`
   - 點擊 **Create Space**。
3. 在 Space 頁面點選 **「Files」** -> **「+ Contribute」** -> **「Upload files」**。
4. 將本資料夾 (`web/`) 內的以下 5 個檔案拖曳上傳並點擊 **Commit changes to main**：
   - `README.md`
   - `Dockerfile`
   - `requirements.txt`
   - `web_app.py`
   - `index.html`
   - `mops_self_profit.db`
5. 等待約 1~2 分鐘建置完成（狀態變為 `Running`），即可直接透過您的公開網址使用！

---

## 🚀 方法二：免費部署至 Render.com（GitHub 自動同步部署）

1. 將本資料夾 (`web/`) 內的檔案推送到您的 GitHub 儲存庫（Repository）。
2. 登入 [Render.com](https://dashboard.render.com/)，點擊 **「New +」** -> **「Web Service」**。
3. 連結您的 GitHub 儲存庫：
   - **Runtime**：`Python 3`
   - **Build Command**：`pip install -r requirements.txt`
   - **Start Command**：`python web_app.py`
   - **Instance Type**：選擇 **Free**
4. 點擊 **Deploy Web Service**，完成後即可取得 `https://xxxx.onrender.com` 免費公開網址。
