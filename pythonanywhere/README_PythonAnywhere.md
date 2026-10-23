# 📊 MOPS 金融股月自結損益彙整系統 — PythonAnywhere 部署專用版

本資料夾 (`pythonanywhere/`) 已針對 **PythonAnywhere (免費版 Beginner / 付費版皆相容)** 完成 WSGI 打包與設定。

---

## 📁 檔案清單
1. **`flask_app.py`**：PythonAnywhere 專用主程式（同時匯出 `app` 與 `application` WSGI 介面，完美對應 PythonAnywhere 的 Flask 預設設定）。
2. **`index.html`**：前端網頁介面。
3. **`mops_self_profit.db`**：SQLite3 資料庫（已預載 **25 檔金融股 11509 自結清單** 及 **全部 25 檔 114、115 年度完整月自結損益資料**，上傳後立即可用）。
4. **`requirements.txt`**：相依套件清單（PythonAnywhere 預設環境已內建 `pandas`、`beautifulsoup4`、`openpyxl`，無須額外安裝即可直接運行）。
5. **`mops_pythonanywhere.zip`**：將上述檔案打包好的壓縮檔（可選擇單檔上傳解壓縮，或直接上傳前 3 個檔案）。

---

## 🚀 PythonAnywhere 三步驟上傳與啟動教學

### 步驟 1：建立 Web App
1. 註冊並登入 [PythonAnywhere](https://www.pythonanywhere.com/)。
2. 點擊上方選單的 **「Web」** 分頁 ➔ 點擊 **「Add a new web app」**。
3. 點擊 **Next** ➔ 選擇開發框架：**「Flask」** ➔ 選擇 Python 版本（例如 **Python 3.10** 或 **Python 3.11**）➔ 路徑維持預設的 `/home/您的帳號/mysite/flask_app.py` ➔ 點擊 **Next** 完成建立。

### 步驟 2：上傳本資料夾內的 3 個核心檔案
1. 點擊上方選單的 **「Files」** 分頁。
2. 點擊左側 Directories 下的 **`mysite/`** 資料夾（進入 `/home/您的帳號/mysite`）。
3. 點擊右側黃色按鈕 **「Upload a file」**，依序上傳本資料夾內的這 **3 個檔案**（會自動覆蓋原本預設空白的 `flask_app.py`）：
   - `flask_app.py`
   - `index.html`
   - `mops_self_profit.db`

   *(💡 進階快速法：您也可以只上傳 `mops_pythonanywhere.zip` 到 `mysite/`，然後點開 Bash console 輸入 `cd ~/mysite && unzip -o mops_pythonanywhere.zip` 一鍵解壓縮！)*

### 步驟 3：點擊 Reload 立即上線！
1. 回到上方選單的 **「Web」** 分頁。
2. 點擊畫面最上方綠色大按鈕 **「Reload 您的帳號.pythonanywhere.com」**。
3. 點擊上方的網址（`https://您的帳號.pythonanywhere.com`）即可開啟您的雲端版系統！

---

## 💡 關於 PythonAnywhere 免費版 (Free Tier) 的貼心設計
- **永久硬碟空間**：PythonAnywhere 免費版提供 512MB 永久磁碟，`mops_self_profit.db` 放在 `mysite/` 內**永遠不會因為休眠或重啟而消失**。
- **完整預載資料**：隨附的 `mops_self_profit.db` 已內建全部 25 檔金融股的 `11509` 清單與 `114`、`115` 兩年度完整月自結損益表，在「🗄️ 資料庫優先」模式下可 **0 秒瞬間讀取全部 25 檔資料並匯出 Excel / CSV**。
- **日後如何更新下個月的資料？**
  由於 PythonAnywhere 免費版對外部非白名單網站有連線限制，日後每個月公布新自結損益時，您只需要在自己的電腦跑一次更新，再把電腦裡最新的 `mops_self_profit.db` 上傳覆蓋到 PythonAnywhere 的 `mysite/` 目錄，雲端網頁就會同步擁有最新月份數據！
