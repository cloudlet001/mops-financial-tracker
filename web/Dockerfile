FROM python:3.11-slim

WORKDIR /app

# 安裝 Python 相依套件
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 複製程式碼與預設 SQLite3 種子資料庫 (mops_self_profit.db)
COPY web_app.py index.html ./
COPY mops_self_profit.db ./mops_self_profit.db

# Hugging Face Spaces 預設使用 7860 port；Render/Koyeb 會自動注入 PORT 環境變數
ENV PORT=7860
ENV HOST=0.0.0.0
ENV PYTHONUNBUFFERED=1

# 開放非 root 權限寫入 SQLite 資料庫 (相容 Hugging Face Spaces user 1000)
RUN chmod -R 777 /app

EXPOSE 7860

CMD ["python", "web_app.py"]
