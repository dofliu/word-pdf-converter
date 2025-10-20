# 使用 Python 官方映像作為基礎
FROM python:3.9-slim-buster

# 設定工作目錄
WORKDIR /app

# 複製 requirements.txt 並安裝依賴
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 複製應用程式程式碼
COPY src/ ./src/

# 暴露 FastAPI 服務的埠
EXPOSE 8000

# 運行 FastAPI 應用程式
# --host 0.0.0.0 允許從外部訪問
CMD ["uvicorn", "src.api_server:app", "--host", "0.0.0.0", "--port", "8000"]
