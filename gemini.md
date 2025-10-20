# 專案重構計畫：Word/PDF 轉換工具 (GUI & AI-Callable MCP Server)

## 1. 專案目標

本計畫旨在將現有的 Word 與 PDF 轉換及合併工具進行重構，使其具備以下兩種運行模式：
1.  **獨立 GUI 應用程式：** 維持現有基於 PyQt5 的桌面應用程式功能，提供使用者友善的圖形介面進行檔案轉換與合併。
2.  **AI 可呼叫的 MCP 伺服器：** 透過建立一個 Web API 服務，將核心轉換功能暴露出來，使其能夠被 AI 服務或其他自動化系統呼叫，並支援多雲部署 (Multi-Cloud Platform)。

## 2. 新架構概覽

新的架構將會清晰地分離核心業務邏輯、GUI 介面和 Web API 介面，實現程式碼重用和模組化。

```
+-------------------+       +-------------------+
|   GUI Application |       |   API Server      |
| (integrated_app.py)|       | (api_server.py)   |
|   (PyQt5)         |       |   (FastAPI)       |
+---------+---------+       +---------+---------+
          |                             |
          |        +-------------------+
          +-------->|   Core Converter  |<--------+
                   | (core_converter.py)|
                   | (Word/PDF Logic)  |
                   +-------------------+
```

## 3. 實施階段與詳細計畫

### 階段 1: 核心邏輯提取 (Core Logic Extraction)

**目標：** 將所有檔案轉換和合併的業務邏輯從 `src/integrated_app.py` 中分離出來，形成一個獨立的、不依賴於 GUI 的 Python 模組。

**具體步驟：**

1.  **建立新檔案：** 在 `src/` 目錄下建立一個新檔案 `core_converter.py`。
2.  **遷移轉換函數：**
    *   將 `convert_word_to_pdf(word_path, pdf_path)` 函數完整遷移到 `src/core_converter.py`。
    *   將 `convert_pdf_to_word(pdf_path, word_path)` 函數完整遷移到 `src/core_converter.py`。
    *   將 `setup_fonts()` 和 `CHINESE_FONT` 相關的字型設定邏輯遷移到 `src/core_converter.py`，確保轉換功能可以獨立運行。
3.  **遷移合併邏輯：**
    *   將 `MergePdfThread` 類別中的核心合併邏輯（例如 `_prepare_pdf_reader`, `create_toc`, `add_page_numbers`, `to_roman` 以及 `run` 方法中實際執行轉換和合併的部分）提取到 `src/core_converter.py` 中，作為獨立的函數或類別方法。
    *   例如，可以建立一個 `PdfMerger` 類別，包含 `merge_files(file_list, output_file, options)` 等方法。
    *   **重要：** 處理密碼輸入的部分 (`password_requested`, `wait_for_password`, `wrong_password`, `decrypt_error`) 是 GUI 特有的互動，這部分不應遷移到 `core_converter.py`。`core_converter.py` 中的合併邏輯應假設所有 PDF 檔案在傳入時已解密，或者在需要密碼時拋出異常，由上層呼叫者（GUI 或 API）處理。
4.  **更新 `src/integrated_app.py`：**
    *   從 `src/core_converter.py` 導入新的核心函數和類別。
    *   修改 `PdfToWordThread`, `WordToPdfThread`, `MergePdfThread` 類別，使其不再包含核心轉換邏輯，而是呼叫 `src/core_converter.py` 中對應的函數。
    *   確保 GUI 介面在呼叫這些核心功能時，仍然能正確處理進度更新、狀態顯示和錯誤訊息。
5.  **更新 `requirements.txt`：** 確保所有核心邏輯所需的第三方庫都已列出。

### 階段 2: API 伺服器實作 (API Server Implementation)

**目標：** 建立一個基於 FastAPI 的 Web API 服務，將 `src/core_converter.py` 中的功能暴露為 RESTful API 端點。

**具體步驟：**

1.  **建立新檔案：** 在 `src/` 目錄下建立一個新檔案 `api_server.py`。
2.  **安裝 FastAPI 及其依賴：**
    *   將 `fastapi` 和 `uvicorn` 添加到 `requirements.txt`。
    *   `pip install fastapi uvicorn`
3.  **實作 FastAPI 應用：**
    *   在 `src/api_server.py` 中初始化 FastAPI 應用。
    *   從 `src/core_converter.py` 導入核心轉換函數。
    *   **定義 API 端點：**
        *   **Word 轉 PDF：**
            *   `POST /convert/word-to-pdf`：接收一個 Word 檔案作為上傳，呼叫 `core_converter.convert_word_to_pdf`，並返回轉換後的 PDF 檔案。
            *   考慮使用 `UploadFile` 處理檔案上傳。
        *   **PDF 轉 Word：**
            *   `POST /convert/pdf-to-word`：接收一個 PDF 檔案作為上傳，呼叫 `core_converter.convert_pdf_to_word`，並返回轉換後的 Word 檔案。
        *   **PDF 合併：**
            *   `POST /merge/pdfs`：接收多個檔案（Word 或 PDF）作為上傳，以及合併選項（例如是否生成目錄、是否添加頁碼、頁碼格式、起始頁碼）。
            *   呼叫 `core_converter` 中對應的合併邏輯，並返回合併後的 PDF 檔案。
            *   需要處理臨時檔案的儲存和清理。
    *   **錯誤處理：** 實作適當的錯誤處理機制，例如檔案類型不符、轉換失敗等，並返回有意義的 HTTP 狀態碼和錯誤訊息。
    *   **安全性考量：** 雖然不在本次重構範圍內，但未來應考慮 API 認證和授權。
4.  **測試 API 服務：** 使用工具（如 Postman, curl 或 Python `requests` 庫）測試每個 API 端點，確保其功能正常。

### 階段 3: 容器化與多雲部署準備 (Containerization for MCP)

**目標：** 建立 Dockerfile，將 API 伺服器及其所有依賴打包成一個可移植的 Docker 映像，為多雲部署做準備。

**具體步驟：**

1.  **建立 Dockerfile：** 在專案根目錄下建立 `Dockerfile`。
2.  **Dockerfile 內容範例：**
    ```dockerfile
    # 使用 Python 官方映像作為基礎
    FROM python:3.9-slim-buster

    # 設定工作目錄
    WORKDIR /app

    # 複製 requirements.txt 並安裝依賴
    COPY requirements.txt .
    RUN pip install --no-cache-dir -r requirements.txt

    # 複製應用程式程式碼
    COPY src/ ./src/
    COPY ./.gitignore ./.gitignore # 複製 .gitignore 以便在容器內使用，如果需要

    # 暴露 FastAPI 服務的埠
    EXPOSE 8000

    # 運行 FastAPI 應用程式
    # --host 0.0.0.0 允許從外部訪問
    CMD ["uvicorn", "src.api_server:app", "--host", "0.0.0.0", "--port", "8000"]
    ```
3.  **建立 `.dockerignore`：** 在專案根目錄下建立 `.dockerignore` 檔案，排除不必要的檔案和目錄，以減小 Docker 映像大小（例如 `__pycache__`, `.git`, `build`, `dist`, `*.spec`, `test_files` 等）。
4.  **建構 Docker 映像：**
    *   `docker build -t word-pdf-converter-api .`
5.  **本地測試 Docker 映像：**
    *   `docker run -p 8000:8000 word-pdf-converter-api`
    *   確認 API 服務在本地 Docker 容器中正常運行。
6.  **多雲部署考量：**
    *   一旦 Docker 映像準備就緒，它可以被推送到任何容器註冊表（如 Google Container Registry, Docker Hub）。
    *   然後，可以使用各雲服務提供商的容器服務進行部署，例如 Google Cloud Run (無伺服器容器), Google Kubernetes Engine (GKE), AWS Fargate, Azure Container Instances 等。

### 階段 4: GUI 介面適應性調整 (GUI Adaptation)

**目標：** 確保 GUI 應用程式在核心邏輯分離後，仍然能夠正常運行並提供所有功能。

**具體步驟：**

1.  **測試 GUI 功能：** 在完成核心邏輯提取後，徹底測試 GUI 應用程式的所有功能（Word 轉 PDF、PDF 轉 Word、PDF 合併），確保沒有引入任何回歸錯誤。
2.  **優化 GUI 互動：** 考慮在 GUI 中加入一些提示，例如當轉換或合併操作需要較長時間時，可以顯示更詳細的進度或預計時間。

## 4. 後續考量

*   **錯誤日誌：** 在 `core_converter.py` 和 `api_server.py` 中實作健壯的日誌記錄機制，以便於問題追蹤和除錯。
*   **效能優化：** 對於大型檔案的轉換和合併，可能需要進一步優化效能，例如使用多進程或異步處理。
*   **安全性：** 對於 API 服務，未來應考慮加入 API 金鑰、OAuth2 等認證和授權機制。
*   **配置管理：** 將一些可配置的參數（例如臨時檔案路徑、字型路徑）從程式碼中提取出來，使用環境變數或配置文件進行管理。

---
