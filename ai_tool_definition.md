# AI 工具定義：Word/PDF 轉換器 (MCP Server)

本文件定義了 Word/PDF 轉換器 MCP 伺服器提供的功能，以便 AI 模型能夠理解並調用這些功能。這些定義基於 FastAPI 應用程式 (`src/api_server.py`) 暴露的 RESTful API 端點。

AI 模型可以使用這些定義來：
1.  識別使用者意圖是否與文件轉換或合併相關。
2.  選擇最合適的工具功能來執行任務。
3.  提取必要的參數並以正確的格式調用 API。
4.  處理 API 的回應並向使用者呈現結果。

## 1. 工具名稱與描述

*   **工具名稱 (Tool Name):** `pdf_converter` (或 `document_converter`)
*   **工具描述 (Tool Description):** 一個強大的文件轉換和合併工具，能夠將 Word 文件轉換為 PDF，將 PDF 文件轉換為 Word，以及將多個 Word 或 PDF 文件合併為單一 PDF 檔案。

## 2. 可用功能 (Functions)

以下是 `pdf_converter` 工具提供的功能及其詳細定義。這些定義通常會以 JSON Schema 格式提供給支援函數調用的 AI 模型。

### 2.1. 功能：Word 轉 PDF (convert_word_to_pdf)

*   **功能名稱 (Function Name):** `convert_word_to_pdf`
*   **功能描述 (Function Description):** 將一個 Word 文件 (.docx 或 .doc) 轉換為 PDF 格式。
*   **API 端點 (API Endpoint):** `POST /convert/word-to-pdf`
*   **參數 (Parameters):**
    *   `file` (類型: `file` / `binary`): 必需。要轉換的 Word 文件。AI 模型需要提供檔案的內容或可訪問的 URL。
*   **回應 (Response):** 轉換後的 PDF 檔案 (application/pdf)。

### 2.2. 功能：PDF 轉 Word (convert_pdf_to_word)

*   **功能名稱 (Function Name):** `convert_pdf_to_word`
*   **功能描述 (Function Description):** 將一個 PDF 文件轉換為 Word (.docx) 格式。
*   **API 端點 (API Endpoint):** `POST /convert/pdf-to-word`
*   **參數 (Parameters):**
    *   `file` (類型: `file` / `binary`): 必需。要轉換的 PDF 文件。AI 模型需要提供檔案的內容或可訪問的 URL。
*   **回應 (Response):** 轉換後的 Word 文件 (application/vnd.openxmlformats-officedocument.wordprocessingml.document)。

### 2.3. 功能：合併 PDF (merge_pdfs)

*   **功能名稱 (Function Name):** `merge_pdfs`
*   **功能描述 (Function Description):** 將多個 Word 文件 (.docx/.doc) 或 PDF 文件合併為一個單一的 PDF 檔案。支援生成目錄和添加頁碼。
*   **API 端點 (API Endpoint):** `POST /merge/pdfs`
*   **參數 (Parameters):**
    *   `files` (類型: `array` of `file` / `binary`): 必需。要合併的 Word 或 PDF 文件列表。AI 模型需要提供多個檔案的內容或可訪問的 URL 列表。
    *   `generate_toc` (類型: `boolean`, 預設值: `false`): 可選。是否在合併後的 PDF 中生成目錄。
    *   `add_page_numbers` (類型: `boolean`, 預設值: `false`): 可選。是否在合併後的 PDF 中添加頁碼。
    *   `page_number_format` (類型: `string`, 預設值: `"數字"`, 可選值: `"數字"`, `"羅馬數字"`): 可選。頁碼的格式。
    *   `start_page_number` (類型: `integer`, 預設值: `1`): 可選。頁碼的起始數字。
*   **回應 (Response):** 合併後的 PDF 檔案 (application/pdf)。

## 3. AI 模型如何使用這些定義

AI 模型（例如，支援函數調用的 LLM）會接收這些工具定義。當使用者提出與文件轉換或合併相關的請求時，AI 模型會根據其內部邏輯和這些定義來判斷：

1.  **是否需要使用工具：** 例如，如果使用者說「請將我的報告轉換為 PDF」，AI 會意識到它需要 `pdf_converter` 工具的 `convert_word_to_pdf` 功能。
2.  **選擇哪個工具功能：** 根據使用者請求的細節（例如，轉換類型、是否需要合併），AI 會選擇最匹配的功能。
3.  **提取參數：** AI 會從使用者提示中提取必要的參數值（例如，輸入檔案的名稱或內容，以及 `generate_toc` 等選項）。
4.  **調用 API：** AI 會構造一個對應的 HTTP 請求，將提取的參數發送到 FastAPI 伺服器的相關端點。
5.  **處理回應：** AI 會接收 API 的回應（例如，下載連結或錯誤訊息），並以使用者友好的方式呈現給使用者。

## 4. 檔案處理考量

由於 API 端點期望接收檔案內容，AI 模型在調用這些功能時，需要能夠以 `multipart/form-data` 的形式提供檔案。這通常意味著：

*   **AI 代理：** 如果是 AI 代理，它可能需要一個機制來從使用者那裡獲取檔案（例如，使用者上傳到一個臨時 URL，或代理直接從使用者介面接收檔案內容）。
*   **回應處理：** API 返回的是檔案內容，AI 代理需要將其作為下載提供給使用者，或根據後續指令進行進一步處理。

---
