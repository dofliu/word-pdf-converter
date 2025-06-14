# Word PDF 轉換與合併工具

這是一個功能完整的Word與PDF文件轉換工具，提供圖形化界面和豐富的功能選項。

## 主要功能

- **Word轉PDF**:
  - 支援多種轉換方法 (docx2pdf, COM自動化, LibreOffice)
  - 自動處理中文字型 (Windows/Linux/macOS)
  - 轉換進度顯示

- **PDF轉Word**:
  - 使用pdf2docx進行高品質轉換
  - 保留原始格式和內容

- **多文件合併PDF**:
  - 支援混合Word和PDF文件合併
  - 自動生成目錄頁
  - 添加頁碼功能
  - 自訂頁碼格式

- **圖形化界面**:
  - 使用PyQt5開發的跨平台界面
  - 進度條顯示轉換狀態
  - 錯誤訊息提示

## 系統需求

- Python 3.7+
- 必要套件:
  - PyQt5
  - python-docx
  - pdf2docx
  - docx2pdf
  - PyPDF2
  - reportlab

## 安裝方式

1. 克隆倉庫:
```bash
git clone https://github.com/dofliu/word-pdf-converter.git
cd word-pdf-converter
```

2. 安裝依賴套件:
```bash
pip install -r requirements.txt
```

## 使用方式

### 圖形化界面
```bash
python src/integrated_app.py
```

### 命令列使用
Word轉PDF:
```bash
python src/integrated_app.py --word input.docx --pdf output.pdf
```

PDF轉Word:
```bash
python src/integrated_app.py --pdf input.pdf --word output.docx
```

合併多個文件:
```bash
python src/integrated_app.py --merge file1.pdf file2.docx --output merged.pdf
```

## 授權資訊

本專案採用 MIT 授權條款。詳見 LICENSE 檔案。

## 貢獻指南

歡迎提交 Pull Request。請確保:
1. 遵循現有程式碼風格
2. 包含適當的測試
3. 更新相關文件
