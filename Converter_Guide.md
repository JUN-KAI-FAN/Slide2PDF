# Slides to PDF 轉換工具

此工具用於將 PPTX/ODP 轉為 PDF，解決 LibreOffice 轉檔常見的字體擠壓問題，並自動移除浮水印與修正 PDF Metadata 標題。

### 1. 安裝環境

#### **安裝 Python**
1.  前往 [Python 官網下載頁面](https://www.python.org/downloads/windows/)。
2.  下載最新穩定版 (Latest Stable Release) 的 Windows Installer。
3.  執行安裝程式，**務必勾選「Add Python to PATH」** 再點擊 Install Now。
4.  安裝完後，開啟 cmd 輸入 `python --version` 即可確認是否安裝成功。

#### **安裝必要套件**
執行以下指令安裝轉檔與 PDF 處理引擎：

```bash
pip install aspose-slides pymupdf
```

*若安裝最新版後執行失敗，可參考已知穩定的環境版本 (Python 3.12.3)：*
```bash
pip install -r requirements.txt
```

---

### 2. 檔案配置

將 GitHub 儲存庫中的以下檔案下載並放置在同一個資料夾：

*   **SlidesToPDF.py**：核心轉檔與標記清理邏輯。
*   **Convert.bat**：Windows 啟動器（若需手動建立，內容如下）：

```batch
@echo off
python "%~dp0SlidesToPDF.py" %*
pause
```

---

### 3. 使用與資安說明

*   **執行方式**：直接將 PPTX/ODP 檔案拖到 `Convert.bat` 上即可產出 PDF。
*   **本地處理**：轉檔邏輯完全在本地 CPU 執行，不涉及任何雲端上傳，可安全處理敏感文件。
*   **高品質產出**：採用 `aspose-slides` 引擎確保排版與字體間距正確，並由 `pymupdf` 於指令層級精確移除浮水印標記，不影響底層內容。
