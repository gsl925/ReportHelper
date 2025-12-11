# Report Helper

快速上傳影像或 PDF，進行 OCR，抽取關鍵句並依 STAR 原則產生可匯出為 PPT 的內容範本。

## 快速開始

1. 建立虛擬環境並安裝依賴
python -m venv .venv
source .venv/bin/activate  # 或 Windows: .venv\Scripts\activate
pip install -r requirement

好的，這是一個專業且清晰的 `README.md` 文件，您可以直接將它與您的專案檔案放在一起。

---

# 報告整理小幫手 v16.0 (Ollama 整合版)

這是一個桌面應用程式，旨在自動化從各種來源（圖片、Email、文字檔）提取資訊，並利用本地大型語言模型 (Ollama) 進行分析，最終一鍵生成結構化的 PowerPoint (PPTX) 報告。

## 功能亮點

-   **多樣化輸入**：支援拖曳、上傳、剪貼簿貼上等多種方式輸入資料。
-   **智慧文字辨識 (OCR)**：自動從圖片中提取文字內容（支援繁中、簡中、英文）。
-   **Email 解析**：可直接讀取 Outlook 的 `.msg` 檔案，並提取寄件人、主旨和內文。
-   **本地 AI 整合**：直接與本地運行的 [Ollama](https://ollama.com/) 連接，無需將資料上傳至外部伺服器，確保資料隱私與安全。
-   **一鍵生成報告**：從文字提取到 AI 分析，再到 PPT 生成，整個流程高度自動化。
-   **智慧排版**：自動識別報告中的 STAR 原則（情境、任務、行動、結果）關鍵字，並在 PowerPoint 中生成具有階層結構的專業版式。
-   **現代化介面**：採用 `ttkbootstrap` 打造美觀且易於使用的圖形介面。

## 系統需求

-   **作業系統**：Windows / macOS / Linux
-   **Python 版本**：3.8 或更高版本
-   **Tesseract OCR 引擎**：用於圖片文字辨識。
-   **Ollama**：用於本地 AI 分析。

## 安裝與設定

### 1. 安裝 Tesseract OCR

Tesseract 是本工具進行圖片文字辨識的核心引擎。

-   **Windows**:
    -   從 [Tesseract at UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki) 下載安裝程式。
    -   **重要**：在安裝過程中，請務必勾選 `Additional language data`，並選擇 `Chinese - Traditional` 和 `Chinese - Simplified`。
    -   安裝完成後，請將 Tesseract 的安裝路徑加入系統的 `Path` 環境變數中（例如 `C:\Program Files\Tesseract-OCR`）。

-   **macOS**:
    ```bash
    brew install tesseract
    brew install tesseract-lang
    ```

-   **Linux (Debian/Ubuntu)**:
    ```bash
    sudo apt update
    sudo apt install tesseract-ocr
    sudo apt install tesseract-ocr-chi-tra tesseract-ocr-chi-sim
    ```

### 2. 安裝與設定 Ollama

Ollama 讓您可以在本機端輕鬆運行大型語言模型。

1.  前往 [Ollama 官網](https://ollama.com/) 下載並安裝適合您作業系統的版本。
2.  安裝完成後，打開終端機或命令提示字元，下載您想使用的模型。建議使用 `llama3`：
    ```bash
    ollama pull llama3
    ```
3.  請確保 Ollama 服務在背景運行。

### 3. 安裝 Python 依賴套件

本專案的所有 Python 依賴項都記錄在 `requirements.txt` 中。

1.  **建立虛擬環境 (建議)**:
    ```bash
    python -m venv venv
    source venv/bin/activate  # macOS/Linux
    .\venv\Scripts\activate  # Windows
    ```

2.  **安裝依賴**:
    ```bash
    pip install -r requirements.txt
    ```

如果您沒有 `requirements.txt` 檔案，可以手動安裝以下套件：
```bash
pip install ttkbootstrap ttkbootstrap-dnd2 pytesseract Pillow python-pptx extract-msg requests
```

## 使用方法

1.  **啟動應用程式**:
    ```bash
    python main.py
    ```

2.  **步驟 1：輸入資料**
    -   **拖曳檔案**：將圖片 (`.png`, `.jpg`)、文字檔 (`.txt`) 或 Email 檔 (`.msg`) 拖曳至程式視窗內。
    -   **上傳/貼上**：點擊「1. 上傳/貼上」按鈕。如果剪貼簿中有圖片，程式會自動辨識；否則，會彈出檔案選擇對話框。
    -   所有辨識出的文字會顯示在左側的「步驟 A: 辨識結果」文字框中。

3.  **步驟 2：生成 AI 報告**
    -   在「專案名稱」欄位中輸入專案名稱（選填）。
    -   根據您的需求，點擊「分析為『單一問題』報告」或「分析為『多個問題』報告」。
    -   程式會自動呼叫本地的 Ollama 模型進行分析。請耐心等待，UI 狀態列會顯示進度。
    -   分析完成後，結果會顯示在右側的「步驟 C: Ollama 生成結果」文字框中。

4.  **步驟 3：生成 PowerPoint 投影片**
    -   確認右側的報告內容無誤後，點擊「步驟 D: 新增至彙總簡報」按鈕。
    -   程式會將報告內容新增至專案資料夾下的 `Weekly Report_JimChuang.pptx` 檔案中。
    -   **注意**：生成前，請確保該 PowerPoint 檔案處於關閉狀態，否則會因權限問題導致儲存失敗。

## 專案結構

本專案採用模組化設計，以提高可讀性與可維護性。

```
.
├── main.py                   # 應用程式主入口
├── app_ui.py                 # UI 介面層 (View)
├── app_controller.py         # 控制器層 (Controller)
├── services.py               # 核心服務層 (Model/Logic)
├── config.py                 # 全域設定檔
├── prompt_single_issue.txt   # 單一問題分析的 Prompt 模板
├── prompt_multi_issue.txt    # 多個問題分析的 Prompt 模板
├── requirements.txt          # Python 依賴套件列表
└── README.md                 # 本說明文件
```

## 客製化設定

您可以透過修改 `config.py` 檔案來客製化應用程式的行為：

-   `MASTER_PPTX_FILENAME`: 修改預設生成的 PowerPoint 檔案名稱。
-   `OLLAMA_API_URL`: 如果您的 Ollama 運行在不同的主機或埠號，請在此修改。
-   `OLLAMA_MODEL`: 更換您想使用的 Ollama 模型名稱（例如 `mistral`, `gemma` 等）。
-   `STAR_KEYWORDS`: 如果您的報告格式需要識別不同的關鍵字，可以在此處修改。

您也可以直接編輯 `prompt_*.txt` 檔案，來調整 AI 生成報告的風格、語氣和格式。

## 疑難排解

-   **Tesseract 未找到錯誤**: 請確認 Tesseract 已正確安裝，並且其路徑已加入系統環境變數 `Path` 中。
-   **無法連接至 Ollama**: 請確認 Ollama 應用程式正在本機端運行，並且 `config.py` 中的 `OLLAMA_API_URL` 設定正確。
-   **PPT 權限錯誤**: 在生成 PowerPoint 之前，請務必關閉正在編輯的目標檔案。

---