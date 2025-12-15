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

好的，這是一個非常重要的步驟，一份清晰的 `README.md` 文件是確保其他人（以及未來的您）能夠順利使用這個工具的關鍵。

我將為您撰寫一份詳細的 `README.md` 文件，內容涵蓋了專案介紹、安裝步驟、設定說明、詳細的使用流程（包括您最新的「附加模式」），以及常見問題排解。

您可以直接將以下內容複製並儲存為專案根目錄下的 `README.md` 檔案。

---

# 報告整理小幫手 v18.0

這是一個桌面應用程式，旨在幫助使用者快速將螢幕截圖、圖片、文字檔或 Outlook 郵件（`.msg`）的內容，整理成符合 STAR 原則（情境、任務、行動、結果）的結構化報告，並能一鍵將報告新增至 PowerPoint 簡報中。

本工具利用本地運行的 Ollama 大型語言模型（LLM）和視覺語言模型（VLM），確保所有資料都在您的本機端處理，兼顧效率與隱私。

## 主要功能

- **智慧圖片辨識 (VLM-OCR)**：利用視覺語言模型（如 LLaVA）進行高精準度的文字辨識，優於傳統 OCR。
- **彈性文字輸入**：支援多種輸入方式：
    - 檔案上傳（`.txt`, `.msg`）
    - 圖片分析（`.png`, `.jpg`, `.jpeg`）
    - 剪貼簿貼上圖片
    - 拖曳檔案至視窗
- **附加模式**：支援連續處理多張圖片，並將辨識結果附加到現有內容中，適合處理長篇內容。
- **智慧報告生成 (LLM)**：呼叫大型語言模型，根據指定的 Prompt（單一問題或多個問題）將原始文字整理成 STAR 報告。
- **一鍵生成簡報**：將生成好的報告快速新增為 PowerPoint 投影片。
- **高度可配置**：透過 `settings.json` 檔案，使用者可以輕鬆自訂模型、檔案名稱和效能參數。
- **Ollama 自動管理**：應用程式啟動時會自動檢查並嘗試在背景啟動 Ollama 服務。

## 環境準備與安裝

在執行本程式前，請確保您的電腦已完成以下設定。

### 1. 安裝 Ollama

本工具依賴 Ollama 在本地端運行 AI 模型。

- 前往 [Ollama 官方網站](https://ollama.com/) 下載並安裝適合您作業系統的版本。
- **（重要）** 為了獲得最佳的 VLM 支援，請確保您的 Ollama 是最新版本。可以在終端機執行以下指令進行更新：
  ```bash
  ollama pull ollama
  ```

### 2. 下載必要的 AI 模型

安裝完 Ollama 後，請在終端機執行以下指令，下載本工具預設需要的模型：

- **下載視覺語言模型 (VLM)** - 用於圖片辨識：
  ```bash
  ollama pull llava:latest
  ```
- **下載大型語言模型 (LLM)** - 用於 STAR 分析：
  ```bash
  ollama pull llama3:8b
  ```
> **提示**：您可以在 `settings.json` 檔案中將模型名稱更換為您偏好的其他模型。

### 3. 安裝 Python 函式庫

本專案使用 Python 進行開發。請在終端機中，進入專案目錄，並執行以下指令安裝所有必要的函式庫：

```bash
pip install -r requirements.txt
```

如果專案中沒有 `requirements.txt` 檔案，請手動安裝以下函式庫：
```bash
pip install ttkbootstrap tkinterdnd2 Pillow python-pptx extract-msg requests
```

### 4. (可選) 安裝 Tesseract-OCR

本工具的 VLM-OCR 功能已能取代傳統 OCR。但為了保留舊有的「載入文字/郵件」流程中對圖片的相容性，建議安裝 Tesseract。

- 前往 [Tesseract 官方 GitHub](https://github.com/tesseract-ocr/tessdoc) 找到安裝說明。
- **Windows 使用者**：安裝時請務必勾選新增繁體中文、簡體中文和英文的語言包。
- **重要**：安裝後請將 Tesseract 的安裝路徑加入系統的環境變數 `PATH` 中。

## 設定 (`settings.json`)

首次執行應用程式時，會在專案目錄下自動生成一個 `settings.json` 檔案。您可以修改此檔案來自訂程式行為。

```json
{
    "OLLAMA_API_URL": "http://localhost:11434/api/generate",
    "OLLAMA_MODEL": "llama3:8b",
    "OLLAMA_VLM_MODEL": "llava:latest",
    "MASTER_PPTX_FILENAME": "Weekly Report_JimChuang.pptx",
    "JPEG_QUALITY": 85
}
```

- `OLLAMA_MODEL`: 用於**STAR 分析**的純文字語言模型。
- `OLLAMA_VLM_MODEL`: 用於**圖片辨識 (OCR)** 的視覺語言模型。
- `MASTER_PPTX_FILENAME`: 最終生成的 PowerPoint 檔案名稱。
- `JPEG_QUALITY`: VLM 處理圖片前的壓縮品質（0-100）。**這是重要的效能調整參數**。
  - **建議值 `70-90`**：在速度和辨識準確率之間取得良好平衡。
  - **降低此值可加快處理速度**，但過低（如低於 60）可能會影響辨識準確率。

## 使用方法

### 啟動程式

在終端機中，進入專案目錄，執行：
```bash
python main.py
```
程式啟動後，請稍待狀態列顯示「Ollama 已就緒！請選擇操作。」，代表模型已預熱完畢。

### 主要工作流程 (推薦)

此流程利用 VLM 進行高效率、高準確率的辨識。

**情境一：處理單張圖片**

1.  在「專案名稱」欄位輸入您的專案名稱（選填）。
2.  確保右下角的「**附加到現有內容**」複選框**未被勾選**。
3.  點擊「**智慧分析圖片**」選擇圖片檔案，或複製圖片後點擊「**貼上圖片分析**」。
4.  等待 VLM 辨識完成，辨識出的純文字會顯示在左側的「步驟 A」文字區。
5.  檢查或手動編輯左側的文字內容。
6.  點擊「**步驟 B**」的按鈕（「分析為單一問題」或「分析為多個問題」）。
7.  LLM 生成的 STAR 報告會顯示在右側的「步驟 C」文字區。
8.  點擊「**步驟 D**」將報告新增至 PowerPoint 簡報。

**情境二：處理多張連續的圖片 (例如長篇對話截圖)**

1.  **處理第一張圖**：按照「情境一」的步驟 1-4，先處理第一張圖片。
2.  **啟用附加模式**：**勾選**右下角的「**附加到現有內容**」複選框。
3.  **處理後續圖片**：重複使用「智慧分析圖片」或「貼上圖片分析」來處理第二張、第三張...圖片。每一次，新的辨識結果都會被加上分隔線並附加到左側文字區的末尾。
4.  **最終分析**：當所有圖片都處理完畢後，檢查並編輯左側的完整內容，然後點擊「步驟 B」進行 STAR 分析。
5.  點擊「步驟 D」生成簡報。

## 常見問題 (FAQ)

1.  **問題：VLM 處理圖片非常慢，甚至超時 (Timeout)。**
    - **原因**：本地端電腦效能不足，尤其是 CPU 執行 VLM 時。
    - **解決方案**：
        1.  打開 `settings.json` 檔案。
        2.  將 `JPEG_QUALITY` 的值從 `85` **調低**，例如 `75` 或 `70`。這會顯著加快處理速度。
        3.  如果速度仍然很慢，可以嘗試更換一個較小的 VLM 模型，例如將 `OLLAMA_VLM_MODEL` 改為 `"llava:7b"`，並在終端機執行 `ollama pull llava:7b`。

2.  **問題：VLM 辨識結果只顯示 `! Picture 1` 或內容不正確。**
    - **原因**：Ollama 服務沒有成功處理圖片資料，這通常是 Ollama 版本過舊導致的。
    - **解決方案**：在終端機執行 `ollama pull ollama` 將 Ollama 更新到最新版本，然後重啟應用程式。

3.  **問題：點擊「新增至彙總簡報」時，出現「權限錯誤 (PermissionError)」。**
    - **原因**：您要寫入的 PowerPoint 檔案（例如 `Weekly Report_JimChuang.pptx`）正被其他程式（如 Microsoft PowerPoint）開啟。
    - **解決方案**：請先將對應的 PowerPoint 檔案完全關閉，然後再點擊按鈕。