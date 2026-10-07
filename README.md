# 研習證書製作工具

將 Excel 參加者名單的資料填入 Word 證書範本，為每位合格學員建立一份 `.docx` 證書。原始 Excel 和 Word 範本只會被讀取，不會被修改。

此儲存庫包含程式碼；Excel 名單與 Word 範本需自行準備，欄位與文字標記格式如下。

## 第一次啟動

在本資料夾開啟 PowerShell，執行：

```powershell
uv sync
uv run certificate-maker
```

預設會讀取 `參加者名單.xlsx` 與 `研習證書範本.docx`，並將證書放在 `輸出證書` 資料夾。完成後請開啟 `輸出證書\產生報告.txt` 檢視結果和待處理資料。

## Excel 欄位

第一列必須包含以下欄名：`學員編號`、`姓名`、`Email`、`課程名稱`、`完成日期`。所有欄位皆為必填；工具會跳過缺漏資料及學員編號重複的資料列，並在報告中列出原因。

完成日期若是 Excel 日期欄位，會以 `YYYY年MM月DD日` 填入證書。

## Word 範本欄位

在範本中放入以下文字標記：

- `{{姓名}}`
- `{{課程名稱}}`
- `{{完成日期}}`

工具只取代這些標記，並保留範本的版面、圖片、頁首頁尾和其他內容。

## 更換檔案與輸出位置

使用完整或相對路徑指定自己的檔案：

```powershell
uv run certificate-maker --roster ".\我的名單.xlsx" --template ".\我的證書範本.docx" --output ".\我的輸出證書"
```

輸出檔名為 `學員編號_姓名_研習證書.docx`。若同名輸出檔已存在，會被新的產生結果覆蓋；建議每次使用不同輸出資料夾保留歷次結果。

## 將既有證書轉為 PDF

此功能透過 Python 的 `pywin32` 直接呼叫 Microsoft Word 匯出 PDF，需在 Windows 上安裝並完成 Word 的首次啟動設定。執行 `uv sync` 會安裝所需 Python 套件。它只讀取既有 `.docx` 證書，不會重新產生或修改 Word 檔案，並會略過 Word 的 `~$` 暫存檔。

```powershell
uv run certificate-to-pdf --input ".\輸出證書" --output ".\輸出 PDF 證書"
```

完成後請查看 `輸出 PDF 證書\轉檔報告.txt`，其中會列出每份 Word 證書對應的 PDF 檔案或失敗原因。

若轉檔超時，工具會在 120 秒後停止等待，保留已完成的結果，並在報告標示卡住的階段。可用 `--timeout 300` 延長整批轉檔的等待時間。輸出資料夾的 `轉檔紀錄.txt` 會記錄開啟、匯出及關閉文件的進度。逾時時工具會嘗試結束本次新建的 Word 程序。

## 一次完成 Word 與 PDF

此批次使用 Microsoft Word 匯出 PDF，與 `certificate-to-pdf` 共用相同的 Python COM 轉檔流程。`--timeout` 是每份文件的等待上限（預設 120 秒）。

每次執行都會建立新的批次資料夾，包含 `Word`、`PDF` 與 `處理結果.txt`，不會與舊批次混在一起或覆寫既有成果。Word 製作失敗時不會進行 PDF 轉換；PDF 失敗時會保留已產生的 Word，並在報告列出原因。

```powershell
uv run certificate-batch --roster ".\參加者名單.xlsx" --template ".\研習證書範本.docx"
```

若要指定批次上層資料夾：

```powershell
uv run certificate-batch --roster ".\我的名單.xlsx" --template ".\我的證書範本.docx" --output-root ".\歷次證書"
```

## Gmail SMTP 測試寄信（不會寄送證書）

此項功能只會在您明確執行指令時寄出一封純文字測試信，**不會讀取名單、不會附加證書，也不會在產生證書後自動寄信**。

### 首次設定

1. 將 `.env.example` 複製為 `.env`，填入自己的 Gmail 位址與 Google「應用程式密碼」：

   ```ini
   GMAIL_SMTP_USER=your-account@gmail.com
   GMAIL_SMTP_APP_PASSWORD=在此填入16碼應用程式密碼
   ```

   `.env.example` 保留了不含任何真實密碼的範例，可在遺失設定時參考。`.env` 已加入 `.gitignore`；若日後初始化 Git，執行 `git status` 時不應看見 `.env`。

2. Gmail 必須先啟用兩步驟驗證，接著在 Google 帳戶的「應用程式密碼」建立一組 Mail 用密碼。請將這組密碼填入 `.env`，不是 Gmail 的一般登入密碼。請勿將密碼貼到終端機、對話或任何版本控制檔案。

### 寄送給自己測試

填妥 `.env` 後，在本資料夾執行：

```powershell
uv run certificate-email-test --to "your-account@gmail.com"
```

可選擇自訂主旨：

```powershell
uv run certificate-email-test --to "your-account@gmail.com" --subject "證書工具 SMTP 測試"
```

程式不會在畫面、錯誤訊息或任何檔案中輸出應用程式密碼；成功時只會顯示收件者，失敗時只會顯示不含敏感值的診斷訊息。

## 建立、轉檔並寄送 PDF 證書

`certificate-email-batch` 會為每位名單資料完整的參加者建立 Word 證書、轉成 PDF，再依 Excel 的 `Email` 欄位寄出 PDF 附件。信件內文含姓名、課程名稱與簡短說明。資料缺漏、Word 製作失敗或 PDF 轉檔失敗時，該名參加者不會收到郵件。

PDF 轉檔使用 Microsoft Word，`--timeout` 是每份文件的等待上限（預設 120 秒）。兩種批次的 `PDF/轉檔紀錄.txt` 都會保留每份文件的轉檔進度。

正式執行會以一個指令完成 Word 製作、PDF 轉檔與 Gmail 寄送，不需要額外寄送指令或中途確認：

```powershell
uv run certificate-email-batch --roster ".\正式名單.xlsx" --template ".\研習證書範本.docx" --send
```

測試執行是可獨立使用的 dry run：會完成相同的 Word/PDF 處理，並在 `處理與寄送結果.txt` 列出每位收件者及其對應 PDF 附件，但完全不連線或寄送 Gmail，也不會標記為已寄送：

```powershell
uv run certificate-email-batch --roster ".\我的測試名單.xlsx" --template ".\研習證書範本.docx"
```

每個批次會保留 Word、PDF、`處理與寄送結果.txt` 與 `寄送狀態.json`。後者只記錄參加者資料、處理狀態、寄送狀態與不含密碼的失敗原因。重新執行同一批時，`email_status=sent` 的項目會略過，因此不會重寄已成功的郵件；失敗或尚未寄送的項目會重試，並沿用已成功建立的 PDF 附件：

```powershell
uv run certificate-email-batch --batch-dir ".\證書寄送批次\寄送批次_YYYYMMDD_HHMMSS" --send
```

若只想更新檔案而不寄信，移除 `--send` 即可再次進行 dry run。
