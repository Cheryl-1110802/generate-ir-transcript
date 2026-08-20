# generate-ir-transcript

從 Excel 財務數據自動產生法說會 Transcript(中/英文)、Management Report,以及法說會場次統計、資訊揭露申請單的內部工具。

給日常操作流程的請看 [SOP.md](SOP.md)。這份文件是給要維護/修改程式的人看的架構說明。

## 架構

```
backend/                     Flask API,網頁前端跟核心邏輯之間的橋接層
  app.py                        所有 API 路由
  config.json                   伺服器設定、腳本清單、法說會資料夾路徑
  serve.py                      用 waitress 跑的正式進入點（取代 app.py 內建的開發用伺服器）

web/
  IR小工具.html                單頁前端，呼叫 backend 的 API

scripts/document_update/     核心報告產生邏輯
  core/
    report_generator.py          主流程：讀 Excel → 解析各 sheet → 套用 Word 範本 → 產出
    data_parser.py                逐一 sheet 的解析邏輯（中英文措辭判斷、千分位、數字轉換）
    excel_table_to_word.py        把 Excel 表格範圍逐格貼進 Word 特定位置
    snapshot_utils.py             跟上一季快照比對，標記沒更新的 sheet（紅字）
    history_utils.py              tape-out 累計數的歷史追蹤
  configs/
    data_config.json              sheet 名稱對應、Word 範本路徑
    quarter_config.json           目前季度／年度／法說會日期
    constant.py                   董事長／總經理／財務長姓名等常數
  templates/                     Word 範本（.docx），裡面是 Jinja2 標籤
  tools/                         docx → html 轉換的實驗性工具（為未來可能的 n8n 遷移做準備）
```

## 本機執行方式

```bash
cd backend
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

打開 `http://127.0.0.1:5001`。

## 重要：不會被版控的檔案

`scripts/document_update/input/*.xlsx`、`output/*.docx`、`output/*history*.json` 都被排除在 git 之外（見 `.gitignore`），因為裡面是真實、可能還沒公開揭露的財報數字。**每次拉新環境都要自己手動放資料，不要把真實資料 commit 進去。**

## 已知限制 / 未來方向

- `/open-excel` 用 `os.startfile`，只有在跑程式的機器上有桌面環境時才有意義，部署到伺服器上這個功能不會動
- `tools/` 裡有把 Word 範本轉成 HTML + Jinja2 渲染的雛型，是為了未來可能把產生邏輯搬去 n8n 之類的環境做準備，目前尚未接上主流程

## 相關連結

- Repo: https://github.com/Cheryl-1110802/generate-ir-transcript
