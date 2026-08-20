# IR 講稿工具 操作手冊

給整個 IR 團隊看的操作說明，不需要懂程式。技術架構說明請看 [README.md](README.md)。

## 這個工具是做什麼的

從 Excel 財務數據自動產生：
- 法說會 Transcript（英文、中文）
- Management Report
- 法說會場次統計查詢
- 待公告論壇提醒 + 資訊發布申請單

## 開始之前

目前這個工具需要在**有裝 Python 的電腦上**啟動後端程式才能用：

1. 打開終端機，進到專案資料夾下的 `backend/`
2. 執行 `python app.py`
3. 打開瀏覽器，連到 `http://127.0.0.1:5001`

> 之後如果部署到公司內部伺服器，這一步會改成直接連伺服器網址，不需要在自己電腦上啟動任何東西。目前還沒到那一步，所以這一步不能省略。

## 每季例行流程（自辦法說會）

1. **準備好這一季的 Excel 資料檔**（`data input.xlsx`），放到 `scripts/document_update/input/` 資料夾，或用網頁上「開啟 Excel」按鈕直接編輯
2. 打開網頁，切到 **「自辦法說會」** 頁籤
3. 填寫**本季度（Q1–Q4）、年度、法說會日期**，按 **「確認季度設定」**——這一步做完，「產生文件」按鈕才會解鎖
4. 確認 Excel 資料都填好，尤其是 `financial_results`、`revenue_streams`、`tech`、`wafer_size` 這幾個會被拿去自動組成英文/中文措辭句子的分頁
5. 按 **「產生文件」**，系統會產出三份文件到 `scripts/document_update/output/`：
   - 英文 Investor Conference Transcript
   - 中文 Investor Conference Transcript
   - Management Report
6. **檢查產出文件裡有沒有紅字。** 紅字代表這個分頁的內容跟上一季的快照一模一樣，系統判斷你可能忘了更新這部分數據——務必人工複查，確認是真的沒變動還是漏填了

## 講稿文字維護

**「講稿文字輸入」** 頁籤可以直接編輯以下三個分頁的中英文內容，存檔會直接寫回 Excel，不用另外開 Excel 手動改：
- Opening remarks
- Future outlook
- Chairman remarks

## 法說會活動管理

切到 **「法說會活動管理」** 頁籤：

- **查詢場次**：輸入起訖日期，查詢自辦／受邀-1on1／受邀-論壇的場次統計，以及國內/外資法人比例
- **檢查待公告場次**：列出未來需要送資訊發布申請單的論壇，附距離天數提醒
- **產生資訊發佈申請單**：批次填寫申請單資料（可逐場切換確認主辦方、時間、地點）

## 常見狀況

| 狀況 | 原因 |
|---|---|
| 「產生文件」按鈕按不下去 | 還沒按「確認季度設定」，或找不到 Excel 檔案 |
| 產出的 Word 裡有紅字段落 | 對應的 Excel 分頁跟上一季內容相同，系統判斷可能沒更新，需人工確認 |
| 查詢場次找不到資料 | 檢查 `ir_meetings_base_dir` 設定的資料夾路徑是否存在、年份資料夾命名是否正確 |

## 資料存放位置

| 內容 | 位置 | 是否進 git |
|---|---|---|
| 本季 Excel 輸入資料 | `scripts/document_update/input/data input.xlsx` | 否，每次手動放 |
| 產出的 Word 文件 | `scripts/document_update/output/` | 否 |
| 法說會歷史紀錄（統計、公告用） | 公司 OneDrive「法說會與座談會」資料夾 | 不適用（存在公司雲端） |
| 程式碼、範本、設定檔 | GitHub：https://github.com/Cheryl-1110802/generate-ir-transcript | 是 |
