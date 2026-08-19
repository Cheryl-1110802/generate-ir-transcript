# 文件自動更新工具 (Transcript & Management Report Generator)

## 版本資訊
**當前版本：v3.0.0 (Web整合版)**
- 🌐 **Web整合**：已整合至統一的 IR小工具 Web 介面
- 🔧 **移除GUI依賴**：main.py 改為純命令行工具，支援 API 調用
- 📊 **穩定可靠**：專注於核心功能，提供穩定的文件生成
- 🎛️ **多模式支援**：Web 介面（推薦）和命令列模式
- ✨ **簡化輸出**：優化執行結果顯示，專注關鍵資訊

**注意：推薦使用統一的 Web 介面進行操作，獲得最佳使用體驗。**

## 專案簡介
這是一個自動化工具，用於生成投資者會議記錄（Transcript）和管理報告（Management Report）。v3.0.0 版本已完全整合至 Web 前端，同時保持命令行功能以支援程式化調用。

## 🚀 功能特色
- 🔄 **自動化文件生成**：從 Excel 自動生成 Word 文件
- 🌐 **多語言支援**：支援中英文投資者會議記錄
- 📊 **智能數據處理**：自動格式化數字、計算變化率
- 🎯 **靈活範本系統**：支援不同季度的範本配置
- 💡 **用戶友好介面**：圖形化操作介面，操作簡單
- 📈 **豐富的數據分析**：財務結果、營運數據、現金流等多維度分析

## 使用方式

### 🌐 Web 介面模式（推薦）
**透過統一的 IR小工具 Web 介面使用**
- 🔗 存取位置：`http://127.0.0.1:5001`（需先啟動 backend/app.py）
- 📋 操作簡單：選擇季度、年度、舉辦日期後點擊「完成設定並開始更新」
- 🎯 功能完整：設定更新、文件生成、結果顯示一站式完成
- 📱 響應式設計：支援各種裝置使用

### �️ 命令列模式
```bash
# 使用預設 Excel 檔案
python main.py

# 指定 Excel 檔案
python main.py --excel "input/data input.xlsx"

# 指定特定報告類型
python main.py --excel "input/data input.xlsx" --reports transcript_en,management_report

# 顯示說明
python main.py --help
```

### 🔧 API 整合模式
**供其他系統呼叫使用**
- 後端 API：`POST /run/script1`
- 配置更新：`POST /update-config`
- 適用於自動化工作流程整合

## 功能特色
- 🔄 **自動化文件生成**：從 Excel 自動生成 Word 文件
- 🌐 **多語言支援**：支援中英文投資者會議記錄
- 📊 **智能數據處理**：自動格式化數字、計算變化率
- 🎯 **靈活範本系統**：支援不同季度的範本配置
- 💡 **用戶友好介面**：圖形化操作介面，操作簡單
- 📈 **豐富的數據分析**：財務結果、營運數據、現金流等多維度分析

## 系統需求
- Windows 作業系統
- Python 3.7+
- 已安裝 Microsoft Word（用於 Word 文件處理）
- Excel 檔案格式支援
- 現代瀏覽器支援（用於 Web 介面）

## 檔案結構
```
document_update/
├── main.py                    # 命令行主程式（移除GUI依賴）
├── configs/
│   ├── config.json           # 配置文件
│   └── constant.py           # 常數定義
├── core/
│   ├── data_parser.py        # 數據解析器
│   ├── excel_table_to_word.py # Excel轉Word核心邏輯
│   ├── report_generator.py   # 報告生成器
│   └── history_utils.py      # 歷史記錄工具
├── templates/                # Word模板
├── input/                    # 輸入Excel文件
└── output/                   # 輸出Word文件
```

**Web 整合檔案** (位於專案根目錄):
```
panel_for_IR/
├── web/
│   └── IR小工具_fixed.html    # 統一Web前端介面
├── backend/
│   └── app.py                # Flask API服務器
└── scripts/
    └── document_update/      # 本模組
```

## 安裝與設定

### 必要套件
```bash
pip install pandas docxtpl jinja2 openpyxl python-docx flask
```

### Web 模式設定
1. 啟動後端服務：`python backend/app.py`
2. 開啟瀏覽器：`http://127.0.0.1:5001`
3. 在「自辦線上法說會」區塊進行操作

### 命令列模式設定
1. 編輯 `configs/config.json` 設定季度和年份（或透過參數指定）
2. 將 Excel 文件放入 `input/` 目錄
3. 確保 `templates/` 目錄包含所需的 Word 模板

## 故障排除
如遇到問題，請檢查：

### Web 介面問題
1. 後端服務是否正確啟動（backend/app.py）
2. 瀏覽器是否能正常存取 `http://127.0.0.1:5001`
3. 防火牆是否阻擋本地端口 5001

### 文件生成問題
1. Excel 文件是否包含所需的工作表和數據
2. Word 模板是否格式正確
3. 是否已安裝所有必要的 Python 套件
4. 檔案權限是否正確設定

### 配置問題
1. `configs/config.json` 格式是否正確
2. 季度、年度、日期設定是否合理
3. 輸入和輸出目錄是否存在且可存取

## 更新記錄
- **v3.0.0** (2025年8月): Web整合版本
  - ✅ 完整整合至統一 Web 介面
  - ✅ 移除 GUI 依賴，改為純命令行工具
  - ✅ 新增 RESTful API 支援
  - ✅ 優化執行結果顯示格式
  - ✅ 支援 Web 前端的配置更新功能

- **v2.1.0**: 簡化版本
  - 移除複雜功能，專注核心文件生成
  - 修復 exe 環境問題
  - 新增數字格式化功能

- **v2.0.x**: 複雜工作流程版本（已移除）
- **v1.x**: 基礎版本

---
**目前狀態**: 🌐 已整合至統一工作台  
**推薦使用方式**: Web 介面操作  
**維護狀態**: 生產就緒版本
