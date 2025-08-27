# 文件自動更新工具 (Transcript & Management Report Generator)

## 版本資訊
**當前版本：v2.1.0**
- 🔧 **簡化版本**：回復到傳統模式，移除複雜的工作流程管理
- 📊 **穩定可靠**：專注於核心功能，提供穩定的文件生成
- 🎛️ **雙模式支援**：GUI 和傳統命令列模式
- 修復 exe 環境中表格更新路徑問題
- 新增千分位分隔符號格式化
- 增強數字單位處理 (ppt/ppts vs percentage point/percentage points)

## 專案簡介
這是一個自動化工具，用於生成投資者會議記錄（Transcript）和管理報告（Management Report）。v2.1.0 版本回復到簡潔穩定的傳統模式，專注於核心的文件生成功能。

## 🚀 功能特色
- 🔄 **自動化文件生成**：從 Excel 自動生成 Word 文件
- 🌐 **多語言支援**：支援中英文投資者會議記錄
- 📊 **智能數據處理**：自動格式化數字、計算變化率
- 🎯 **靈活範本系統**：支援不同季度的範本配置
- 💡 **用戶友好介面**：圖形化操作介面，操作簡單
- 📈 **豐富的數據分析**：財務結果、營運數據、現金流等多維度分析

## 使用方式

### 🖥️ GUI 模式（推薦）
```bash
python main.py
```
- 圖形界面操作簡單直觀
- 支援多種文件類型選擇

### 🔧 命令列模式
```bash
# 基本使用
python main.py --traditional "path/to/excel.xlsx"

# 指定特定報告類型
python main.py --traditional "input/data.xlsx" --reports transcript_en,management_report
```

### 📖 顯示說明
```bash
python main.py --help
```

## 功能特色
- 🔄 **自動化文件生成**：從 Excel 自動生成 Word 文件
- 🌐 **多語言支援**：支援中英文投資者會議記錄
- 📊 **智能數據處理**：自動格式化數字、計算變化率
- 🎯 **靈活範本系統**：支援不同季度的範本配置
- 💡 **用戶友好介面**：圖形化操作介面，操作簡單
- 📈 **豐富的數據分析**：財務結果、營運數據、現金流等多維度分析

## 系統需求
- Windows 作業系統
- 已安裝 Microsoft Word（用於 Word 文件處理）
- Excel 檔案格式支援

## 檔案結構
```
Document_Update/
├── main.py                    # 主程式入口
├── configs/
│   ├── config.json           # 配置文件
│   └── constant.py           # 常數定義
├── core/
│   ├── data_parser.py        # 數據解析器
│   ├── excel_table_to_word.py # Excel轉Word核心邏輯
│   ├── report_generator.py   # 報告生成器
│   └── history_utils.py      # 歷史記錄工具
├── gui/
│   └── interface.py          # 圖形界面
├── templates/                # Word模板
├── input/                    # 輸入Excel文件
└── output/                   # 輸出Word文件
```

## 安裝與設定

### 必要套件
```bash
pip install pandas docxtpl jinja2 openpyxl python-docx
```

### 配置說明
1. 編輯 `configs/config.json` 設定季度和年份
2. 將 Excel 文件放入 `input/` 目錄
3. 確保 `templates/` 目錄包含所需的 Word 模板

## 注意事項
- 請確保 Excel 文件格式正確，包含所需的工作表
- Word 模板必須使用 docxtpl 格式的變數標記
- 輸出文件會自動保存到 `output/` 目錄

## 故障排除
如遇到問題，請檢查：
1. Excel 文件是否包含所需的工作表和數據
2. Word 模板是否格式正確
3. 是否已安裝所有必要的 Python 套件
4. 檔案權限是否正確設定

## 更新記錄
- v2.1.0: 簡化版本，移除複雜功能，專注核心文件生成
- v2.0.x: 複雜工作流程版本（已移除）
- v1.x: 基礎版本
