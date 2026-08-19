# -*- coding: utf-8 -*-
# 文件自動更新工具
import sys
import os
import argparse

# 設定正確的工作目錄為腳本所在目錄
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# 版本資訊
VERSION = "2.1.0"

def get_resource_path(relative_path):
    """獲取資源檔案的絕對路徑，適用於開發環境和 PyInstaller 打包環境"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包環境
        base_path = sys._MEIPASS
    else:
        # 開發環境
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def get_output_path():
    """獲取輸出目錄的路徑，確保可寫入"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 環境中，找到專案根目錄下的 output 目錄
        exe_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        # 嘗試找到專案根目錄（可能在上一層或同一層）
        possible_paths = [
            os.path.join(exe_dir, "..", "output"),  # exe 在 dist 目錄中，output 在上一層
            os.path.join(exe_dir, "output"),        # exe 和 output 在同一層
        ]
        
        for path in possible_paths:
            abs_path = os.path.abspath(path)
            if os.path.exists(abs_path):
                return abs_path
        
        # 如果都找不到，就在 exe 同級目錄創建
        output_dir = os.path.join(exe_dir, "output")
    else:
        # 開發環境
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, "output")
    
    # 確保輸出目錄存在
    os.makedirs(output_dir, exist_ok=True)
    return output_dir

def print_usage():
    """顯示使用說明"""
    print(f"""
文件自動更新工具 v{VERSION}

使用方式:
1. 自動處理模式 (使用預設 Excel 檔案):
   python main.py
   
2. 指定 Excel 檔案:
   python main.py --excel <excel_path> [--reports <report_types>]
   
   範例:
   python main.py --excel "input/data input.xlsx"
   python main.py --excel "input/data input.xlsx" --reports transcript_en,management_report

參數說明:
  --excel          指定 Excel 檔案路徑 (預設: input/data input.xlsx)
  --reports        要生成的報告類型，用逗號分隔 (預設: transcript_en,transcript_zh,management_report)
  --help, -h       顯示此說明
""")

def run_document_update(excel_path=None, selected_reports=None):
    """執行文件更新"""
    try:
        # 預設值
        if not excel_path:
            excel_path = os.path.join(script_dir, "input", "data input.xlsx")
        
        if not selected_reports:
            selected_reports = ['transcript_en', 'transcript_zh', 'management_report']
        
        # 檢查 Excel 檔案是否存在
        if not os.path.exists(excel_path):
            print(f"錯誤: 找不到 Excel 檔案: {excel_path}")
            return False
        
        print(f"=== 開始處理文件更新 ===")
        print(f"Excel 檔案: {excel_path}")
        print(f"報告類型: {', '.join(selected_reports)}")
        
        from core.report_generator import main
        main(excel_path, selected_reports)
        print("文件更新處理完成！")
        return True
        
    except Exception as e:
        print(f"執行時發生錯誤: {e}")
        import traceback
        traceback.print_exc()
        return False

if __name__ == "__main__":
    # 解析命令列參數
    parser = argparse.ArgumentParser(description=f"文件自動更新工具 v{VERSION}", add_help=False)
    parser.add_argument('--excel', help='Excel 檔案路徑')
    parser.add_argument('--reports', help='要生成的報告類型 (逗號分隔)', default='transcript_en,transcript_zh,management_report')
    parser.add_argument('--help', '-h', action='store_true', help='顯示說明')
    
    args = parser.parse_args()
    
    print(f"文件自動更新工具 v{VERSION}")
    
    if args.help:
        print_usage()
        sys.exit(0)
    
    # 處理報告類型參數
    selected_reports = [r.strip() for r in args.reports.split(',')]
    
    # 執行文件更新
    success = run_document_update(args.excel, selected_reports)
    
    if not success:
        sys.exit(1)