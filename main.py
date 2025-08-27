from gui.interface import launch_gui
import sys
import os
import argparse

# 版本資訊 - 每次修改後更新這個版本號
VERSION = "2.1.0"  # 移除工作流程管理器，回復傳統模式

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
1. GUI 模式 (預設，使用智慧工作流程):
   python main.py
   
2. 工作流程模式 (命令列):
   python main.py --workflow <excel_path> [--previous-word <word_path>] [--reports <report_types>]
   
   範例:
   python main.py --workflow "input/data.xlsx"
   python main.py --workflow "input/data.xlsx" --previous-word "output/previous.docx" --reports transcript_en,management_report

3. 傳統模式:
   python main.py --traditional <excel_path>

參數說明:
  --workflow       使用新的六步驟工作流程
  --traditional    使用原有的報告生成方式  
  --previous-word  上一季講稿路徑 (可選)
  --reports        要生成的報告類型，用逗號分隔 (預設: transcript_en,transcript_zh,management_report)
  --help, -h       顯示此說明
""")

if __name__ == "__main__":
    # 解析命令列參數
    parser = argparse.ArgumentParser(description=f"文件自動更新工具 v{VERSION}", add_help=False)
    parser.add_argument('--workflow', help='使用工作流程模式的 Excel 檔案路徑')
    parser.add_argument('--traditional', help='使用傳統模式的 Excel 檔案路徑')
    parser.add_argument('--previous-word', help='上一季講稿路徑')
    parser.add_argument('--reports', help='要生成的報告類型 (逗號分隔)', default='transcript_en,transcript_zh,management_report')
    parser.add_argument('--help', '-h', action='store_true', help='顯示說明')
    
    args = parser.parse_args()
    
    print(f"文件自動更新工具 v{VERSION}")
    
    if args.help:
        print_usage()
        sys.exit(0)
    
    try:
        if args.workflow:
            # 工作流程模式已移除，改用傳統模式
            print("=== 工作流程模式已移除，使用傳統模式 ===")
            from core.report_generator import main
            selected_reports = [r.strip() for r in args.reports.split(',')]
            main(args.workflow, selected_reports)
            print("處理完成！")
                
        elif args.traditional:
            # 傳統模式
            print("=== 使用傳統模式 ===")
            from core.report_generator import main
            selected_reports = [r.strip() for r in args.reports.split(',')]
            main(args.traditional, selected_reports)
            print("傳統模式處理完成！")
            
        else:
            # GUI 模式 (預設)
            print("=== 啟動 GUI 介面 ===")
            launch_gui()
            
    except Exception as e:
        print(f"執行時發生錯誤: {e}")
        import traceback
        traceback.print_exc()
        input("按 Enter 鍵退出...")