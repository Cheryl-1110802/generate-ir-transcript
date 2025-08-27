import json
import pandas as pd
import re
import os
import sys

from jinja2 import StrictUndefined
from docxtpl import DocxTemplate
from datetime import datetime, date
from num2words import num2words
from .data_parser import TranscriptParser, ManagementReportParser
from .history_utils import check_if_history_already_updated, mark_history_as_updated

# 使用絕對路徑 import configs.constant
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from configs import constant

def get_resource_path(relative_path):
    """獲取資源檔案的絕對路徑，支援開發環境和打包後的 exe 環境"""
    try:
        # PyInstaller 打包後的臨時資料夾路徑
        base_path = sys._MEIPASS
    except Exception:
        # 開發環境中的路徑
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
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
        script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        output_dir = os.path.join(script_dir, "output")
    
    # 確保輸出目錄存在
    os.makedirs(output_dir, exist_ok=True)
    return output_dir

def format_date(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    day = dt.day
    suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
    return dt.strftime(f"%B {day}{suffix}, %Y")

def format_date_zh(date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return f"{dt.year}年{dt.month}月{dt.day}日"

def get_previous_quarter_end_date(this_quarter, this_year):
    """計算上一季的最後一天日期"""
    quarter_to_month = {
        "Q1": 12,  # Q1的上一季是Q4，結束於12月31日
        "Q2": 3,   # Q2的上一季是Q1，結束於3月31日
        "Q3": 6,   # Q3的上一季是Q2，結束於6月30日
        "Q4": 9    # Q4的上一季是Q3，結束於9月30日
    }
    # 獲取上一季結束月份
    end_month = quarter_to_month.get(this_quarter)
    if not end_month:
        return None
    # 如果是Q1，年份要減1（因為上一季是去年的Q4）
    year = int(this_year) - 1 if this_quarter == "Q1" else int(this_year)
    
    # 根據月份確定最後一天
    if end_month in [3, 12]:  # 3月和12月都是31天
        last_day = 31
    elif end_month == 6:      # 6月是30天
        last_day = 30
    elif end_month == 9:      # 9月是30天
        last_day = 30
    return f"{year}-{end_month:02d}-{last_day:02d}"

def create_transcript_context(config, lang, this_quarter):
    # quarter_map 設定
    quarter_map = {
        "en": {
            "Q1": "first-quarter",
            "Q2": "second-quarter",
            "Q3": "third-quarter",
            "Q4": "fourth-quarter"
        },
        "zh": {
            "Q1": "第一季",
            "Q2": "第二季",
            "Q3": "第三季",
            "Q4": "第四季"
        }
    }
    # 年度累積（Q2~Q4才有）
    ytd_map = {
        "en": {
            "Q2": "first half",
            "Q3": "first three quarters",
            "Q4": "full year"
        },
        "zh": {
            "Q2": "上半年",
            "Q3": "前三季",
            "Q4": "全年"
        }
    }
    context = {}
    if lang == "en":
        # 自動組合 event_quarter 例如 2Q25
        quarter_num = str(config["this_quarter"]).replace("Q", "")
        year_short = str(config["this_year"])[-2:]
        context["event_quarter"] = f"{quarter_num}Q{year_short}"
        context["event_date"] = format_date(config["event_date"])
        context["this_quarter_en"] = quarter_map["en"].get(this_quarter, this_quarter)
        context["ytd"] = ytd_map["en"].get(this_quarter, "")
        context["chairman"] = constant.CHAIRMAN_EN
        context["president"] = constant.PRESIDENT_EN
        context["financial_officer"] = constant.FINANCIAL_OFFICER_EN
        context["chairman_name"] = constant.CHAIRMAN_FIRST_NAME
        context["president_name"] = constant.PRESIDENT_FIRST_NAME
    else:
        context["event_date_zh"] = format_date_zh(config["event_date"])
        context["this_quarter_zh"] = quarter_map["zh"].get(this_quarter, this_quarter)
        context["ytd"] = ytd_map["zh"].get(this_quarter, "")
        context["chairman"] = constant.CHAIRMAN
        context["president"] = constant.PRESIDENT
        context["financial_officer"] = constant.FINANCIAL_OFFICER
    return context

def create_management_context(config, this_year, this_quarter):
    quarter_map = {
        "Q1": "first quarter",
        "Q2": "second quarter",
        "Q3": "third quarter",
        "Q4": "fourth quarter"
    }
    fiscal_quarter_map = {
        "Q1": "First Fiscal Quarter",
        "Q2": "Second Fiscal Quarter",
        "Q3": "Third Fiscal Quarter",
        "Q4": "Fourth Fiscal Quarter"
    }
    quarter_end_date_map = {
        "Q1": "3-31",
        "Q2": "6-30",
        "Q3": "9-30",
        "Q4": "12-31"
    }
    ytd_map = {
        "Q2": "H1",
        "Q3": "Q1-Q3",
        "Q4": "FY"
    }
    # Management report 專用的 context
    context = {}
    context["event_date"] = format_date(config["event_date"])
    context["this_quarter_en"] = quarter_map.get(this_quarter, this_quarter)
    context["ytd"] = ytd_map.get(this_quarter, "")
    context["this_fiscal_quarter"] = fiscal_quarter_map.get(this_quarter, "")
    this_quarter_end_date = quarter_end_date_map.get(this_quarter, "")
    context["this_quarter_end_date"] = format_date(f"{this_year}-{this_quarter_end_date}")
    previous_quarter_end = get_previous_quarter_end_date(this_quarter, this_year)
    context["previous_quarter_end_date"] = format_date(previous_quarter_end)
    return context

def build_context(config, xls, this_quarter, this_year, parser, should_save_history=True):
    # 這裡 parser 可能是 TranscriptParser 或 ManagementReportParser
    shared_context = {}
    transcript_context = {}
    management_context = {}

    # 共用欄位
    shared_context = {
        "this_quarter": config["this_quarter"],
        "this_year": config["this_year"],
        "this_quarter_year": f"{this_quarter} {this_year}",
        "previous_year": str(int(this_year) - 1)
    }
    
    transcript_context.update(shared_context)
    management_context.update(shared_context)

    # 由 parser 產生各自專屬欄位
    if isinstance(parser, TranscriptParser) and not isinstance(parser, ManagementReportParser):
        context = create_transcript_context(config, parser.lang, this_quarter)
        transcript_context.update(context)
    elif isinstance(parser, ManagementReportParser):
        context = create_management_context(config, this_year, this_quarter)
        management_context.update(context)

    # parse each sheet
    for sheet_name, prefix in config["sheet_mapping"].items():
        df = xls.parse(sheet_name, index_col=0, header=0)
        
        if sheet_name == "future_outlook":
            # 使用同樣的 parser，不需要分別判斷
            parsed = parser.parse_opening_remarks(df)
            transcript_context[sheet_name] = parsed
            management_context[sheet_name] = parsed
                
        elif sheet_name == "new_tapeouts":
            # 兩種 parser 都處理，但可能有不同的實作
            if isinstance(parser, TranscriptParser) and not isinstance(parser, ManagementReportParser):
                parsed = parser.parse_new_tapeouts(df)
                transcript_context[sheet_name] = parsed
                
            elif isinstance(parser, ManagementReportParser):
                # 為 ManagementReportParser 添加歷史數據文件路徑
                history_file_path = os.path.join(get_output_path(), "new_tapeouts_history.json")
                parsed = parser.parse_new_tapeouts(df, history_file_path, should_save_history)  # 傳遞 should_save_history 參數
                management_context[sheet_name] = parsed
                
        elif sheet_name in ["financial_results", "revenue_streams", "tech", "wafer_size", "opening_remarks", "chairman_remarks"]:
            # 這些 sheet 只有 TranscriptParser 處理
            if isinstance(parser, TranscriptParser) and not isinstance(parser, ManagementReportParser):
                if hasattr(parser, f"parse_{sheet_name}"):
                    parsed = getattr(parser, f"parse_{sheet_name}")(df)
                    transcript_context[sheet_name] = parsed

        else:
            # 其他 sheet 由 ManagementReportParser 處理
            if isinstance(parser, ManagementReportParser):
                if sheet_name in ["operating_results", "financial_condition", "annual_cash_flow"]:
                    if hasattr(parser, f"parse_{sheet_name}"):
                        parsed = getattr(parser, f"parse_{sheet_name}")(df)
                    else:
                        parsed = parser.parse_remaining_information(df, sheet_name)  # 傳遞 sheet_name 參數
                else:
                    # 對於其他未明確定義的sheet，使用 parse_remaining_information
                    parsed = parser.parse_remaining_information(df, sheet_name)
                management_context[sheet_name] = parsed  
    return transcript_context, management_context

def fix_ampersand_in_context(data):
    """
    遞迴處理 context 中的 & 符號，確保在 docxtpl 模板渲染時正確顯示
    
    docxtpl 使用 Jinja2 + XML，& 符號會被當作 XML 特殊字符處理。
    我們需要使用 HTML 實體 &amp; 來確保正確渲染。
    """
    if isinstance(data, str):
        # 將 & 替換為 &amp; 以確保在 XML 中正確顯示
        return data.replace('&', '&amp;')
    elif isinstance(data, dict):
        return {k: fix_ampersand_in_context(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [fix_ampersand_in_context(item) for item in data]
    else:
        return data

def generate_report(context, template_path, output_path):
    doc = DocxTemplate(template_path)
    try:
        # 修復 context 中的 & 符號
        fixed_context = fix_ampersand_in_context(context)
        doc.render(fixed_context)
        doc.save(output_path)
        print(f"報告已產出：{output_path}")
        return True  # 返回成功標誌
    except Exception as e:
        print("渲染 Word 時失敗：", e)
        return False  # 返回失敗標誌    

def main(excel_path=None, selected_reports=None):
    # 使用資源路徑函數來獲取配置檔案的正確路徑
    config_path = get_resource_path("configs/config.json")
    with open(config_path, encoding='utf-8') as f:
        config = json.load(f)

    # 設定法說會季度和年份
    # 直接從 config 讀取 this_quarter 和 this_year
    this_quarter = config.get("this_quarter", "Q2")
    this_year = config.get("this_year", "2025")

    # input data - 如果有提供 excel_path 就使用它，否則使用 config 中的路徑
    if excel_path:
        xls = pd.ExcelFile(excel_path)
    else:
        # 使用資源路徑函數來獲取輸入檔案的完整路徑
        input_path = get_resource_path(config["input_excel_path"])
        xls = pd.ExcelFile(input_path)
    
    # output - 使用正確的輸出路徑函數
    today = date.today().isoformat()
    output_dir = get_output_path()
    
    # 確保輸出目錄存在
    os.makedirs(output_dir, exist_ok=True)
    
    all_outputs = [
        ("transcript_en", os.path.join(output_dir, f"{this_year} {this_quarter} Investor Conference Transcript_en_{today}.docx"), "en", TranscriptParser),
        ("transcript_zh", os.path.join(output_dir, f"{this_year} {this_quarter} Investor Conference Transcript_zh_{today}.docx"), "zh", TranscriptParser),
        ("management_report", os.path.join(output_dir, f"{this_year} {this_quarter} Management Report_{today}.docx"), "en", ManagementReportParser)
    ]
    
    # 如果指定了選擇的報告，只產生選擇的報告
    if selected_reports:
        outputs = [output for output in all_outputs if output[0] in selected_reports]
    else:
        outputs = all_outputs  # 預設產生所有報告
    
    print(f"產出檔案: {[output[0] for output in outputs]}")
    
    # 獲取法說會日期
    event_date = config.get("event_date", "")
    
    # 檢查 new_tapeouts 歷史數據是否已經針對這次法說會更新過
    history_file_path = os.path.join(get_output_path(), "new_tapeouts_history.json")
    history_already_updated_for_this_event = check_if_history_already_updated(history_file_path, event_date)
    
    if history_already_updated_for_this_event:
        print(f"✓ new_tapeouts 歷史數據已更新過")
                
        # 針對每個報告類型生成對應的輸出
    
    # 添加標誌來追踪是否已經在本次執行中更新過 new_tapeouts 歷史數據
    history_updated_this_run = False

    for name, output_path, lang, ParserClass in outputs:
        parser = ParserClass(lang, this_quarter, this_year, config)
        
        # 對於 ManagementReportParser，先嘗試不保存歷史數據生成 context
        should_save_history = False if isinstance(parser, ManagementReportParser) else True
        transcript_context, management_context = build_context(config, xls, this_quarter, this_year, parser, should_save_history)
        
        # 根據報告類型選擇對應的 context
        if isinstance(parser, TranscriptParser) and not isinstance(parser, ManagementReportParser):
            context = transcript_context
        elif isinstance(parser, ManagementReportParser):
            context = management_context
        else:
            context = transcript_context  # 預設使用 transcript_context
            
        # 依照你的模板選擇邏輯取得 template_path
        if this_quarter == "Q1":
            tpl_path = config["template_paths"][name].get("Q1")
        else:
            tpl_path = config["template_paths"][name].get("default")
        
        # 使用資源路徑函數來獲取模板的完整路徑
        tpl_path = get_resource_path(tpl_path)
        
        # 嘗試渲染 Word 文檔
        render_success = generate_report(context, tpl_path, output_path)
        
        # 如果是 ManagementReportParser 且渲染成功，執行後續處理
        if isinstance(parser, ManagementReportParser) and render_success:
            # 檢查是否需要更新歷史數據
            if not history_already_updated_for_this_event and not history_updated_this_run:
                print("正在更新 new_tapeouts 歷史數據...")
                transcript_context, management_context = build_context(config, xls, this_quarter, this_year, parser, should_save_history=True)
                context = management_context  # 更新 context 為包含保存歷史數據的版本
                
                # 標記歷史數據已更新，傳入 event_date
                mark_history_as_updated(history_file_path, event_date, this_quarter, this_year)
                history_updated_this_run = True
                print("✓ new_tapeouts 歷史數據已更新")
            else:
                if history_already_updated_for_this_event:
                    print(f"⚠️ new_tapeouts 歷史數據已更新過")
                else:
                    print("⚠️ new_tapeouts 歷史數據在本次執行中已更新過")
            
            # 執行 Excel 到 Word 表格貼上功能
            print("正在更新表格...")
            try:
                from .excel_table_to_word import ExcelTable2Word
                integration = ExcelTable2Word()
                
                # 從檔案名稱中提取基本名稱（去掉路徑和副檔名）
                word_filename = os.path.basename(output_path)
                excel_filename = os.path.basename(excel_path) if excel_path else "latest data for transcript & management report.xlsx"
                
                # 根據您的新規格配置：
                # Q2~Q4: index=0: financial_results, index=1: revenue_streams, index=2: tech(單季A1:J8), index=3: tech(年度累計A1:A8+K1:P8), index=4: wafer_size(A1:F4), index=5: new_tech_platform(A1:P4)
                # Q1: index=0: financial_results, index=1: revenue_streams, index=2: tech(A1:J8), index=3: wafer_size(A1:D4), index=4: new_tech_platform(A1:P4)
                
                paste_configs = []
                
                if this_quarter == 'Q1':
                    # Q1 配置
                    paste_configs = [
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'financial_results',
                            'word_filename': word_filename,
                            'table_identifier': 0,  # index=0
                            'cell_range': 'A1:F10',
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'revenue_streams',
                            'word_filename': word_filename,
                            'table_identifier': 1,  # index=1
                            'cell_range': 'A1:F4',
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 2,  # index=2
                            'cell_range': 'A1:J8',  # Q1有合併儲存格
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'wafer_size',
                            'word_filename': word_filename,
                            'table_identifier': 3,  # index=3
                            'cell_range': 'A1:D4',
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'new_tech_platform',
                            'word_filename': word_filename,
                            'table_identifier': 4,  # index=4
                            'cell_range': 'A1:P4',
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        }
                    ]
                else:
                    # Q2~Q4 配置
                    paste_configs = [
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'financial_results',
                            'word_filename': word_filename,
                            'table_identifier': 0,  # index=0
                            'cell_range': 'A1:I10',
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'revenue_streams',
                            'word_filename': word_filename,
                            'table_identifier': 1,  # index=1
                            'cell_range': 'A1:I4',
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 2,  # index=2: 單季表格
                            'cell_range': 'A1:J8',
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 3,  # index=3: 年度累計表格 - rowname部分
                            'cell_range': 'A1:A8',  # rowname在A1:A8
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 3,  # index=3: 年度累計表格 - 表格內容部分
                            'cell_range': 'K1:P8',  # 表格內容在K1:P8 (有合併儲存格)
                            'start_position': (0, 1),  # 從 (0,1) 開始貼上
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'wafer_size',
                            'word_filename': word_filename,
                            'table_identifier': 4,  # index=4
                            'cell_range': 'A1:F4',  # 有合併儲存格
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'new_tech_platform',
                            'word_filename': word_filename,
                            'table_identifier': 5,  # index=5
                            'cell_range': 'A1:P4',
                            'start_position': (0, 0),  # 從 (0,0) 開始
                            'clear_existing': False
                        }
                    ]
                
                # 執行批量貼上
                results = integration.batch_paste_excel_to_word(paste_configs)
                
                success_count = sum(1 for result in results.values() if result)
                total_count = len(results)
                
                if success_count > 0:
                    print(f"✓ 表格更新完成：{success_count}/{total_count}")
                else:
                    print("⚠️ 表格更新未成功")
                    
            except Exception as e:
                print(f"⚠️ 表格更新錯誤: {e}")
                # 不影響主流程，繼續執行





