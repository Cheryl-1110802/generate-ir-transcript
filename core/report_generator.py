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
from .snapshot_utils import (
    get_stale_sheets, mark_stale_values, color_stale_paragraphs, SNAPSHOT_FILENAME
)

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
    """Current assets/liabilities 段落固定跟「前一年年底」比較（去年12月31日），
    而不是逐季比較的上一季末，所以不管 this_quarter 是哪一季都回傳去年12月31日"""
    return f"{int(this_year) - 1}-12-31"

def create_transcript_context(lang, this_quarter, this_year, event_date):
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
        quarter_num = str(this_quarter).replace("Q", "")
        year_short = str(this_year)[-2:]
        context["event_quarter"] = f"{quarter_num}Q{year_short}"
        context["event_date"] = format_date(event_date)
        context["this_quarter_en"] = quarter_map["en"].get(this_quarter, this_quarter)
        context["ytd"] = ytd_map["en"].get(this_quarter, "")
        context["chairman"] = constant.CHAIRMAN_EN
        context["president"] = constant.PRESIDENT_EN
        context["financial_officer"] = constant.FINANCIAL_OFFICER_EN
        context["chairman_name"] = constant.CHAIRMAN_FIRST_NAME
        context["president_name"] = constant.PRESIDENT_FIRST_NAME
    else:
        context["event_date_zh"] = format_date_zh(event_date)
        context["this_quarter_zh"] = quarter_map["zh"].get(this_quarter, this_quarter)
        context["ytd"] = ytd_map["zh"].get(this_quarter, "")
        context["chairman"] = constant.CHAIRMAN
        context["president"] = constant.PRESIDENT
        context["financial_officer"] = constant.FINANCIAL_OFFICER
    return context

def create_management_context(event_date, this_year, this_quarter):
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
    context["event_date"] = format_date(event_date)
    context["this_quarter_en"] = quarter_map.get(this_quarter, this_quarter)
    context["ytd"] = ytd_map.get(this_quarter, "")
    context["this_fiscal_quarter"] = fiscal_quarter_map.get(this_quarter, "")
    this_quarter_end_date = quarter_end_date_map.get(this_quarter, "")
    context["this_quarter_end_date"] = format_date(f"{this_year}-{this_quarter_end_date}")
    previous_quarter_end = get_previous_quarter_end_date(this_quarter, this_year)
    context["previous_quarter_end_date"] = format_date(previous_quarter_end)
    return context

def build_context(data_config, xls, this_quarter, this_year, parser, event_date, should_save_history=True, stale_sheets=None):
    # 這裡 parser 可能是 TranscriptParser 或 ManagementReportParser
    shared_context = {}
    transcript_context = {}
    management_context = {}

    # 共用欄位
    # use explicit quarter/year passed from quarter_config.json
    shared_context = {
        "this_quarter": this_quarter,
        "this_year": this_year,
        "this_quarter_year": f"{this_quarter} {this_year}",
        "previous_year": str(int(this_year) - 1)
    }
    
    transcript_context.update(shared_context)
    management_context.update(shared_context)

    # 由 parser 產生各自專屬欄位
    if isinstance(parser, TranscriptParser) and not isinstance(parser, ManagementReportParser):
        context = create_transcript_context(parser.lang, this_quarter, this_year, event_date)
        transcript_context.update(context)
    elif isinstance(parser, ManagementReportParser):
        context = create_management_context(event_date, this_year, this_quarter)
        management_context.update(context)

    # parse each sheet
    for sheet_name, prefix in data_config["sheet_mapping"].items():
        df = xls.parse(sheet_name, index_col=0, header=0)
        
        if sheet_name == "future_outlook":
            # 使用同樣的 parser，不需要分別判斷
            parsed = parser.parse_opening_remarks(df)
            if stale_sheets and sheet_name in stale_sheets:
                parsed = mark_stale_values(parsed)
            transcript_context[sheet_name] = parsed
            management_context[sheet_name] = parsed
                
        elif sheet_name == "new_tapeouts":
            # 兩種 parser 都處理，但可能有不同的實作
            if isinstance(parser, TranscriptParser) and not isinstance(parser, ManagementReportParser):
                parsed = parser.parse_new_tapeouts(df)
                if stale_sheets and sheet_name in stale_sheets:
                    parsed = mark_stale_values(parsed)
                transcript_context[sheet_name] = parsed
                
            elif isinstance(parser, ManagementReportParser):
                # 為 ManagementReportParser 添加歷史數據文件路徑
                history_file_path = os.path.join(get_output_path(), "new_tapeouts_history.json")
                parsed = parser.parse_new_tapeouts(df, history_file_path, should_save_history)  # 傳遞 should_save_history 參數
                if stale_sheets and sheet_name in stale_sheets:
                    parsed = mark_stale_values(parsed)
                management_context[sheet_name] = parsed
                
        elif sheet_name in ["financial_results", "revenue_streams", "tech", "wafer_size", "opening_remarks", "chairman_remarks"]:
            # 這些 sheet 只有 TranscriptParser 處理
            if isinstance(parser, TranscriptParser) and not isinstance(parser, ManagementReportParser):
                if hasattr(parser, f"parse_{sheet_name}"):
                    parsed = getattr(parser, f"parse_{sheet_name}")(df)
                    if stale_sheets and sheet_name in stale_sheets:
                        parsed = mark_stale_values(parsed)
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
                if stale_sheets and sheet_name in stale_sheets:
                    parsed = mark_stale_values(parsed)
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
    # 讀取季度／年度設定（例如 this_quarter, this_year）
    config_path = get_resource_path("configs/quarter_config.json")
    with open(config_path, encoding='utf-8') as f:
        quarter_config = json.load(f)

    # 讀取資料相關設定（例如 input_excel_path、其他路徑或參數）
    data_config_path = get_resource_path("configs/data_config.json")
    with open(data_config_path, encoding='utf-8') as f:
        data_config = json.load(f)

    # 直接從 config 讀取 this_quarter 和 this_year
    this_quarter = quarter_config["this_quarter"]
    this_year = quarter_config["this_year"]
    event_date = quarter_config.get("event_date", "")

    # input data - 如果有提供 excel_path 就使用它，否則嘗試從 scripts/document_update/input/ 取得最新的 Excel 檔
    if excel_path:
        xls = pd.ExcelFile(excel_path)
    else:
        # 嘗試在專案的 input 目錄尋找最近的 .xls/.xlsx 檔案
        script_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        input_dir = os.path.join(script_dir, "input")
        excel_candidates = []
        if os.path.isdir(input_dir):
            for name in os.listdir(input_dir):
                if name.lower().endswith(('.xls', '.xlsx')):
                    excel_candidates.append(os.path.join(input_dir, name))

        if excel_candidates:
            # 選最新修改時間的檔案
            excel_candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            chosen = excel_candidates[0]
            print(f"使用 input 資料檔: {chosen}")
            xls = pd.ExcelFile(chosen)
        else:
            raise FileNotFoundError(
                "找不到 Excel 輸入檔 (沒有傳入 excel_path，且 scripts/document_update/input/ 內無 .xls/.xlsx)。"
                " 請傳入 excel_path 或把檔案放到 scripts/document_update/input/，或在 data_config.json 中加入 'input_excel_path'。"
            )
    
    # ── 快照比對：找出未更新的 sheet（顯示為紅字）────────────────────────────
    snapshot_path = os.path.join(get_output_path(), SNAPSHOT_FILENAME)
    sheet_names = list(data_config["sheet_mapping"].keys())
    stale_sheets = get_stale_sheets(
        excel_path if excel_path else "", sheet_names,
        this_quarter, str(this_year), snapshot_path
    )
    if stale_sheets:
        print(f"⚠️ 以下 sheet 與上季快照相同，將以紅字標記: {sorted(stale_sheets)}")
    else:
        print("✓ 所有 sheet 已更新（或為首次執行），不標記紅字")

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

    
    # 檢查 new_tapeouts 歷史數據是否已經針對這次法說會更新過
    history_file_path = os.path.join(get_output_path(), "new_tapeouts_history.json")
    history_already_updated_for_this_event = check_if_history_already_updated(history_file_path, event_date)
    
    if history_already_updated_for_this_event:
        print(f"✓ new_tapeouts 歷史數據已更新過")
                
        # 針對每個報告類型生成對應的輸出
    
    # 添加標誌來追踪是否已經在本次執行中更新過 new_tapeouts 歷史數據
    history_updated_this_run = False

    for name, output_path, lang, ParserClass in outputs:
        parser = ParserClass(lang, this_quarter, this_year, data_config)
        
        # 對於 ManagementReportParser，先嘗試不保存歷史數據生成 context
        should_save_history = False if isinstance(parser, ManagementReportParser) else True
        transcript_context, management_context = build_context(data_config, xls, this_quarter, this_year, parser, event_date, should_save_history, stale_sheets=stale_sheets)
        
        # 根據報告類型選擇對應的 context
        if isinstance(parser, TranscriptParser) and not isinstance(parser, ManagementReportParser):
            context = transcript_context
        elif isinstance(parser, ManagementReportParser):
            context = management_context
        else:
            context = transcript_context  # 預設使用 transcript_context
        
        # 取得模板的完整路徑（template_paths 在 data_config 中）
        tpl_rel = data_config.get("template_paths", {}).get(name)
        if not tpl_rel:
            print(f"找不到 template_paths 中對應 {name} 的設定，跳過此報告。")
            continue
        tpl_path = get_resource_path(tpl_rel)
        
        # 嘗試渲染 Word 文檔
        render_success = generate_report(context, tpl_path, output_path)

        # 如果有 stale sheet，對渲染完成的 Word 執行紅字著色
        if render_success and stale_sheets:
            try:
                color_stale_paragraphs(output_path)
                print(f"✓ 紅字標記完成（{len(stale_sheets)} 個 stale sheet）")
            except Exception as _ce:
                print(f"⚠️ 紅字著色失敗（不影響主流程）: {_ce}")

        
        # 如果是 ManagementReportParser 且渲染成功，執行後續處理
        if isinstance(parser, ManagementReportParser) and render_success:
            # 檢查是否需要更新歷史數據
            if not history_already_updated_for_this_event and not history_updated_this_run:
                print("正在更新 new_tapeouts 歷史數據...")
                transcript_context, management_context = build_context(data_config, xls, this_quarter, this_year, parser, event_date, should_save_history=True, stale_sheets=stale_sheets)
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
                
                # 優先使用 data_config 中的 paste_configs（如果有），否則使用通用的貼上設定
                paste_configs = data_config.get("paste_configs")
                if not paste_configs:
                    paste_configs = [
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'financial_results',
                            'word_filename': word_filename,
                            'table_identifier': 0,
                            'cell_range': 'A1:I10',
                            'start_position': (0, 0),
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'revenue_streams',
                            'word_filename': word_filename,
                            'table_identifier': 1,
                            'cell_range': 'A1:I4',
                            'start_position': (0, 0),
                            'clear_existing': False
                        },
                        {
                            # Revenue Analysis (US$) - 新增的美金表格
                            'excel_filename': excel_filename,
                            'sheet_name': 'revenue_streams',
                            'word_filename': word_filename,
                            'table_identifier': 2,
                            'cell_range': 'A6:I9',
                            'start_position': (0, 0),
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 3,
                            'cell_range': 'A1:J8',
                            'start_position': (0, 0),
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 4,
                            'cell_range': 'A1:A8',
                            'start_position': (0, 0),
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 4,
                            'cell_range': 'K1:P8',
                            'start_position': (0, 1),
                            'clear_existing': False
                        },
                        {
                            # Revenue analysis by technology (US$, main) - 新增的美金表格
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 5,
                            'cell_range': 'A10:J17',
                            'start_position': (0, 0),
                            'clear_existing': False
                        },
                        {
                            # Revenue analysis by technology (US$, YTD side) - 新增的美金表格
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 6,
                            'cell_range': 'A10:A17',
                            'start_position': (0, 0),
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'tech',
                            'word_filename': word_filename,
                            'table_identifier': 6,
                            'cell_range': 'K10:P17',
                            'start_position': (0, 1),
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'wafer_size',
                            'word_filename': word_filename,
                            'table_identifier': 7,
                            'cell_range': 'A1:F4',
                            'start_position': (0, 0),
                            'clear_existing': False
                        },
                        {
                            # Wafer size (US$) - 新增的美金表格
                            'excel_filename': excel_filename,
                            'sheet_name': 'wafer_size',
                            'word_filename': word_filename,
                            'table_identifier': 8,
                            'cell_range': 'A6:F9',
                            'start_position': (0, 0),
                            'clear_existing': False
                        },
                        {
                            'excel_filename': excel_filename,
                            'sheet_name': 'new_tech_platform',
                            'word_filename': word_filename,
                            'table_identifier': 9,
                            'cell_range': 'A1:P4',
                            'start_position': (0, 0),
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





