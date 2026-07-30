import pandas as pd
import re
from multiprocessing import context
from datetime import datetime
from num2words import num2words
from cn2an import an2cn

def safe_float(val, default=0.0):
    """將 Excel 讀取的值安全轉為 float；字串/NaN/空值 → default"""
    if val is None:
        return default
    try:
        import pandas as pd_inner
        if pd_inner.isna(val):
            return default
    except Exception:
        pass
    try:
        return float(val)
    except (TypeError, ValueError):
        return default



# 換成百分比字串，取到小數點後第一位，如果是 0 就省略
def format_digits(val, digits=1):
    try:
        num = abs(float(val))  # 取絕對值，負號省略
        if num == int(num):
            return f"{int(num)}"
        else:
            fmt = f"{{:.{digits}f}}"
            return fmt.format(num).rstrip('0').rstrip('.')
    except:
        return str(val)

# 從字符串中提取數值，去除單位
def extract_number_from_string(value):
    """從字符串中提取數值，去除 ppt, ppts 等單位"""
    try:
        # 轉為字符串並去除常見單位
        value_str = str(value).replace("ppts", "").replace("ppt", "").replace("%", "").strip()
        return float(value_str)
    except:
        return float(value)

# QoQ, YoY 描述
def use_updown(val):
    try:
        num = float(val)
        num_str = format_digits(num)
        return f"up {num_str}" if num > 0 else f"down {num_str}"
    except:
        return ""

def use_updown_ppt(val):
    try:
        num = extract_number_from_string(val)
        num_str = format_digits(num)
        abs_num = abs(num)
        unit = "ppts" if abs_num >= 2.0 else "ppt"
        return f"up {num_str} {unit}" if num > 0 else f"down {num_str} {unit}"
    except:
        return ""

def use_verb_ed_ppt(val):
    try:
        num = extract_number_from_string(val)
        num_str = format_digits(num)
        abs_num = abs(num)
        unit = "percentage points" if abs_num >= 1.0 else "percentage point"
        return f"increased by {num_str} {unit}" if num > 0 else f"decreased by {num_str} {unit}"
    except:
        return ""
            
def use_verb_ing(val):
    try:
        num = float(val)
        num_str = format_digits(num)
        return f"increasing {num_str}" if num > 0 else f"decreasing {num_str}"
    except:
        return ""

def use_verb_ed(val):
    try:
        num = float(val)
        num_str = format_digits(num)
        return f"increased by {num_str}" if num > 0 else f"decreased by {num_str}"
    except:
        return ""

def use_noun(val):
    try:
        num = float(val)
        num_str = format_digits(num)
        return f"an increase of {num_str}" if num > 0 else f"a decrease of {num_str}"
    except:
        return ""
        
def safe_int(val, default=0):
    """將 Excel 讀取的值安全轉為 int；NaN/None/文字 → default"""
    if val is None:
        return default
    try:
        import math
        if isinstance(val, float) and math.isnan(val):
            return default
    except Exception:
        pass
    try:
        return int(round(float(val)))
    except (TypeError, ValueError):
        return default


def get_connector(qoq_val, yoy_val):
    try:
        if float(qoq_val) * float(yoy_val) < 0:
            return "but "
        else:
            return "and "
    except (TypeError, ValueError):
        return "and "
    
def use_chinese_updown(key=None, val=None):
    try:
        num = float(val)
        num_str = format_digits(num)
        if any(keyword in key for keyword in ["expense", "licensing", "royalty"]):
            return f"增加 {num_str}" if num > 0 else f"減少 {num_str}"
        elif "margin" in key:
            return f"上升 {num_str}" if num > 0 else f"下降 {num_str}"
        else:
            return f"成長 {num_str}" if num > 0 else f"衰退 {num_str}"
    except:
        return ""

class TranscriptParser:
    def __init__(self, lang, this_quarter=None, this_year=None, config=None):
        self.lang = lang
        self.this_quarter = this_quarter
        self.this_year = this_year
        self.config = config

    def parse_financial_results(self, df):
        item = {}
        def num2chinesewords(num):
            units = ["", "拾", "佰", "仟"]
            big_units = ["", "萬", "億", "兆"]
            
            # 將浮點乘上千，代表「千元」轉為「元」
            num = safe_int(round(safe_float(num) * 1000))
            num_str = str(num).zfill(((len(str(num)) + 3) // 4) * 4)  # 補0到4的倍數
            
            result = []
            sections = [num_str[i:i+4] for i in range(0, len(num_str), 4)]
            
            for i, section in enumerate(sections):
                part = ""
                zero_flag = False
                for j, ch in enumerate(section):
                    if ch != "0":
                        if zero_flag:
                            part += "0"  # 保留0數字
                            zero_flag = False
                        part += ch + units[3 - j]
                    else:
                        zero_flag = True
                part = part.rstrip("0")
                if part:
                    result.append(part + big_units[len(sections) - i - 1])
            final = "".join(result)
            # 特殊處理：開頭是零的情況
            if final.startswith("0"):
                final = final[1:]
            return final
    
        def add_amount_to_item(num, lang):
            if lang == "en":
                amount_en = num2words(round(num / 1000), lang="en") + " million"
                return amount_en
            else:
                amount_zh = num2chinesewords(num)
                return amount_zh
            
        for row_index, row in df.iterrows():
            row_index = str(row_index)

            # 先刪除括號及括號內的內容
            key = re.sub(r'\([^)]*\)', '', row_index)  # 刪除小括號 ()
            key = re.sub(r'\[[^\]]*\]', '', key)  # 刪除中括號 []
            key = re.sub(r'\{[^}]*\}', '', key)  # 刪除大括號 {}
            key = re.sub(r'（[^）]*）', '', key)  # 刪除中文括號 （）
            
            # 然後處理其他字符替換
            key = key.strip().lower().replace(" ", "_").replace("\xa0", "").replace("*", "").replace("-", "_")
            key = key.strip()  # 最後清理前後空格
            
            
            # Assign Values: Transform and Store in English/Chinese
            val = row.iloc[0]
            qoq = row.iloc[2]
            yoy = row.iloc[4]

            qoq_val = safe_float(qoq)*100
            yoy_val = safe_float(yoy)*100

            if key in ["revenue", "operating_expenses", "operating_income", "net_income"]:
                if self.lang == "en":
                    item[f"{key}"] = add_amount_to_item(val, self.lang)
                    item[f"{key}_abbv"] = f"{round(val / 1000):,} mil"

                    if key in ["revenue", "operating_expenses"]:    
                        item[f"{key}_qoq"] = use_updown(qoq_val)
                        connector = get_connector(qoq_val, yoy_val)
                        item[f"{key}_yoy"] = connector + use_updown(yoy_val)
                    elif key in ["operating_income", "net_income"]:
                        item[f"{key}_qoq"] = use_noun(qoq_val)
                        connector = get_connector(qoq_val, yoy_val)
                        item[f"{key}_yoy"] = connector + use_noun(yoy_val)
                else:
                    item[f"{key}"] = add_amount_to_item(val, self.lang)
                    item[f"{key}_qoq"] = use_chinese_updown(key, qoq_val)
                    item[f"{key}_yoy"] = use_chinese_updown(key, yoy_val)                   
            elif key == "operating_margin":
                margin_val = safe_float(val)*100
                item[f"{key}"] = format_digits(margin_val)
                # 使用新函數提取數值，去掉單位
                qoq_val = extract_number_from_string(qoq)
                yoy_val = extract_number_from_string(yoy)

                if self.lang == "en":
                    item[f"{key}_qoq"] = use_verb_ed_ppt(qoq_val)
                    connector = get_connector(qoq_val, yoy_val)
                    item[f"{key}_yoy"] = connector + use_verb_ed_ppt(yoy_val)
                else:
                    item[f"{key}_qoq"] = use_chinese_updown(key, qoq_val)
                    item[f"{key}_yoy"] = use_chinese_updown(key, yoy_val)
            elif key == "eps":
                item[f"{key}"] = val
            else:
                continue # 其他 key 不處理，直接跳到下一行
        return item

    def parse_revenue_streams(self, df):
        item = {}
        for row_index, row in df.iterrows():
            row_index = str(row_index)
            key = row_index.strip().lower().replace(" ", "_").replace("\xa0", "").replace("*", "")

            # Assign Values
            if key in ["licensing", "royalty"]:
                val = safe_float(row.iloc[8])*100
                item[f"{key}"] = format_digits(val)
                # 年度累積（Q2~Q4才有）
                if self.this_quarter in ["Q2", "Q3", "Q4"]:
                    val_ytd = safe_float(row.iloc[9])*100
                    item[f"{key}_ytd"] = format_digits(val_ytd)

            # Assign QoQ and YoY
            qoq = row.iloc[2]
            yoy = row.iloc[4]

            qoq_val = safe_float(qoq)*100
            yoy_val = safe_float(yoy)*100

            if self.lang == "en":
                if key == "licensing":
                    item[f"{key}_qoq"] = use_updown(qoq_val)
                    connector = get_connector(qoq_val, yoy_val)
                    item[f"{key}_yoy"] = connector + use_updown(yoy_val)
                elif key == "royalty":
                    item[f"{key}_qoq"] = use_verb_ing(qoq_val)
                    connector = get_connector(qoq_val, yoy_val)
                    item[f"{key}_yoy"] = connector + use_verb_ing(yoy_val)
                else:
                    item[f"{key}_qoq"] = use_verb_ed(qoq_val)
                    connector = get_connector(qoq_val, yoy_val)
                    item[f"{key}_yoy"] = connector + use_verb_ed(yoy_val)
            else:
                item[f"{key}_qoq"] = use_chinese_updown(key, qoq_val)
                item[f"{key}_yoy"] = use_chinese_updown(key, yoy_val)

            # 年度累積（Q2~Q4才有）
            if self.this_quarter in ["Q2", "Q3", "Q4"]:
                yoy_ytd = row.iloc[7]      
                yoy_ytd_val = safe_float(yoy_ytd)*100         
                
                if self.lang == "en":
                    if key == "licensing":
                        item[f"{key}_yoy_ytd"] = use_updown(yoy_ytd_val)
                    elif key == "royalty":
                        item[f"{key}_yoy_ytd"] = use_verb_ing(yoy_ytd_val)
                    else:
                        item[f"{key}_yoy_ytd"] = use_verb_ed(yoy_ytd_val)
                else:
                    item[f"{key}_yoy_ytd"] = use_chinese_updown(key, yoy_ytd_val)
        return item

    def parse_tech(self, df):
        item = {}
        for row_index, row in df.iterrows():
            row_index = str(row_index)
            key = row_index.strip().lower().replace(" ", "_").replace("\xa0", "").replace("*", "")
            key = re.sub(r'[^a-z0-9_]', '', key)
            if pd.isna(key) or key == "nan":
                continue  # 跳過這一行

            # Assign Values
            total = row.iloc[0]
            licensing = row.iloc[3]
            royalty = row.iloc[6]

            total_val = safe_float(total)*100
            licensing_val = safe_float(licensing)*100
            royalty_val = safe_float(royalty)*100

            if key in ["neobit", "neofuse", "pufbased", "mtp"]:
                item[f"{key}_total"] = format_digits(total_val)
                item[f"{key}_licensing"] = format_digits(licensing_val)
                if key == "pufbased":
                        if royalty <= 1:
                            if self.lang == "en":
                                item[f"{key}_royalty"] = "less than 1"
                            else:
                                item[f"{key}_royalty"] = "小於 1"
                else:
                    item[f"{key}_royalty"] = format_digits(royalty_val)

            licensing_qoq = row.iloc[4]
            licensing_yoy = row.iloc[5] 

            royalty_qoq = row.iloc[7]
            royalty_yoy = row.iloc[8]

            licensing_qoq_val = safe_float(licensing_qoq)*100
            licensing_yoy_val = safe_float(licensing_yoy)*100
            royalty_qoq_val = safe_float(royalty_qoq)*100
            royalty_yoy_val = safe_float(royalty_yoy)*100

            
            # QoQ and YoY
            if self.lang == "en":
                if key == "neobit":
                    item[f"{key}_licensing_qoq"] = use_verb_ing(licensing_qoq_val)
                    connector = get_connector(licensing_qoq_val, licensing_yoy_val)
                    item[f"{key}_licensing_yoy"] = connector + use_verb_ing(licensing_yoy_val)

                    item[f"{key}_royalty_qoq"] = use_updown(royalty_qoq_val)
                    connector = get_connector(royalty_qoq_val, royalty_yoy_val)
                    item[f"{key}_royalty_yoy"] = connector + use_verb_ing(royalty_yoy_val)
                elif key == "neofuse":
                    item[f"{key}_licensing_qoq"] = use_updown(licensing_qoq_val)
                    connector = get_connector(licensing_qoq_val, licensing_yoy_val)
                    item[f"{key}_licensing_yoy"] = connector + use_updown(licensing_yoy_val)

                    item[f"{key}_royalty_qoq"] = use_verb_ed(royalty_qoq_val)
                    connector = get_connector(royalty_qoq_val, royalty_yoy_val)
                    item[f"{key}_royalty_yoy"] = connector + use_verb_ed(royalty_yoy_val)
                elif key == "pufbased":
                    item[f"{key}_licensing_qoq"] = use_verb_ing(licensing_qoq_val)
                    connector = get_connector(licensing_qoq_val, licensing_yoy_val)
                    item[f"{key}_licensing_yoy"] = connector + use_verb_ing(licensing_yoy_val)

                    item[f"{key}_royalty_qoq"] = use_updown(royalty_qoq_val)
                    connector = get_connector(royalty_qoq_val, royalty_yoy_val)
                    item[f"{key}_royalty_yoy"] = connector + use_verb_ing(royalty_yoy_val)                   
                else:
                    item[f"{key}_licensing_qoq"] = use_verb_ing(licensing_qoq_val)
                    connector = get_connector(licensing_qoq_val, licensing_yoy_val)
                    item[f"{key}_licensing_yoy"] = connector + use_verb_ing(licensing_yoy_val)

                    item[f"{key}_royalty_qoq"] = use_updown(royalty_qoq_val)
                    connector = get_connector(royalty_qoq_val, royalty_yoy_val)
                    item[f"{key}_royalty_yoy"] = connector + use_verb_ed(royalty_yoy_val)
            else:
                item[f"{key}_licensing_qoq"] = use_chinese_updown(key, licensing_qoq_val)
                item[f"{key}_licensing_yoy"] = use_chinese_updown(key, licensing_yoy_val)

                item[f"{key}_royalty_qoq"] = use_chinese_updown(key, royalty_qoq_val)
                item[f"{key}_royalty_yoy"] = use_chinese_updown(key, royalty_yoy_val)

            # 年度累積（Q2~Q4才有）
            if self.this_quarter in ["Q2", "Q3", "Q4"]:
                ratio_ytd = row.iloc[9]
                ratio_ytd_val = safe_float(ratio_ytd)*100
                item[f"{key}_ytd"] = format_digits(ratio_ytd_val)

                licensing_yoy_ytd = row.iloc[12]
                royalty_yoy_ytd = row.iloc[14]

                licensing_yoy_ytd_val = safe_float(licensing_yoy_ytd)*100
                royalty_yoy_ytd_val = safe_float(royalty_yoy_ytd)*100

                if self.lang == "en":
                    item[f"{key}_licensing_yoy_ytd"] = use_verb_ed(licensing_yoy_ytd_val)
                    connector = get_connector(licensing_yoy_ytd_val, royalty_yoy_ytd_val)
                    item[f"{key}_royalty_yoy_ytd"] = connector + 'royalty ' + use_verb_ed(royalty_yoy_ytd_val)
                else:
                    item[f"{key}_licensing_yoy_ytd"] = use_chinese_updown(key, licensing_yoy_ytd_val)
                    item[f"{key}_royalty_yoy_ytd"] = use_chinese_updown(key, royalty_yoy_ytd_val)
        return item

    def parse_wafer_size(self, df):
        item = {}
        for row_index, row in df.iterrows():
            row_index = str(row_index)
            # key 特殊處理
            key = row_index.lower().strip()
            match = re.search(r'(\d+)[-_ ]?inch', key)
            if match:
                num = int(match.group(1))
                word = num2words(num)
                key = re.sub(rf'{num}[-_ ]?inch', f'{word}_inch', key)
            key = key.replace("-", "_").replace(" ", "_")
            
            if pd.isna(key) or key == "nan":
                continue  # 跳過這一行

            val = row.iloc[0]
            ratio_val = safe_float(val)*100

            item[f"{key}"] = format_digits(ratio_val)

            qoq = row.iloc[1]
            yoy = row.iloc[2]

            qoq_val = safe_float(qoq)*100
            yoy_val = safe_float(yoy)*100

            if self.lang == "en":
                item[f"{key}_qoq"] = use_updown(qoq_val)
                connector = get_connector(qoq_val, yoy_val)
                item[f"{key}_yoy"] = connector + use_updown(yoy_val)
            else:
                item[f"{key}_qoq"] = use_chinese_updown(key, qoq_val)
                item[f"{key}_yoy"] = use_chinese_updown(key, yoy_val)
        return item

    def parse_opening_remarks(self, df):
        item = {}
        for row_index, row in df.iterrows():
            row_index = str(row_index)
            key = f"{str(row_index).strip().lower().replace(' ', '_')}"
            if self.lang == "en":
                v = row.get("content_en", "")
            else:
                v = row.get("content_zh", "")
            item[f"{key}"] = v
        return item

    def parse_chairman_remarks(self, df):
        paragraphs = []
        for row_index, row in df.iterrows():
            row_index = str(row_index)
            if self.lang == "en":
                content = row.get("content_en", "")
            else:
                content = row.get("content_zh", "")
            paragraphs.append({
                "page": row_index,
                "title": row.get("title", ""),
                "content": content
            })
        return paragraphs


    def parse_new_tapeouts(self, df):
        item = {}
        total_value = df.columns[0]
        # 如果是數字，則添加千分位分隔符
        try:
            if isinstance(total_value, (int, float)):
                item["total"] = f"{safe_int(total_value):,}"
            else:
                # 嘗試轉換為數字
                num_value = safe_int(total_value)
                item["total"] = f"{num_value:,}"
        except (ValueError, TypeError):
            # 如果無法轉換為數字，保持原樣
            item["total"] = total_value
        return item


class ManagementReportParser(TranscriptParser):    
    def parse_new_tapeouts(self, df, history_file_path=None, should_save_history=True):
        import json
        import os

        process_mapping = {
            "3nm": "3nm",
            "4nm": "5nm",
            "5nm": "5nm",
            "6nm": "7nm",
            "7nm": "7nm",
            "12nm": "12nm/14nm/16nm",
            "16nm": "12nm/14nm/16nm",
            "22nm": "22nm/28nm",
            "28nm": "22nm/28nm",
            "40nm": "40nm",
            "55nm": "55nm/65nm",
            "65nm": "55nm/65nm",
            "80nm": "80nm/90nm",
            "90nm": "80nm/90nm",
        }
        
        item = {}
        summary = []

        # df 處理
        df = df.fillna(0).astype(int)
        
        # 讀取歷史總數
        history_total = 0
        if history_file_path and os.path.exists(history_file_path):
            try:
                with open(history_file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    history_total = data.get("total", 0)
            except Exception as e:
                print(f"讀取歷史數據失敗: {e}")
                history_total = 0
        

        # 遍歷每個列（process）
        for col_name in df.columns:

            process = process_mapping.get(col_name, col_name)

            if col_name == "Total":
                continue  # 跳過 Total 列

            # 取得該列中非零值的 row_index，並去掉nan
            non_zero_mask = (df[col_name] > 0) & df[col_name].notna()
            non_zero_rows = df[non_zero_mask].index.tolist()
            # 進一步過濾掉 row_index 本身是 nan 的情況
            non_zero_rows = [row for row in non_zero_rows if pd.notna(row) and str(row).lower() != 'nan']

            # 計算總計，排除 NaN 值
            total = safe_int(df.loc[non_zero_rows, col_name].sum())

            if total > 0:
                # 生成描述文字，只顯示應用名稱，最後一個前加 "and"
                app_names = [str(row) for row in non_zero_rows]
                if len(app_names) > 1:
                    app_descriptions_str = ", ".join(app_names[:-1]) + " and " + app_names[-1]
                else:
                    app_descriptions_str = app_names[0] if app_names else ""
                
                if process == "DRAM":
                    summary.append(f"{total} tape-out{'s' if total > 1 else ''} for {app_descriptions_str}")
                elif isinstance(process, str):
                    summary.insert(0, f"{total} tape-out{'s' if total > 1 else ''} at {process} for {app_descriptions_str}")

        # 計算當前季度總計
        current_total = safe_int(df.iloc[0, 0:].sum())
        
        # 根據 should_save_history 決定如何處理歷史總數
        if should_save_history and history_file_path:
            # 需要更新歷史數據：累加當前季度到歷史總數
            new_history_total = safe_int(history_total + current_total)
            try:
                import os
                os.makedirs(os.path.dirname(history_file_path), exist_ok=True)
                with open(history_file_path, 'w', encoding='utf-8') as f:
                    json.dump({"total": safe_int(new_history_total)}, f, ensure_ascii=False, indent=2)
                print(f"NTO總數已更新: {new_history_total:,} (保存到: {history_file_path})")
            except Exception as e:
                print(f"保存NTO數據失敗: {e}")
        else:
            # 不需要更新歷史數據：使用現有的歷史總數（本季數據已經包含在內）
            new_history_total = safe_int(history_total)
            print(f"本季數據已更新過，本次不再更新NTO總數 (維持: {history_total:,})")

        item["total"] = f"{current_total:,}"
        item["summary"] = summary
        item["total_utd"] = f"{new_history_total:,}"
        return item
    

    def parse_operating_results(self, df):
        item = {}
        for row_index, row in df.iterrows():
            row_index = str(row_index)
            
            # 先刪除括號及括號內的內容
            key = re.sub(r'\([^)]*\)', '', row_index)  # 刪除小括號 ()
            key = re.sub(r'\[[^\]]*\]', '', key)  # 刪除中括號 []
            key = re.sub(r'\{[^}]*\}', '', key)  # 刪除大括號 {}
            key = re.sub(r'（[^）]*）', '', key)  # 刪除中文括號 （）
            
            # 然後處理其他字符替換
            key = key.strip().lower().replace(" ", "_").replace("\xa0", "").replace("*", "").replace("-", "_")
            key = key.strip()  # 最後清理前後空格
            
            # Assign Values
            val = row.iloc[1]
            qoq = row.iloc[5]
            yoy = row.iloc[7]

            qoq_val = safe_float(qoq)*100
            yoy_val = safe_float(yoy)*100


            if key in ["net_revenue", "operating_expenses", "net_profit_shareholders", "interest_income"]:
                # 加入千分位分隔符
                val_millions = val/1000
                item[f"{key}"] = f"{val_millions:,.2f} million"

                if key == "net_profit_shareholders":
                    item[f"{key}_qoq"] = use_noun(qoq_val)
                    connector = get_connector(qoq_val, yoy_val)
                    item[f"{key}_yoy"] = connector + use_noun(yoy_val)
                elif key != "interest_income":
                    item[f"{key}_qoq"] = use_updown(qoq_val)
                    connector = get_connector(qoq_val, yoy_val)
                    item[f"{key}_yoy"] = connector + use_updown(yoy_val)

            elif key in ["non_operating_items", "net_foreign_exchange"]:
                # 判斷是 gain 還是 loss
                if val >= 0:
                    gain_loss = "gain"
                else:
                    gain_loss = "loss"
                    val = abs(val)  # 取絕對值顯示

                num = val/1000

                if key == "non_operating_items":
                    item[f"{key}"] = f"{gain_loss} for the quarter was NT$ {num:,.2f} million"
                elif key == "net_foreign_exchange":
                    item[f"{key}"] = f"a foreign exchange {gain_loss} of NT$ {num:,.2f} million"

            elif key == "operating_margin":
                margin_val = safe_float(val)*100
                item[f"{key}"] = format_digits(margin_val)

                item[f"{key}_qoq"] = use_updown_ppt(qoq_val)
                connector = get_connector(qoq_val, yoy_val)
                item[f"{key}_yoy"] = connector + use_updown_ppt(yoy_val)

            elif key == "eps":
                item[f"{key}"] = val
            else:
                continue # 其他 key 不處理，直接跳到下一行
        
        net_revenue = safe_float(df.iloc[1, 1])
        operating_expenses = safe_float(df.iloc[8, 1])
        if net_revenue != 0:
            operating_expenses_ratio = (operating_expenses / net_revenue) * 100
            item["operating_expenses_ratio"] = format_digits(operating_expenses_ratio)
        else:
            item["operating_expenses_ratio"] = ""

        # effective_tax_rate_ytd 只有 Q2~Q4 才有年度累積數據
        if self.this_quarter in ["Q2", "Q3", "Q4"]:
            effective_tax_rate_ytd = safe_float(df.iloc[20, 1]) * 100
            item["effective_tax_rate_ytd"] = format_digits(effective_tax_rate_ytd)
        else:
            item["effective_tax_rate_ytd"] = ""
        return item
    
    def parse_financial_condition(self, df):
        def use_verb_ed_million(val):
            try:
                num = abs(float(val/1000))
                num_str = f"{num:,.2f}"
                return f"increased by NT$ {num_str} million" if val > 0 else f"decreased by NT$ {num_str} million"
            except:
                return ""
            
        def use_noun_million(val):
            try:
                num = abs(float(val/1000))
                num_str = f"{num:,.2f}"
                return f"an increase of NT$ {num_str} million" if val > 0 else f"a decrease of NT$ {num_str} million"
            except:
                return ""
            
        item = {}
        for row_index, row in df.iterrows():
            row_index = str(row_index)
            
            # 先刪除括號及括號內的內容
            key = re.sub(r'\([^)]*\)', '', row_index)  # 刪除小括號 ()
            key = re.sub(r'\[[^\]]*\]', '', key)  # 刪除中括號 []
            key = re.sub(r'\{[^}]*\}', '', key)  # 刪除大括號 {}
            key = re.sub(r'（[^）]*）', '', key)  # 刪除中文括號 （）
            
            # 然後處理其他字符替換
            key = key.strip().lower().replace(" ", "_").replace("\xa0", "").replace("*", "").replace("-", "_")
            key = key.strip()  # 最後清理前後空格
            
            # Assign Values
            val = row.iloc[1]
            val_previous_quarter = row.iloc[2]
            qoq = row.iloc[4]

            num = val/1000
            num_previous_quarter = val_previous_quarter / 1000

            if key in ["cash", "total_current_assets", "total_current_liabilities", "net_working_capital"]:
                
                item[f"{key}"] = f"{num:,.2f} million"

                if key == "total_current_liabilities":
                    item[f"{key}_previous_quarter"] = f"{num_previous_quarter:,.2f} million"
                    item[f"{key}_qoq"] = use_verb_ed_million(qoq)
                elif key in ["cash", "total_current_liabilities"]:
                    item[f"{key}_qoq"] = use_verb_ed_million(qoq)
                elif key == "total_current_assets":
                    item[f"{key}_qoq"] = use_noun_million(qoq)
            elif key == "current_ratio":
                item[f"{key}"] = format_digits(val)
            else:
                continue
        return item
    
    def parse_annual_cash_flow(self, df):
        item = {}
        previous_row = None

        for row_index, row in df.iterrows():
            row_index = str(row_index)

            # 先刪除括號及括號內的內容
            key = re.sub(r'\([^)]*\)', '', row_index)  # 刪除小括號 ()
            key = re.sub(r'\[[^\]]*\]', '', key)  # 刪除中括號 []
            key = re.sub(r'\{[^}]*\}', '', key)  # 刪除大括號 {}
            key = re.sub(r'（[^）]*）', '', key)  # 刪除中文括號 （）
            
            # 然後處理其他字符替換
            key = key.strip().lower().replace(" ", "_").replace("\xa0", "").replace("*", "").replace("-", "_")
            # 去掉結尾的斜線
            key = key.rstrip("/")
            key = key.strip()  # 最後清理前後空格

            if key in ["total_operating_sources", "income_before_income_tax", "depreciation_and_amortization", "other_operating_sources", "net_investing_sources", "net_financing_sources", "net_cash_position_changes", "ending_cash_balance"]:
                # Assign Values
                val = row.iloc[1]
                num = val/1000
                num = abs(num)

                if key == "ending_cash_balance":
                    # 如果有前一行數據，計算 beginning_cash_balance
                    if previous_row is not None:
                        net_cash_position_changes = previous_row.iloc[1] / 1000
                        beginning_cash_balance_val = num - net_cash_position_changes
                        # 顯示變動為 increase 或 decrease
                        change_label = "increased from" if net_cash_position_changes > 0 else "decreased from"
                        item["beginning_cash_balance"] = f"{change_label} NT$ {beginning_cash_balance_val:,.2f} million"
                        item[f"{key}"] = f"{num:,.2f} million"

                elif key in ["other_operating_sources", "net_investing_sources", "net_financing_sources"]:
                    # 判斷是 gain 還是 loss
                    if val >= 0:
                        inflow_outflow = "generated from"
                    else:
                        inflow_outflow = "used in"

                    if key == "other_operating_sources":
                        item[f"{key}"] = f"{num:,.2f} million {inflow_outflow} other operating activities"
                    elif key == "net_investing_sources":
                        item[f"{key}"] = f"{inflow_outflow} investing activities was NT$ {num:,.2f} million"
                    elif key == "net_financing_sources":
                        item[f"{key}"] = f"{inflow_outflow} financing activities was NT$ {num:,.2f} million"
                else:
                    item[f"{key}"] = f"{num:,.2f} million"
                
                # 保存當前行為下一次迭代的前一行
                previous_row = row

            else:
                continue
        return item

    def parse_remaining_information(self, df, sheet_name=None):
        item = {}
        if sheet_name == "new_tech_licenses":
            item["total"] = len(df)
            return item
        elif sheet_name == "new_tech_platform":
            key = str(df.index[1]).lower()
            developing_num = df.loc[df.index[1], df.columns[-1]]  # 第二列，最後一欄
            # 確保轉換為Python原生類型
            if pd.isna(developing_num):
                item[f"{key}"] = 0
            else:
                item[f"{key}"] = safe_int(developing_num) if str(developing_num).replace('.', '').isdigit() else str(developing_num)
            return item
        else:
            # 計算 eMemory 總人數：優先使用 ememory_sum 欄，若不存在則加總 ememory + ememory_jp
            if "ememory_sum" in df.columns:
                total_series = df.loc["Total", "ememory_sum"]
                mask = df.index.astype(str).str.contains("R", na=False)
                rd = df.loc[mask, "ememory_sum"].sum()
            else:
                ememory_cols = [c for c in df.columns if c.startswith("ememory")]
                total_series = df.loc["Total", ememory_cols].fillna(0).sum()
                mask = df.index.astype(str).str.contains("R", na=False)
                rd = df.loc[mask, ememory_cols].fillna(0).sum().sum()
            item["total"] = safe_int(total_series)
            item["rd"] = safe_int(rd)

            ratio = (rd / item["total"])*100 if item["total"] != 0 else 0
            item["rd_ratio"] = format_digits(ratio)
            return item
        return None
    
    def get_tapeouts_summary(self, history_file_path=None):
        """
        獲取 tape-outs 歷史總數
        """
        import json
        import os
        
        if not history_file_path or not os.path.exists(history_file_path):
            return {"total": 0, "error": "歷史數據文件不存在"}
        
        try:
            with open(history_file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            return {"total": data.get("total", 0)}
            
        except Exception as e:
            return {"total": 0, "error": f"讀取歷史數據失敗: {e}"}
    
