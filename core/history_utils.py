import json
import os
from datetime import date

def check_if_history_already_updated(history_file_path, event_date):
    """
    檢查歷史數據是否已經針對特定法說會日期更新過
    """
    if not os.path.exists(history_file_path):
        return False
    
    try:
        with open(history_file_path, 'r', encoding='utf-8') as f:
            history_data = json.load(f)
        
        # 檢查是否有當前法說會日期的更新記錄
        last_update = history_data.get("last_update")
        
        # 如果有 last_update 且 event_date 相同，表示已經針對這次法說會更新過
        if last_update and last_update.get("event_date") == event_date:
            return True
        
        return False
    except (json.JSONDecodeError, KeyError):
        return False

def mark_history_as_updated(history_file_path, event_date, this_quarter, this_year):
    """
    在歷史數據中標記已針對特定法說會更新
    """
    try:
        if os.path.exists(history_file_path):
            with open(history_file_path, 'r', encoding='utf-8') as f:
                history_data = json.load(f)
        else:
            history_data = {}
        
        # 添加更新記錄
        today = date.today().isoformat()
        history_data["last_update"] = {
            "event_date": event_date,  # 關鍵：記錄法說會日期
            "quarter": this_quarter,
            "year": this_year,
            "update_date": today,
            "timestamp": date.today().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        with open(history_file_path, 'w', encoding='utf-8') as f:
            json.dump(history_data, f, ensure_ascii=False, indent=2)
            
    except Exception as e:
        print(f"警告：無法標記歷史更新狀態: {e}")


