import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
from datetime import datetime
from core.report_generator import main

def launch_gui():
    # 動態獲取版本號
    import sys
    import os
    
    # 獲取 main.py 的版本號
    main_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    version = "2.1.7"  # 預設版本
    try:
        main_path = os.path.join(main_dir, "main.py")
        with open(main_path, 'r', encoding='utf-8') as f:
            content = f.read()
            import re
            match = re.search(r'VERSION\s*=\s*["\']([^"\']+)["\']', content)
            if match:
                version = match.group(1)
    except Exception as e:
        print(f"無法讀取版本號: {e}")
        pass
    # 固定使用微軟正黑體
    chinese_font = ("Microsoft JhengHei UI", 12)
    title_font = ("Microsoft JhengHei UI", 16, "bold")
    label_font = ("Microsoft JhengHei UI", 12)
    button_font = ("Microsoft JhengHei UI", 11)
    small_font = ("Microsoft JhengHei UI", 10)
    
    def update_config():
        """更新 config.json 檔案"""
        try:
            # 獲取專案根目錄
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            config_path = os.path.join(project_root, "configs", "config.json")
            
            # 讀取現有的 config
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 更新配置
            config["this_quarter"] = quarter_var.get()
            config["this_year"] = year_var.get()
            
            # 組合日期格式 YYYY-MM-DD
            selected_date = f"{year_var.get()}-{month_var.get().zfill(2)}-{day_var.get().zfill(2)}"
            config["event_date"] = selected_date
            
            # 寫回檔案
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
            return True
        except Exception as e:
            messagebox.showerror("錯誤", f"更新配置檔案失敗: {str(e)}")
            return False
    
    def on_submit():
        # 先驗證輸入
        if not quarter_var.get():
            messagebox.showwarning("警告", "請選擇季度！")
            return
        if not year_var.get():
            messagebox.showwarning("警告", "請輸入年份！")
            return
        if not month_var.get() or not day_var.get():
            messagebox.showwarning("警告", "請選擇日期！")
            return
            
        # 更新配置檔案
        if not update_config():
            return
        
        excel_path = filedialog.askopenfilename(
            title="選擇 Excel 檔案",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if excel_path:
            # 獲取選擇的報告類型
            selected_reports = []
            if transcript_var.get():
                selected_reports.extend(["transcript_en", "transcript_zh"])
            if management_var.get():
                selected_reports.append("management_report")
            
            if not selected_reports:
                messagebox.showwarning("警告", "請至少選擇一種文件類型！")
                return
            
            # 檢查是否使用工作流程模式（已移除，使用傳統模式）
            use_workflow = False
            previous_word_path = None
            
            try:
                # 使用傳統模式
                print("使用傳統模式...")
                main(excel_path, selected_reports)
                messagebox.showinfo("成功", "文件產出成功！")
                    
            except Exception as e:
                print(f"錯誤: {str(e)}")
                import traceback
                traceback.print_exc()
                messagebox.showerror("錯誤", f"執行時發生錯誤: {str(e)}")

    root = tk.Tk()
    root.title("文件自動更新工具")
    root.geometry("550x600")
    
    # 載入現有配置
    def load_current_config():
        try:
            project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            config_path = os.path.join(project_root, "configs", "config.json")
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {
                "this_quarter": "Q2",
                "this_year": "2025",
                "event_date": "2025-08-15"
            }
    
    current_config = load_current_config()
    
    # 變數定義
    quarter_var = tk.StringVar(value=current_config.get("this_quarter", "Q2"))
    year_var = tk.StringVar(value=current_config.get("this_year", "2025"))
    
    # 解析現有日期
    try:
        current_date = datetime.strptime(current_config.get("event_date", "2025-08-15"), "%Y-%m-%d")
        month_var = tk.StringVar(value=str(current_date.month))
        day_var = tk.StringVar(value=str(current_date.day))
    except:
        month_var = tk.StringVar(value="8")
        day_var = tk.StringVar(value="15")
    
    # 標題
    title_label = tk.Label(root, text="文件自動更新工具", font=title_font)
    title_label.pack(pady=20)
    
    # 法說會資訊設定框架
    config_frame = tk.LabelFrame(root, text="法說會資訊設定", font=label_font)
    config_frame.pack(pady=10, padx=20, fill="x")
    
    # 季度選擇
    quarter_frame = tk.Frame(config_frame)
    quarter_frame.pack(fill="x", padx=15, pady=5)
    tk.Label(quarter_frame, text="季度:", font=label_font).pack(side="left")
    quarter_combo = ttk.Combobox(quarter_frame, textvariable=quarter_var, 
                                values=["Q1", "Q2", "Q3", "Q4"], 
                                state="readonly", width=10, font=chinese_font)
    quarter_combo.pack(side="left", padx=(10, 0))
    
    # 年份輸入
    year_frame = tk.Frame(config_frame)
    year_frame.pack(fill="x", padx=15, pady=5)
    tk.Label(year_frame, text="年份:", font=label_font).pack(side="left")
    year_entry = tk.Entry(year_frame, textvariable=year_var, width=10, font=chinese_font)
    year_entry.pack(side="left", padx=(10, 0))
    
    # 舉辦日期
    date_frame = tk.Frame(config_frame)
    date_frame.pack(fill="x", padx=15, pady=5)
    tk.Label(date_frame, text="舉辦日期:", font=label_font).pack(side="left")
    
    # 月份選擇
    month_combo = ttk.Combobox(date_frame, textvariable=month_var,
                              values=[str(i) for i in range(1, 13)],
                              state="readonly", width=5, font=chinese_font)
    month_combo.pack(side="left", padx=(10, 5))
    tk.Label(date_frame, text="月", font=label_font).pack(side="left")
    
    # 日期選擇
    day_combo = ttk.Combobox(date_frame, textvariable=day_var,
                            values=[str(i) for i in range(1, 32)],
                            state="readonly", width=5, font=chinese_font)
    day_combo.pack(side="left", padx=(5, 5))
    tk.Label(date_frame, text="日", font=label_font).pack(side="left")
    
    # 文件類型選擇框架
    report_frame = tk.LabelFrame(root, text="選擇要更新的文件類型", font=label_font)
    report_frame.pack(pady=10, padx=20, fill="x")
    
    # Transcript 選項
    transcript_var = tk.BooleanVar(value=True)
    transcript_check = tk.Checkbutton(
        report_frame, 
        text="法說會講稿-中英文 (Transcript-English/Chinese)", 
        variable=transcript_var,
        font=chinese_font
    )
    transcript_check.pack(anchor="w", padx=15, pady=8)
    
    # Management Report 選項
    management_var = tk.BooleanVar(value=True)
    management_check = tk.Checkbutton(
        report_frame, 
        text="管理報告 (Management Report)", 
        variable=management_var,
        font=chinese_font
    )
    management_check.pack(anchor="w", padx=15, pady=8)
    
    # 選擇檔案按鈕
    submit_button = tk.Button(
        root, 
        text="選擇 Excel 檔案並開始處理", 
        command=on_submit, 
        font=button_font,
        bg="#4CAF50", 
        fg="white", 
        padx=20, 
        pady=15
    )
    submit_button.pack(pady=30)
    
    # 版本資訊
    version_label = tk.Label(root, text=f"v{version}", 
                           font=small_font, fg="#666666")
    version_label.pack(side="bottom", pady=10)
    
    root.mainloop()