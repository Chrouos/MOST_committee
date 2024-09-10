import tkinter as tk
from tkinter import messagebox
from utils.get_setting import value_of_key
from utils.script import load_into_chroma_bge_manager, search_v3, filter_committee, excel_process_VBA, statistic_committee
from utils.get_setting import setting_data

def execute_mode(mode, current_plan):
    # 檢查 "目前執行計畫" 並設定 is_industry
    
    if current_plan == "產學合作":
        is_industry = True
    elif current_plan == "研究計畫":
        is_industry = False
    else:
        messagebox.showerror("錯誤", "「目前執行計畫」沒有填寫正確（產學合作｜研究計畫）")
        return
    
    # 根據選擇的模式執行相應的功能
    if mode == '存入資料庫':
        load_into_chroma_bge_manager(is_industry)
        messagebox.showinfo("成功", "資料已成功存入資料庫")
    elif mode == '輸出推薦委員':
        search_v3(is_industry)
        statistic_committee()  # 統計清單人才資料_RDF
        filter_committee(is_industry)  # 篩選人員
        excel_process_VBA()
        messagebox.showinfo("成功", "已成功輸出推薦委員")

def create_gui():
    
    current_plan = value_of_key("目前執行計畫")
    
    # 建立主視窗
    window = tk.Tk()
    window.title(f"腳本執行器")
    window.geometry("300x200")

    # 標題標籤
    label = tk.Label(window, text=f"請選擇模式 (目前計畫「{current_plan}」)", font=("Arial", 14))
    label.pack(pady=20)

    # 按鈕 - 存入資料庫
    button_db = tk.Button(window, text="存入資料庫(請確保記憶體足夠)", command=lambda: execute_mode('存入資料庫', current_plan))
    button_db.pack(pady=10)

    # 按鈕 - 輸出推薦委員
    button_committee = tk.Button(window, text="輸出推薦委員", command=lambda: execute_mode('輸出推薦委員', current_plan))
    button_committee.pack(pady=10)

    # 啟動主視窗循環
    window.mainloop()

if __name__ == "__main__":
    create_gui()
