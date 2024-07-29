import tkinter as tk
from tkinter import messagebox
from utils.get_setting import setting_data, print_setting_data, find_key_path
from utils.script import statistic_committee, load_into_chroma_bge_manager, search_v3, filter_committee, excel_process_VBA

def run_script(print_setting, is_industry, is_load_chroma_bge):
    if print_setting:
        print_setting_data()
    if is_load_chroma_bge:
        load_into_chroma_bge_manager(is_industry)
    search_v3(is_industry)
    statistic_committee()
    filter_committee()
    excel_process_VBA()

def start_script():
    try:
        print_setting = print_setting_var.get()
        is_industry = is_industry_var.get()
        is_load_chroma_bge = is_load_chroma_bge_var.get()
        run_script(print_setting, is_industry, is_load_chroma_bge)
        messagebox.showinfo("成功", "腳本執行完成")
    except Exception as e:
        messagebox.showerror("錯誤", str(e))

# 建立主視窗
root = tk.Tk()
root.title("腳本執行器")
root.geometry("500x400")  # 設置窗口大小

# 設置標題
tk.Label(root, text="腳本執行器", font=("Helvetica", 20)).pack(pady=20)

# 建立變數
print_setting_var = tk.BooleanVar(value=False)
is_industry_var = tk.BooleanVar(value=True)
is_load_chroma_bge_var = tk.BooleanVar(value=True)

# 建立 GUI 元件框架
frame = tk.Frame(root)
frame.pack(pady=20)

tk.Checkbutton(frame, text="打印設定數據", variable=print_setting_var, font=("Helvetica", 14)).grid(row=0, column=0, sticky='w', pady=10, padx=20)
tk.Checkbutton(frame, text="產業專案 (是/否)", variable=is_industry_var, font=("Helvetica", 14)).grid(row=1, column=0, sticky='w', pady=10, padx=20)
tk.Checkbutton(frame, text="匯入資料庫 (是/否)", variable=is_load_chroma_bge_var, font=("Helvetica", 14)).grid(row=2, column=0, sticky='w', pady=10, padx=20)

tk.Button(frame, text="執行", command=start_script, font=("Helvetica", 14), width=20, height=2).grid(row=3, column=0, pady=30)

# 啟動主循環
root.mainloop()
