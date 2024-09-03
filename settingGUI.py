from ruamel.yaml import YAML
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, Toplevel
from tkinter.ttk import Combobox
import os
from openpyxl import load_workbook

# 定義設定檔路徑
DEFAULT_SETTING_YAML = "./setting.yaml"

# 初始化 YAML 物件
yaml = YAML()
yaml.preserve_quotes = True

def select_and_update_project_aim(setting_data):
    """
    使用 GUI 讓使用者選擇目前執行計畫 ('產學合作' 或 '研究計畫') 並更新設定資料。
    :param setting_data: 讀取的設定資料字典
    """
    root = Tk()
    root.title("計畫目標選擇")
    root.geometry("300x150")
    
    valid_project_aim_options = ["產學合作", "研究計畫"]

    selected_aim = StringVar()
    selected_aim.set(valid_project_aim_options[0])  # 預設值

    label = Label(root, text="請選擇目前執行計畫:")
    label.pack(pady=10)

    combobox = Combobox(root, textvariable=selected_aim, values=valid_project_aim_options, state="readonly")
    combobox.pack(pady=10)

    def on_select():
        current_project_aim = selected_aim.get()
        setting_data['SOURCE']['field']['目前執行計畫'] = current_project_aim
        messagebox.showinfo("成功", f"目前執行計劃目標已經設定為: {current_project_aim}")
        root.destroy()  # 正確關閉視窗

    button = Button(root, text="確認", command=on_select)
    button.pack(pady=10)
    
    root.mainloop()

def select_the_file_update_project_name(setting_data):
    """
    使用 GUI 讓使用者選擇一個檔案，並根據目前執行計畫更新對應的檔案名稱和副檔名。
    限制檔案選擇在 ./data 資料夾下的任何子資料夾。
    :param setting_data: 讀取的設定資料字典
    """
    root = Tk()
    root.title("選擇資料名稱")
    root.geometry("300x150")
    
    current_aim = setting_data['SOURCE']['field']['目前執行計畫']

    if current_aim == "研究計畫": 
        base_directory = os.path.abspath("./data/research_proj/")
    elif current_aim == "產學合作": 
        base_directory = os.path.abspath("./data/industry_coop/")

    file_window = Toplevel(root)
    file_window.title("選擇檔案")
    file_window.geometry("400x200")

    def open_file_dialog():
        file_path = filedialog.askopenfilename(initialdir=base_directory, title=f"選擇 {current_aim} 的初始資料檔案")
        if file_path:
            file_path = os.path.abspath(file_path)
            if file_path.startswith(base_directory):
                file_name = os.path.basename(file_path)
                if current_aim == '研究計畫':
                    setting_data['SOURCE']['data']['research_proj']['研究計畫申請名冊'] = file_name
                elif current_aim == '產學合作':
                    setting_data['SOURCE']['data']['industry_coop']['產學合作申請名冊'] = file_name
                messagebox.showinfo("成功", f"已更新檔案名稱為: {file_name}")
                file_window.destroy()  # 正確關閉視窗
                root.destroy()  # 關閉主視窗
                # 呼叫下一步驟的函數來選擇 Sheet
                select_sheet_from_excel(file_path, setting_data)
            else:
                messagebox.showerror("錯誤", "選擇的檔案不在允許的目錄內。")
        else:
            messagebox.showerror("錯誤", "未選擇任何檔案。")

    label = Label(file_window, text=f"請選擇 {current_aim} 的初始資料檔案")
    label.pack(pady=20)

    select_button = Button(file_window, text="選擇檔案", command=open_file_dialog)
    select_button.pack(pady=20)

    root.withdraw()
    file_window.mainloop()

def select_sheet_from_excel(file_path, setting_data):
    """
    讀取 Excel 檔案中的 Sheet，並讓使用者選擇要使用的 Sheet 名稱。
    更新設定資料的 '計畫SHEET' 欄位。
    :param file_path: Excel 檔案的路徑
    :param setting_data: 讀取的設定資料字典
    """
    root = Tk()
    root.title("選擇計畫 SHEET")
    root.geometry("300x150")

    # 讀取 Excel 檔案中的 Sheets
    workbook = load_workbook(file_path, read_only=True)
    sheets = workbook.sheetnames

    selected_sheet = StringVar()
    selected_sheet.set(sheets[0])  # 預設選擇第一個 sheet

    label = Label(root, text="請選擇計畫 SHEET:")
    label.pack(pady=10)

    combobox = Combobox(root, textvariable=selected_sheet, values=sheets, state="readonly")
    combobox.pack(pady=10)

    def on_select():
        current_sheet = selected_sheet.get()
        setting_data['SOURCE']['field']['計畫SHEET'] = current_sheet
        messagebox.showinfo("成功", f"計畫 SHEET 已更新為: {current_sheet}")
        root.destroy()  # 正確關閉視窗

    button = Button(root, text="確認", command=on_select)
    button.pack(pady=10)
    
    root.mainloop()

try:
    with open(DEFAULT_SETTING_YAML, 'r', encoding='utf-8') as file:
        setting_data = yaml.load(file)

    select_and_update_project_aim(setting_data)
    select_the_file_update_project_name(setting_data)

    with open(DEFAULT_SETTING_YAML, 'w', encoding='utf-8') as file:
        yaml.dump(setting_data, file)

    messagebox.showinfo("完成", "設定已成功更新並存回檔案。")

except Exception as e:
    messagebox.showerror("錯誤", f"發生錯誤，無法更新設定檔。\n{e}")
finally:
    pass
