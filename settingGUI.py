from ruamel.yaml import YAML
from tkinter import Tk, filedialog, messagebox, StringVar, Label, Button, Toplevel
from tkinter.ttk import Combobox
import os

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
    
    # 建立主 Tk 視窗
    root = Tk()
    root.title("計畫目標選擇")
    root.geometry("300x150")
    
    valid_project_aim_options = ["產學合作", "研究計畫"]

    # 設定選擇變數
    selected_aim = StringVar()
    selected_aim.set(valid_project_aim_options[0])  # 預設值

    # 標籤和 Combobox
    label = Label(root, text="請選擇目前執行計畫:")
    label.pack(pady=10)

    # 使用 Combobox
    combobox = Combobox(root, textvariable=selected_aim, values=valid_project_aim_options, state="readonly")
    combobox.pack(pady=10)

    def on_select():
        current_project_aim = selected_aim.get()
        setting_data['SOURCE']['field']['目前執行計畫'] = current_project_aim
        messagebox.showinfo("成功", f"目前執行計劃目標已經設定為: {current_project_aim}")
        root.quit()  # 結束主視窗的事件循環

    # 確認按鈕
    button = Button(root, text="確認", command=on_select)
    button.pack(pady=10)
    
    root.mainloop()  # 啟動主事件循環，等待使用者選擇

def select_the_file_update_project_name(setting_data):
    """
    使用 GUI 讓使用者選擇一個檔案，並根據目前執行計畫更新對應的檔案名稱和副檔名。
    限制檔案選擇在 ./data 資料夾下的任何子資料夾。
    :param setting_data: 讀取的設定資料字典
    """
    # 建立主 Tk 視窗
    root = Tk()
    root.title("計畫目標選擇")
    root.geometry("300x150")
    
    current_aim = setting_data['SOURCE']['field']['目前執行計畫']

    # 定義目錄路徑，允許在 ./data/ 資料夾下的任何子資料夾
    if current_aim == "研究計畫": base_directory = os.path.abspath("./data/research_proj/")
    elif current_aim == "產學合作": base_directory = os.path.abspath("./data/industry_coop/")

    # 建立檔案選擇視窗
    file_window = Toplevel(root)
    file_window.title("選擇檔案")
    file_window.geometry("400x200")

    def open_file_dialog():
        # 限制在 ./data/ 資料夾下選擇檔案
        file_path = filedialog.askopenfilename(initialdir=base_directory, title=f"選擇 {current_aim} 的初始資料檔案")
        if file_path:
            # 將選擇的檔案路徑標準化並檢查是否在允許的目錄內
            file_path = os.path.abspath(file_path)
            if file_path.startswith(base_directory):
                file_name = os.path.basename(file_path)  # 獲取檔案名稱和副檔名
                if current_aim == '研究計畫':
                    setting_data['SOURCE']['data']['research_proj']['研究計畫申請名冊'] = file_name
                elif current_aim == '產學合作':
                    setting_data['SOURCE']['data']['industry_coop']['產學合作申請名冊'] = file_name
                messagebox.showinfo("成功", f"已更新檔案名稱為: {file_name}")
                file_window.quit()  # 結束檔案選擇視窗的事件循環
            else:
                messagebox.showerror("錯誤", "選擇的檔案不在允許的目錄內。")
        else:
            messagebox.showerror("錯誤", "未選擇任何檔案。")

    # 標籤說明
    label = Label(file_window, text=f"請選擇 {current_aim} 的初始資料檔案")
    label.pack(pady=20)

    # 檔案選擇按鈕
    select_button = Button(file_window, text="選擇檔案", command=open_file_dialog)
    select_button.pack(pady=20)

    root.withdraw()  # 隱藏主視窗
    file_window.mainloop()  # 啟動檔案選擇視窗的事件循環

try:
    # 讀取現有設定資料
    with open(DEFAULT_SETTING_YAML, 'r', encoding='utf-8') as file:
        setting_data = yaml.load(file)

    # 呼叫選擇計畫目標函數
    select_and_update_project_aim(setting_data)

    # 呼叫選擇檔案並更新計畫名稱函數
    select_the_file_update_project_name(setting_data)

    # 將更新後的設定資料存回檔案
    with open(DEFAULT_SETTING_YAML, 'w', encoding='utf-8') as file:
        yaml.dump(setting_data, file)

    messagebox.showinfo("完成", "設定已成功更新並存回檔案。")

except Exception as e:
    messagebox.showerror("錯誤", f"發生錯誤，無法更新設定檔。\n{e}")
finally:
    pass
