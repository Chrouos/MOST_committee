from ruamel.yaml import YAML

# 定義設定檔路徑
DEFAULT_SETTING_YAML = "./setting.yaml"

# 初始化 YAML 物件
yaml = YAML()
yaml.preserve_quotes = True

def select_and_update_project_aim(setting_data):
    """
    讓使用者選擇目前執行計畫 ('產學合作' 或 '研究計畫') 並更新設定資料。
    :param setting_data: 讀取的設定資料字典
    """
    valid_project_aim_options = ["產學合作", "研究計畫"]

    # 反覆詢問直到輸入正確
    while True:
        current_project_aim = input("Please select the project aim ('產學合作'|'研究計畫'): ")
        if current_project_aim in valid_project_aim_options:
            # 更新設定資料
            setting_data['SOURCE']['field']['目前執行計畫'] = current_project_aim
            print(f"目前執行計劃目標已經設定為: {current_project_aim}")
            break
        else:
            print("輸入錯誤，請輸入 '產學合作' 或 '研究計畫'。")
            
def select_the_file_update_project_name(setting_data):
    if setting_data['SOURCE']['field']['目前執行計畫'] == '研究計畫':
        setting_data['SOURCE']['field']['研究計畫申請名冊'] = ...
        
    elif setting_data['SOURCE']['field']['目前執行計畫'] == '產學合作':
        setting_data['SOURCE']['field']['產學合作申請名冊'] = ...

# 讀取現有設定資料
with open(DEFAULT_SETTING_YAML, 'r', encoding='utf-8') as file:
    setting_data = yaml.load(file)

# - 改寫中
try:
    select_and_update_project_aim(setting_data)

    with open(DEFAULT_SETTING_YAML, 'w', encoding='utf-8') as file:
        yaml.dump(setting_data, file)

    print("設定已成功更新並存回檔案。")
except Exception as e:
    print("設定檔格式錯誤，無法找到正確的結構。")
    print(e)
