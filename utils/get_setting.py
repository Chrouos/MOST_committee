import yaml
import os

# - 打開 YAML 檔案並讀取數據
setting_data = None
with open('./setting.yaml', 'r', encoding='utf-8') as file:
    setting_data = yaml.safe_load(file)

# - 找到路徑過程
def find_key_path_list(dictionary, target_key, path=''):
    for key, value in dictionary.items():
        
        if key == target_key:
            temp_path = path + str(value)
            result_list = temp_path.split(" -> ")
            return result_list
        
        elif isinstance(value, dict):
            result = find_key_path_list(value, target_key, path + key + ' -> ')
            if result: return result
        
    return []

# - 輸出路徑
def find_key_path(target_key):
    result_text = "."
    path_list = find_key_path_list(setting_data, target_key)
    for node in path_list[1:]:
        result_text = result_text + "/" + node
        
    directory = os.path.dirname(result_text)
    if not os.path.exists(directory):
        os.makedirs(directory)
        
    return result_text

# - 直接找到資料 value
def value_of_key(target_key, dictionary=setting_data):
    
    for key, value in dictionary.items():
        if key == target_key: return value
        elif isinstance(value, dict):
            result = value_of_key(target_key, value)
            if result: return result
        
    return ""

# - 輸出設定數據
def print_setting_data():
    
    print("v ================================ DATA SETTING ================================")
    def print_dict(d, indent=0):
        for key, value in d.items():
            print('  ' * indent + str(key))
            if isinstance(value, dict):
                print_dict(value, indent+1)
            else:
                print('  ' * (indent+1) + str(value))
    
    # 使用遞迴函數打印 settings 字典
    print_dict(setting_data)
    print("^ ================================ DATA SETTING ================================")
    
    
    
if __name__ == '__main__':
    print_setting_data()