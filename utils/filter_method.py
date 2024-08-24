import pandas as pd
import re

# 處理機關名稱分成學校和系所
def split_institution(department_full):
    if not isinstance(department_full, str) or department_full.strip() == "":
        return pd.Series(['', ''])
    
    keywords = ['大學', '院', '博物館', '學校', '法人']  # 列出所有可能的分割關鍵字
    for keyword in keywords:
        if keyword in department_full:
            school, department = department_full.split(keyword, 1)
            school += keyword  # 將關鍵字加回學校名稱中
            return pd.Series([school.strip(), department.strip()])
    
    return pd.Series([department_full.strip(), ''])  # 如果沒有關鍵字，就只有學校沒有系所

# 找到括號的位置
def extract_text_in_parentheses(text):
    if isinstance(text, str):
        if text == "":
            return [['', '']]
        
        # 捕捉名字和括號內的內容
        pattern = r'([^;]+)\(([^)]+)\)' 
        matches = re.findall(pattern, text)
        
        if matches:  details = [[match[0].strip(), match[1]] for match in matches] # 成對的
        else:
            # 如果沒有匹配到括號中的內容，檢查是否是部門名稱
            if "大學" in text or "學系" in text or "研究所" in text:
                details = [["", text.strip()]]
            else:
                details = [[text.strip(), '']]
        
        return details
    
    elif isinstance(text, list):
        result = []
        for item in text:
            result.extend(extract_text_in_parentheses(item))
        return result

    return []
    
# 找到碩博士論文網中學生的名
def find_crawler_person_relative_school(person, crawler_RDF_data):
    person_data = crawler_RDF_data[crawler_RDF_data['學生姓名'] == person]
    
    if len(person_data) == 0:
        return []
    else:
        result_list = []
        for department in person_data['畢業學校']:
            result_list.append(department.split("／")[0])
            
        return list(set(result_list))

# 取得 dict 所有的 value (unique)
def dict_value_to_list(dict_list, key):
    unique_schools = set()
    for item in dict_list:
        for school in item[key]:
            unique_schools.add(school)

    unique_schools_list = list(unique_schools)
    return unique_schools_list

def filter_committee_person_by_school(apply_school, temp_list):
    print(f"{apply_school} \n=>{temp_list}")
    apply_school_set = set(apply_school) 
    filter = []

    for row in temp_list:
        if not apply_school_set & set(row["相關學校"]):  
            filter.append(row)
            
    return filter

def filter_committee_advanced(schools_info, committee_members, filter_pairs):
    """
    進階過濾委員名單，根據具體的配對關係進行過濾，並提供過濾的具體原因。

    :param schools_info: 包含學校相關資訊的字典
    :param committee_members: 包含委員相關資訊的列表
    :param filter_pairs: 列表，包含過濾配對條件，例如 [("申請學校", "就職學校")]
    :return: 一個字典，包含過濾前後的委員名單和未過濾的委員名單，以及過濾原因
    """
    
    filtered_members = set()
    filter_reasons = {}

    # 根據配對條件進行過濾
    for school_type, member_field in filter_pairs:
        if school_type in schools_info and schools_info[school_type]:
            school_list = schools_info[school_type] if isinstance(schools_info[school_type], list) else [schools_info[school_type]]
            for member in committee_members:
                matching_schools = [school for school in member[member_field] if school in school_list and school]
                if matching_schools:
                    filtered_members.add(member['委員名稱'])
                    filter_reasons[member['委員名稱']] = f"{school_type} 與 {member_field} ({', '.join(matching_schools)}) 重疊"

    # 創建過濾後的委員名單
    remaining_members = [member['委員名稱'] for member in committee_members if member['委員名稱'] not in filtered_members]

    # 返回結果
    return {
        'Filtered Members': list(filtered_members),
        'Remaining Members': remaining_members,
        'Filter Reasons': filter_reasons
    }
    
def merge_committee_advanced(result1, result2):
    """
    合併兩個過濾結果。

    :param result1: 第一個過濾結果字典
    :param result2: 第二個過濾結果字典
    :return: 合併後的過濾結果字典
    """
    print(result1)
    print(result2)
    
    # 合併 'Filtered Members'
    merged_filtered_members = list(set(result1['Filtered Members'] + result2['Filtered Members']))
    
    # 合併 'Remaining Members'
    merged_remaining_members = list(set(result1['Remaining Members'] + result2['Remaining Members']) - set(merged_filtered_members))
    
    # 合併 'Filter Reasons'
    merged_filter_reasons = {**result1['Filter Reasons'], **result2['Filter Reasons']}
    
    return {
        'Filtered Members': merged_filtered_members,
        'Remaining Members': merged_remaining_members,
        'Filter Reasons': merged_filter_reasons
    }