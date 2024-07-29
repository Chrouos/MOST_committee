import pandas as pd
import re

# 處理機關名稱分成學校和系所
def split_institution(department_full):
    keywords = ['大學', '院', '博物館', '學校', '法人']  # 列出所有可能的分割關鍵字
    for keyword in keywords:
        if keyword in department_full:
            school, department = department_full.split(keyword, 1)
            school += keyword  # 將關鍵字加回學校名稱中
            return pd.Series([school.strip(), department.strip()])
    return pd.Series([department_full.strip(), ''])  # 如果沒有關鍵字，就只有學校沒有系所

# 找到括號的位置
def extract_text_in_parentheses(text):
    if text: return [['', '']]
    
    pattern = r'([^;]+)\(([^)]+)\)'  # 捕捉名字和括號內的內容
    matches = re.findall(pattern, text)
    
    details = [[match[0].strip(), match[1]] for match in matches]
    
    return details
    
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
    
    # print("schools_info:\n", schools_info)
    # print("committee_members:\n", committee_members)
    # print("filter_pairs:\n", filter_pairs)
    # print()
    
    filtered_members = set()
    filter_reasons = {}

    # 根據配對條件進行過濾
    for school_type, member_field in filter_pairs:
        if school_type in schools_info and schools_info[school_type]:
            for member in committee_members:
                # 檢查是否有匹配的學校導致過濾
                matching_schools = [school for school in member[member_field] if school == schools_info[school_type]]
                if matching_schools:
                    filtered_members.add(member['委員名稱'])
                    filter_reasons[member['委員名稱']] = f"{school_type} ({schools_info[school_type]}) 與 {member_field} ({', '.join(matching_schools)}) 重疊"

    # 創建過濾後的委員名單
    remaining_members = [member['委員名稱'] for member in committee_members if member['委員名稱'] not in filtered_members]

    # 返回結果
    return {
        'Filtered Members': list(filtered_members),
        'Remaining Members': remaining_members,
        'Filter Reasons': filter_reasons
    }