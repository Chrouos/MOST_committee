import pandas as pd
from typing import List
from utils.get_setting import setting_data, print_setting_data, find_key_path, value_of_key

def get_project_df() -> List[pd.DataFrame]:
    
    # : Field
    apply_project_folder_path = find_key_path("計畫過去申請案件")
    apply_project_file_year = value_of_key("計畫過去申請案件年分範圍")
    statistic_folder_path = find_key_path("統計清單")
    
    # : Variable
    apply_project_excel_file = pd.ExcelFile(apply_project_folder_path)
    
    result_df = {}
    for year in apply_project_file_year:
        
        # - Load the Apply_Project file.
        apply_project_df = pd.read_excel(apply_project_excel_file, year)
        apply_project_df.columns = apply_project_df.iloc[0]
        apply_project_df = apply_project_df.iloc[1:]
        
        # - Load the 
        statistic_excel_file = pd.ExcelFile(statistic_folder_path)
        statistic_df = pd.read_excel(statistic_excel_file, f"{year}總計畫清單")
        pass_proj_name_list = statistic_df['計畫中文名稱'].to_list()
        
        # - Combine the Apply_Project and Admission_List(Statistic)
        pass_list = []
        for i in range(len(apply_project_df)):
            if apply_project_df.iloc[i]['計畫中文名稱'] in pass_proj_name_list:
                pass_list.append('true')
            else:
                pass_list.append('false')
                
        apply_project_df['通過'] = pass_list  
        result_df[year] = apply_project_df
    
    return result_df

def get_industry_coop_proj():
    # 產學計劃
    industry_folder_path = find_key_path("產學過去申請名冊")
    print("industry_folder_path", industry_folder_path)
    xls = pd.ExcelFile(industry_folder_path)
    
    year = ['專題計畫綜合查詢']
    df_list = {}
    for y in year:
        df1 = pd.read_excel(xls, y)
        df_list[y] = df1

    return df_list
