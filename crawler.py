import pandas as pd
import random
import time
from utils.package import generate_search_query, surf, reload_cookies, re_school_department, re_post_request
from bs4 import BeautifulSoup

from utils.get_setting import find_key_path

# : 全局變數設定
URL = 'https://ndltd.ncl.edu.tw/' # = 碩博士論文網
COOKIE = "oMevOx" # = 預設 cookie

FILE_PATH = find_key_path("查找碩博士名單")
OUTPUT_FILE = find_key_path("碩博士論文")
OUTPUT_FILE_RDF = find_key_path("碩博士論文_RDF")

# = 預設 Excel 呈現樣式
clear_excel_dict_template = {   
    '計畫主持人': '',
    '學校': '',
    "碩士畢業學年度": "", 
    "碩士畢業學校": "", 
    "碩士指導教授":"",    
    "碩士論文題目": "",
    "博士畢業學年度": "",
    "博士畢業學校": "",
    "博士指導教授": "",
    "博士論文題目": ""
} 

def save_to_excel(excel_df, output_path):
    """
    將 DataFrame 儲存為 Excel 檔案。
    """
    excel_df.to_excel(output_path, index=False, engine='openpyxl', sheet_name="研究人才")
    
def crawl_thesis_info(row_data):
    """
    爬取指定「名稱」的論文資訊。
    """
    
    student_name = row_data['計畫主持人']
    school_name = row_data['學校']
        
    # : 準備 Excel 存擋資料
    temp_dict = clear_excel_dict_template.copy()
    temp_dict['計畫主持人'] = student_name
    temp_dict['學校'] = school_name
    temp_dict["備註"] = ""
    
    query = generate_search_query(student_name=student_name) # = 爬蟲的搜尋關鍵字
    retry = True # = 是否重複爬取
    error = 0    # = 紀控錯誤次數
    
    while retry and error <= 2:
        try:
            cookie, rs, res_post, h1 = re_post_request(cookie, query, headers, h1) # ! re_post
            soup = BeautifulSoup(res_post.text, 'html.parser')
        
            # - 找到共搜索幾筆資料
            brwrestable = soup.find('table', {'class': 'brwrestable'})
            if not brwrestable:
                return Exception("Can't find brwrestable...")
            
            brwreSpan = brwrestable.findAll("span", {"class": "etd_e"})
            search_counts = int([brwres.text.replace('\xa0', '') for brwres in brwreSpan if brwres.text != query][0])
            temp_dict["查獲人數"] = f"{search_counts}"
            
            # - 找到內文
            PHD_count = 0
            temp_dict["查獲博士人數"] = f"{PHD_count}"
            if search_counts <= 10: 
                
                if search_counts > 2: 
                    temp_dict["備註"] += f"人數多於2人以上，跳過碩士學位"
                
                for r1 in range(1, int(search_counts) + 1):
                    
                    current_data_dict = {} # = 找到內文存成字典
                    
                    # - 爬蟲資料
                    res_get = surf(cookie, rs, r1, h1) # ! re_get
                    soup_content = BeautifulSoup(res_get.text, 'html.parser')

                    # - 內文找到格式
                    contentTable = soup_content.find('table', {'id': 'format0_disparea'})
                    if contentTable:
                        contents = [(content.find('th',{'class':'std1'}).text.replace(":", ""), content.find('td',{'class':'std2'}).text) 
                                    for content in contentTable.findAll('tr') 
                                    if not content.find("td", {"class": "push_td"}) and not content.find("img", {"alt": "被引用"})]

                        for column, value in contents:
                            current_data_dict[column] = value
                            
                            
                    # - 處理後綴
                    endWith = ""
                    if current_data_dict["學位類別"] == "博士":
                        PHD_count += 1
                        if PHD_count > 1:
                            endWith = f"_{PHD_count}"
                            
                        temp_dict["查獲博士人數"] = f"{PHD_count}"
                            
                    if (current_data_dict["學位類別"] == "碩士" and search_counts <= 2) or (current_data_dict["學位類別"] == "博士"):
                        # - 填入
                        temp_dict[f"{current_data_dict['學位類別']}畢業學年度" + endWith] = current_data_dict["畢業學年度"]
                        temp_dict[f"{current_data_dict['學位類別']}畢業學校" + endWith] = current_data_dict["校院名稱"] + "／" + current_data_dict["系所名稱"]
                        temp_dict[f"{current_data_dict['學位類別']}指導教授" + endWith] = current_data_dict["指導教授"]
                        if "論文名稱" in current_data_dict:
                            temp_dict[f"{current_data_dict['學位類別']}論文題目" + endWith] = current_data_dict["論文名稱"]
                        elif "論文名稱(外文)" in current_data_dict:
                            temp_dict[f"{current_data_dict['學位類別']}論文題目" + endWith] = current_data_dict["論文名稱(外文)"]
                            
            else:
                print("人數過多，僅搜尋十筆內資料 => 跳過", end="")
                temp_dict["備註"] += f"人數過多，僅搜尋十筆內資料"
                    
            # - 結束
            rs.close()
            retry = False
            time.sleep(random.randint(2, 5))
            
        except Exception as e:
            cookie, rs, res_post, headers, h1 = reload_cookies(URL, query) # ! reload cookies
            retry = True
            error += 1
            
    return temp_dict

def to_RDF(df):
    result_rdf_df = pd.DataFrame()
    
    for index, row in df.iterrows():
        
        # 檢查碩士資料是否不為空或NaN
        if pd.notnull(row["碩士畢業學年度"]) or pd.notnull(row["碩士畢業學校"]) or pd.notnull(row["碩士論文題目"]):
            current_dict = {
                "學生姓名": row["計畫主持人"],
                "畢業學年度": int(row[f"碩士畢業學年度"]),
                "畢業學校": row["碩士畢業學校"],
                "論文題目": row["碩士論文題目"],
                "學籍": "碩士",
                "計畫發表學校": row["學校"]
            }
            result_rdf_df = pd.concat([result_rdf_df, pd.DataFrame([current_dict])], ignore_index=True)
        
        # 處理博士資料
        for PHD_count in range(int(row["查獲博士人數"])):
            
            endWith = f"_{PHD_count}" if PHD_count > 1 else ""
            
            current_dict = {
                "學生姓名": row["計畫主持人"],
                "畢業學年度": int(row[f"博士畢業學年度{endWith}"]),
                "畢業學校": row[f"博士畢業學校{endWith}"],
                "論文題目": row[f"博士論文題目{endWith}"],
                "學籍": "博士",
                "計畫發表學校": row["學校"]
            }
            result_rdf_df = pd.concat([result_rdf_df, pd.DataFrame([current_dict])], ignore_index=True)


    # 處理重複
    columns_to_check = ["學生姓名", "畢業學年度", "畢業學校", "論文題目", "學籍"]
    result_rdf_df = result_rdf_df.drop_duplicates(subset=columns_to_check, keep="first")
    return result_rdf_df

def main():
    columns = list(clear_excel_dict_template.keys())
    excel_data = pd.read_excel(FILE_PATH, sheet_name="研究人才") # = 讀取的 Excel Data
    result_crawler_df = pd.DataFrame(columns=columns) # = 要存檔的 Excel Data
    
    # - 爬蟲
    for index, row in excel_data.iterrows():
        temp_dict = crawl_thesis_info(row) # = 爬蟲 Action
        temp_df = pd.DataFrame([temp_dict])
        result_crawler_df = pd.concat([result_crawler_df, temp_df], ignore_index=True)
    save_to_excel(result_crawler_df, OUTPUT_FILE)
    
    # - 轉換 RDF
    result_red_df = to_RDF(result_crawler_df)
    save_to_excel(result_red_df, OUTPUT_FILE_RDF)
    
if __name__ == "__main__":
    main()