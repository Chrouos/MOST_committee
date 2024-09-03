import argparse
from utils.get_setting import value_of_key
from utils.script import load_into_chroma_bge_manager, search_v3, filter_committee, excel_process_VBA, statistic_committee

def main():
    
    # : parser setting
    parser = argparse.ArgumentParser(description="腳本執行器")
    parser.add_argument('--choose_mode', default='存入資料庫', type=str, help='選擇模式（"存入資料庫", "輸出推薦委員"）', choices=['存入資料庫', '輸出推薦委員'])
    args = parser.parse_args()
    
    if value_of_key("目前執行計畫") == "產學合作": is_industry = True
    elif value_of_key("目前執行計畫") == "研究計畫": is_industry = False
    else: 
        print("「目前執行計畫」沒有填寫正確（產學合作｜研究計畫）")
        return
    
    # 匯入資料庫
    if args.choose_mode == '存入資料庫':    
        load_into_chroma_bge_manager(is_industry)
    elif args.choose_mode == '輸出推薦委員':
        search_v3(is_industry) 
        statistic_committee() #= output: 統計清單人才資料_RDF
        filter_committee(is_industry) #= 篩選人員
        excel_process_VBA()

if __name__ == "__main__":
    
    main()
    