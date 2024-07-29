import argparse
from utils.get_setting import setting_data, print_setting_data, find_key_path
from utils.script import load_into_chroma_bge_manager, search_v3, filter_committee, excel_process_VBA, statistic_committee

def str2bool(v):
    if isinstance(v, bool):
        return v
    if v.lower() in ('yes', 'true', 't', 'y', '1'):
        return True
    elif v.lower() in ('no', 'false', 'f', 'n', '0'):
        return False
    else:
        raise argparse.ArgumentTypeError('Boolean value expected.')

def main():
    
    # : parser setting
    parser = argparse.ArgumentParser(description="檢查是否要打印設定數據")
    parser.add_argument('--print_setting', action='store_true', help='如果提供這個選項，則打印設定數據')
    parser.add_argument('--is_industry', type=str2bool, default=True, help='(研究計畫＝False, 產業專案＝True)')
    
    args = parser.parse_args()
    
    # : actual choice
    if args.print_setting: print_setting_data()
    is_industry = args.is_industry
    
    # load_into_chroma_bge_manager(is_industry)
    search_v3(is_industry) 
    statistic_committee() #= output: 統計清單人才資料_RDF
    filter_committee() #= 篩選人員
    excel_process_VBA()

if __name__ == "__main__":
    
    main()
    