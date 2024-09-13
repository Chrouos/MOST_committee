## How to use (huai)
1. Preprare the data and write the `setting.yaml` to waiting for start the process. 
2. run crawler.pt to spider the data.
3. run main.py to committee

+ Documentation: [Slides](https://docs.google.com/presentation/d/1CEhxtZ017oe7CfgC6S8-L7yWSMi0eQIIFs5ens1QY4I/edit?usp=sharing)
+ `git clone https://github.com/Chrouos/MOST_committee.git`

## Strategy 

過濾審查委員的策略，將下列重疊部分則一律篩選：
1. [計畫申請的學校, 共同主持計畫的學校, 計畫主持人過去的畢業學校, 共同主持人過去的畢業學校]
2. [審查委員的學校, 審查委員過去的畢業學校（碩士+博士）]


# Rulter
+ 「統計清單」中的 Sheet Name 的名稱固定是 「年分+總計畫清單」，例如: 112總計畫清單。
+ 「計畫過去申請案件」
    + Sheet Name 請以 「年分」 取名。
    + 欄位必須是：計畫主持人、計畫編號、機關名稱、計畫類別、計畫中文名稱、學門、職稱


# 使用
安裝 GIT: https://git-scm.com/downloads


1. 安裝環境
2. 執行程式碼

## 創建環境
如果沒有 Conda.
1. 下載 Windows 版本 miniconda (下方Latest Miniconda installer links)
    + 網址: https://docs.anaconda.com/miniconda/

2. 執行 miniconda
    + 步驟全部按next

3. 將 miniconda 安裝路徑 加入使用者環境變數(或使用 Anaconda Prompt)
    + D:\path\to\miniconda3
    + D:\path\to\miniconda3\Scripts
    + D:\path\to\miniconda3\Library\bin

安裝 Python 環境：
```
conda update conda
conda create --name MOST python=3.11.5
conda init

# 重新啟動 terminal
conda activate MOST

pip install -r requirement.txt
```

## 執行
```
# 研究計畫
python main.py --is_industry False --is_load_chroma_bge False

# 產學合作
python main.py --is_industry True --is_load_chroma_bge False

#!!! 如果已經存入過資料庫了(預設) 若有需要重新匯入資料 is_load_chroma_bge 需要是 True

# 爬蟲碩博士論文網
python crawler.py
```

## 更新
```
git pull
```
