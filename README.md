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

小筆記：：

啟動虛擬環境：

Windows：
windowsenv\Scripts\activate

macOS/Linux：
source macenv/bin/activate

更新requirement.txt
pip install -r requirement.txt

問題1:
根據錯誤訊息，你在安裝 chroma-hnswlib 時遇到以下問題：
error: Microsoft Visual C++ 14.0 or greater is required. Get it with "Microsoft C++ Build Tools": 

解決：要裝
https://visualstudio.microsoft.com/visual-cpp-build-tools/
=> 勾選 Desktop development with C++

問題2:
根據錯誤訊息，問題出在 uvloop 這個套件不支援 Windows，因此在安裝過程中遇到了錯誤：
RuntimeError: uvloop does not support Windows at the moment

解決：uvloop==0.20.0; sys_platform != 'win32' (略過)
