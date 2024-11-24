#!/usr/bin/env python

import git
import os

def git_pull(repo_path):
    """
    執行 git pull 操作
    :param repo_path: Git 儲存庫的路徑
    """
    try:
        # 嘗試打開指定路徑的儲存庫
        repo = git.Repo(repo_path)
        repo.remotes.origin.pull()
        print(f"Git pull 成功 in {repo_path}")
    except Exception as e:
        print(f"Git pull 失敗: {e}")

def main():
    """
    主程式入口，執行當前目錄的 git pull
    """
    # 獲取當前目錄
    current_path = os.getcwd()

    # 檢查是否為有效的 Git 儲存庫
    if not os.path.exists(os.path.join(current_path, ".git")):
        print("當前目錄不是一個有效的 Git 儲存庫")
        return

    # 執行 git pull 操作
    git_pull(current_path)

if __name__ == "__main__":
    main()
