#!/usr/bin/env python

import git
import argparse
import os

def git_pull(repo_path):
    try:
        repo = git.Repo(repo_path)
        repo.remotes.origin.pull()
        print(f"Git pull 成功 in {repo_path}")
    except Exception as e:
        print(f"Git pull 失敗: {e}")

def main():
    parser = argparse.ArgumentParser(description="模擬 git pull 的功能")
    parser.add_argument("path", help="Git 儲存庫的路徑")
    args = parser.parse_args()

    if not os.path.exists(args.path):
        print("指定的路徑不存在")
        return

    git_pull(args.path)

if __name__ == "__main__":
    main()
